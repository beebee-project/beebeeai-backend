const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const XLSX = require("xlsx");
const User = require("../models/User");
const {
  downloadToBuffer,
  saveJsonObject,
  readJsonObject,
  saveBufferObject,
} = require("../utils/storage");
const { readWorkbookFromBuffer } = require("../utils/workbookReader");
const { getOrBuildAllSheetsData } = require("../utils/sheetPreprocessor");
const { buildDownloadFileName } = require("../utils/downloadFileNameBuilder");
const {
  OUTPUT_TYPES,
  getOutputArtifact,
  getOutputExtension,
  getOutputMimeType,
  getOutputDefaultTitle,
  getOutputVersion,
  inferOutputArtifact,
} = require("../automation/config/outputArtifactConfig");
const { decryptBuffer } = require("../services/encryptedFileService");
const {
  readEncryptedQueryJson,
  saveEncryptedQueryJson,
} = require("../services/encryptedJsonStorageService");
const { assertCanUse, bumpUsage } = require("../services/usageService");
const {
  buildQueryTablesFromWorkbook,
} = require("../automation/queryTableBuilder");
const { parseQueryIntent } = require("../automation/queryIntentParser");
const { executeQueryIntent } = require("../automation/queryExecutor");
const {
  buildSummaryWorkbook,
  buildAutomationTemplateWorkbook,
  workbookToBuffer,
  buildChartSpec,
} = require("../automation/summarySheetBuilder");
const { buildReportSections } = require("../automation/reportSectionBuilder");
const { renderReportPpt } = require("../automation/reportPptRenderer");
const {
  buildNormalizedQueryTables,
} = require("../automation/normalizedQueryTableBuilder");
const {
  executeAnalysisRecipeCandidate,
} = require("../automation/analysisRecipeExecutor");
const {
  executeBusinessTemplate,
} = require("../automation/businessTemplateExecutor");
let candidateGenerationModule = {};

try {
  candidateGenerationModule = require("../automation/candidateGeneration");
} catch (error) {
  console.warn(
    "[candidateGeneration] module load failed. Falling back to deterministic builder:",
    error?.message || error,
  );
}
const {
  buildDeterministicCandidateBundle,
} = require("../automation/candidateGeneration/deterministicCandidateBuilder");
const {
  validateCandidateBundle,
} = require("../automation/candidateGeneration/candidateValidator");

function resolveGenerateCandidateBundle() {
  const direct = candidateGenerationModule;

  const fn =
    (typeof direct === "function" && direct) ||
    (typeof direct?.generateCandidateBundle === "function" &&
      direct.generateCandidateBundle) ||
    (typeof direct?.default === "function" && direct.default);

  if (fn) return fn;

  console.warn(
    "[candidateGeneration] generateCandidateBundle export missing. Using deterministic fallback.",
  );

  return async function generateCandidateBundleFallback({
    normalizedQueryTables = [],
    fileName = "",
    source = "candidate-generation-fallback",
  } = {}) {
    const deterministic = buildDeterministicCandidateBundle({
      normalizedQueryTables,
      fileName,
      source,
    });

    const validated = validateCandidateBundle(
      deterministic,
      normalizedQueryTables,
    );

    return {
      ...validated,
      candidateGeneration: {
        ...(validated.candidateGeneration || {}),
        version:
          validated.candidateGeneration?.version || "candidate_generation_v1",
        source,
        fallbackUsed: true,
        fileName,
        aiReranker: {
          enabled: false,
          used: false,
          skippedReason: "INVALID_CANDIDATE_GENERATION_EXPORT",
        },
        generatedAt: new Date().toISOString(),
      },
    };
  };
}

const generateCandidateBundle = resolveGenerateCandidateBundle();
const {
  isBusinessTemplateResult,
  normalizeBusinessTemplateResult,
  outputTypeLabel,
} = require("../automation/businessTemplateContract");

function getGeneratedLocalDir(outputType) {
  const artifact = getOutputArtifact(outputType);
  const dirName = artifact?.localDirName || String(outputType || "generated");

  return path.join(process.cwd(), ".local_uploads", "generated", dirName);
}

function getGeneratedStoragePrefix(outputType) {
  const artifact = getOutputArtifact(outputType);
  return artifact?.storagePrefix || String(outputType || "generated");
}

const REPORT_DIR = getGeneratedLocalDir(OUTPUT_TYPES.ANALYSIS_REPORT);
const PPT_DIR = getGeneratedLocalDir(OUTPUT_TYPES.PPT);
const AUTOMATION_DIR = getGeneratedLocalDir(OUTPUT_TYPES.SUMMARY_SHEET);
const SUMMARY_SHEET_STORAGE_PREFIX = getGeneratedStoragePrefix(
  OUTPUT_TYPES.SUMMARY_SHEET,
);

async function assertTemplateGenerationUsage(req, res) {
  if (!req.user?.id) return true;

  try {
    await assertCanUse(req.user.id, "templateGenerations", 1);
    return true;
  } catch (e) {
    res.status(e.status || 429).json({
      ok: false,
      error: "Usage limit exceeded",
      code: e.code || "LIMIT_EXCEEDED",
      ...(e.meta || {}),
    });
    return false;
  }
}

async function bumpTemplateGenerationUsage(req) {
  if (!req.user?.id) return;
  await bumpUsage(req.user.id, "templateGenerations", 1);
}

function normalizeExecutedResult(result, templateCandidate = null) {
  if (isBusinessTemplateResult(result)) {
    return normalizeBusinessTemplateResult(result, templateCandidate || {});
  }
  return result;
}

function resultTemplateTitle(
  result = {},
  templateCandidate = null,
  fallback = "보고서",
) {
  return (
    result?.title ||
    templateCandidate?.title ||
    result?.templateId ||
    templateCandidate?.templateId ||
    result?.operation ||
    fallback
  );
}

function buildGeneratedFileName({
  sourceFileName,
  templateTitle,
  outputType,
  extension,
}) {
  return buildDownloadFileName({
    sourceFileName,
    templateTitle,
    outputType,
    extension: extension || getOutputExtension(outputType),
  });
}

function encodeDownloadName(fileName = "download") {
  return encodeURIComponent(String(fileName || "download").trim());
}

function contentTypeForGeneratedFile(fileName = "", outputType = "") {
  const artifact = inferOutputArtifact({ outputType, fileName });
  return artifact?.mimeType || "application/octet-stream";
}

function buildGeneratedDownloadUrl({
  storageKey = "",
  filePath = "",
  displayName = "",
  outputType = "",
} = {}) {
  const params = new URLSearchParams();

  if (storageKey) params.set("storageKey", storageKey);
  if (filePath) params.set("filePath", filePath);
  if (displayName) params.set("displayName", displayName);
  if (outputType) params.set("outputType", outputType);

  return `/api/automation/download?${params.toString()}`;
}

function assertGeneratedStorageKeyAccess(req, storageKey = "") {
  const key = String(storageKey || "");
  if (!key) return false;

  const userId = req.user?.id ? String(req.user.id) : "local-dev";

  return (
    key.startsWith(`${SUMMARY_SHEET_STORAGE_PREFIX}/${userId}/`) ||
    key.startsWith(`${SUMMARY_SHEET_STORAGE_PREFIX}/local-dev/`)
  );
}

function resolveSafeGeneratedLocalPath(filePath = "") {
  if (!filePath) return null;

  const resolved = path.resolve(filePath);
  const allowedDirs = [REPORT_DIR, PPT_DIR, AUTOMATION_DIR].map((dir) =>
    path.resolve(dir),
  );

  const allowed = allowedDirs.some(
    (dir) => resolved === dir || resolved.startsWith(`${dir}${path.sep}`),
  );

  return allowed ? resolved : null;
}

function isLocalDevBypassMode() {
  return (
    process.env.NODE_ENV !== "production" &&
    process.env.LOCAL_DEV === "1" &&
    process.env.DEV_BYPASS_AUTH === "1"
  );
}

function findUserFile(user, fileName) {
  if (!user || !fileName) return null;
  return user.uploadedFiles?.find((f) => f.originalName === fileName) || null;
}

const MOJIBAKE_PATTERN =
  /[¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ¸¼½¾±ºµ]/;

function hasMojibakeQueryPayload(payload = {}) {
  const tables = payload.tables || payload.normalizedQueryTables || [];

  const headers = [];

  for (const table of tables) {
    for (const col of table.columns || []) {
      headers.push(col.header || col.originalHeader || "");
    }
  }

  const text = headers.join(" ");
  const suspiciousCount = (text.match(MOJIBAKE_PATTERN) || []).length;

  return suspiciousCount >= 2;
}

async function loadSavedQueryJsonForFile(req, fileName) {
  // 로컬 회귀 테스트/개발 인증 우회 모드에서는 MongoDB 연결 없이
  // .local_uploads의 원본 파일을 직접 읽는다.
  // 이 가드가 없으면 MONGO_URI 없이 실행한 로컬 서버에서
  // User.findById()가 buffering timeout을 발생시킨다.
  if (isLocalDevBypassMode()) return null;

  if (!req.user?.id || !fileName) return null;

  const user = await User.findById(req.user.id).select("uploadedFiles");
  if (!user) return null;

  const fileInfo = findUserFile(user, fileName);
  if (!fileInfo?.queryJsonKey) {
    return { user, fileInfo, payload: null };
  }

  const payload = await readEncryptedQueryJson(fileInfo.queryJsonKey);

  if (!payload) {
    return { user, fileInfo, payload: null };
  }

  return { user, fileInfo, payload };
}

async function saveQueryJsonForFile({ user, fileInfo, fileName, payload }) {
  if (!user || !fileInfo || !payload) return null;

  const meta = await saveEncryptedQueryJson({
    userId: String(user._id),
    fileName,
    payload,
  });

  if (!meta?.queryJsonKey) return null;

  fileInfo.queryJsonKey = meta.queryJsonKey;
  await user.save();

  return meta.queryJsonKey;
}

async function buildQueryTablesForFile(req, fileName) {
  let buffer;
  const savedQueryJson = await loadSavedQueryJsonForFile(req, fileName);

  if (
    savedQueryJson?.payload &&
    !hasMojibakeQueryPayload(savedQueryJson.payload)
  ) {
    const normalizedQueryTables =
      savedQueryJson.payload.normalizedQueryTables || [];

    let analysisRecipeCandidates =
      savedQueryJson.payload.analysisRecipeCandidates || [];
    let categoryCandidates = savedQueryJson.payload.categoryCandidates || [];
    let businessTemplateCandidates =
      savedQueryJson.payload.businessTemplateCandidates || [];
    let candidateGeneration =
      savedQueryJson.payload.candidateGeneration || null;

    if (
      normalizedQueryTables.length &&
      (!analysisRecipeCandidates.length || !businessTemplateCandidates.length)
    ) {
      const candidateBundle = await generateCandidateBundle({
        normalizedQueryTables,
        fileName,
        source: "encrypted-query-json",
      });

      analysisRecipeCandidates = candidateBundle.analysisRecipeCandidates || [];
      categoryCandidates = candidateBundle.categoryCandidates || [];
      businessTemplateCandidates =
        candidateBundle.businessTemplateCandidates || [];
      candidateGeneration = candidateBundle.candidateGeneration || null;
    }

    let candidateBundle = {
      analysisRecipeCandidates:
        savedQueryJson.payload.analysisRecipeCandidates || [],
      categoryCandidates: savedQueryJson.payload.categoryCandidates || [],
      businessTemplateCandidates:
        savedQueryJson.payload.businessTemplateCandidates || [],
      candidateGeneration: savedQueryJson.payload.candidateGeneration || null,
    };

    // 구버전 query-json에는 candidateGeneration 메타 또는 business 후보가 없을 수 있다.
    // 이 경우에도 AI 실패와 무관하게 deterministic 후보를 다시 만든다.
    if (
      !candidateBundle.candidateGeneration ||
      !candidateBundle.analysisRecipeCandidates.length ||
      !candidateBundle.businessTemplateCandidates.length
    ) {
      candidateBundle = await generateCandidateBundle({
        normalizedQueryTables,
        fileName,
        source: "encrypted-query-json",
      });
    }

    return {
      fileHash: savedQueryJson.payload.fileHash,
      sheetStateSig: savedQueryJson.payload.sheetStateSig,
      tables: savedQueryJson.payload.tables || [],
      normalizedQueryTables,
      analysisRecipeCandidates: candidateBundle.analysisRecipeCandidates || [],
      categoryCandidates: candidateBundle.categoryCandidates || [],
      businessTemplateCandidates:
        candidateBundle.businessTemplateCandidates || [],
      candidateGeneration: candidateBundle.candidateGeneration || null,
      source: "encrypted-query-json",
    };
  }

  if (
    savedQueryJson?.payload &&
    hasMojibakeQueryPayload(savedQueryJson.payload)
  ) {
    console.warn(
      "[query-json] mojibake detected. Rebuilding from source file.",
      {
        fileName,
        fileHash: savedQueryJson.payload.fileHash,
      },
    );
  }

  if (isLocalDevBypassMode()) {
    const path = require("path");
    const fs = require("fs");

    const localPath = path.join(__dirname, "..", ".local_uploads", fileName);

    if (!fs.existsSync(localPath)) {
      const uploadDir = path.join(__dirname, "..", ".local_uploads");
      const availableFiles = fs.existsSync(uploadDir)
        ? fs.readdirSync(uploadDir)
        : [];

      const err = new Error("로컬 테스트 파일을 찾을 수 없습니다.");
      err.status = 404;
      err.payload = { path: localPath, availableFiles };
      throw err;
    }

    buffer = fs.readFileSync(localPath);
  } else {
    const user = await User.findById(req.user.id).select("uploadedFiles");
    if (!user) {
      const err = new Error("사용자 없음");
      err.status = 404;
      throw err;
    }

    const fileInfo = findUserFile(user, fileName);
    if (!fileInfo) {
      const err = new Error("파일을 찾을 수 없습니다.");
      err.status = 404;
      throw err;
    }

    const storageName = fileInfo.localName || fileInfo.gcsName;
    buffer = await downloadToBuffer(storageName);

    if (fileInfo.encrypted) {
      buffer = decryptBuffer(buffer, {
        encryptionIv: fileInfo.encryptionIv,
        encryptionTag: fileInfo.encryptionTag,
      });
    }
  }

  const { fileHash, allSheetsData, sheetStateSig } =
    await getOrBuildAllSheetsData(buffer);

  const workbook = readWorkbookFromBuffer(buffer);
  const tables = buildQueryTablesFromWorkbook(workbook, allSheetsData);

  console.log("[query-table-debug]", {
    fileName,
    tableCount: tables.length,
    tables: tables.map((t) => ({
      tableId: t.tableId,
      rowCount: t.rowCount,
      headerRow: t.headerRow,
      dataStartRow: t.dataStartRow,
      dataEndRow: t.dataEndRow,
      range: t.range,
    })),
  });

  const normalizedQueryTables = buildNormalizedQueryTables(tables);
  const candidateBundle = await generateCandidateBundle({
    normalizedQueryTables,
    fileName,
    source: "rebuilt-from-xlsx",
  });

  const analysisRecipeCandidates =
    candidateBundle.analysisRecipeCandidates || [];
  const categoryCandidates = candidateBundle.categoryCandidates || [];
  const businessTemplateCandidates =
    candidateBundle.businessTemplateCandidates || [];
  const candidateGeneration = candidateBundle.candidateGeneration || null;

  if (savedQueryJson?.user && savedQueryJson?.fileInfo) {
    await saveQueryJsonForFile({
      user: savedQueryJson.user,
      fileInfo: savedQueryJson.fileInfo,
      fileName,
      payload: {
        version: "query_tables_v2",
        fileName,
        fileHash,
        sheetStateSig,
        tableCount: tables.length,
        createdAt: new Date().toISOString(),
        tables,
        normalizedQueryTables,
        analysisRecipeCandidates,
        categoryCandidates,
        businessTemplateCandidates,
        candidateGeneration,
      },
    });
  }

  return {
    fileHash,
    sheetStateSig,
    tables,
    normalizedQueryTables,
    analysisRecipeCandidates,
    categoryCandidates,
    businessTemplateCandidates,
    candidateGeneration,
    source: "rebuilt-from-xlsx",
  };
}

const DEFAULT_ANALYSIS_CANDIDATE_CATEGORY = "automation";

function normalizeAnalysisCandidates(analysisRecipeCandidates = []) {
  return (analysisRecipeCandidates || []).map((candidate, index) => {
    const id =
      candidate.candidateId ||
      candidate.id ||
      candidate.recipeId ||
      candidate.type ||
      `candidate_${index + 1}`;

    return {
      candidateId: id,
      title:
        candidate.title ||
        candidate.name ||
        candidate.label ||
        `자동화 후보 ${index + 1}`,
      description:
        candidate.description ||
        candidate.reason ||
        candidate.summary ||
        "업로드된 파일 구조를 기반으로 생성 가능한 자동화입니다.",
      category:
        candidate.category ||
        candidate.type ||
        candidate.recipeId ||
        DEFAULT_ANALYSIS_CANDIDATE_CATEGORY,
      priority: Number.isFinite(candidate.priority)
        ? candidate.priority
        : index + 1,
      candidate,
    };
  });
}

async function executeAnalysisCandidate(req, res) {
  try {
    const { queryTablesKey, normalizedQueryTables, candidate } = req.body || {};

    let tablesForExecution = normalizedQueryTables;

    if (!Array.isArray(tablesForExecution) && queryTablesKey) {
      const saved = await readJsonObject(queryTablesKey);
      tablesForExecution =
        saved.normalizedQueryTables ||
        buildNormalizedQueryTables(saved.tables || []);
    }

    if (!Array.isArray(tablesForExecution)) {
      return res.status(400).json({
        ok: false,
        code: "NORMALIZED_QUERY_TABLES_REQUIRED",
        message: "normalizedQueryTables 또는 queryTablesKey가 필요합니다.",
      });
    }

    if (!candidate || !candidate.recipeType || !candidate.tableId) {
      return res.status(400).json({
        ok: false,
        code: "ANALYSIS_CANDIDATE_REQUIRED",
        message: "실행할 분석 후보가 필요합니다.",
      });
    }

    const result = executeAnalysisRecipeCandidate({
      normalizedQueryTables: tablesForExecution,
      candidate,
    });

    const status = result.ok ? 200 : 400;
    return res.status(status).json(result);
  } catch (error) {
    console.error("executeAnalysisCandidate error:", error);

    return res.status(500).json({
      ok: false,
      code: "ANALYSIS_CANDIDATE_EXECUTE_FAILED",
      message: "분석 후보 실행 중 오류가 발생했습니다.",
    });
  }
}

async function executeBusinessTemplateCandidate(req, res) {
  try {
    const { queryTablesKey, normalizedQueryTables, templateCandidate } =
      req.body || {};

    let tablesForExecution = normalizedQueryTables;

    if (!Array.isArray(tablesForExecution) && queryTablesKey) {
      const saved = await readJsonObject(queryTablesKey);
      tablesForExecution =
        saved.normalizedQueryTables ||
        buildNormalizedQueryTables(saved.tables || []);
    }

    if (!Array.isArray(tablesForExecution)) {
      return res.status(400).json({
        ok: false,
        code: "NORMALIZED_QUERY_TABLES_REQUIRED",
        message: "normalizedQueryTables 또는 queryTablesKey가 필요합니다.",
      });
    }

    if (!templateCandidate || !templateCandidate.templateId) {
      return res.status(400).json({
        ok: false,
        code: "BUSINESS_TEMPLATE_REQUIRED",
        message: "실행할 업무 템플릿 후보가 필요합니다.",
      });
    }

    const result = executeBusinessTemplate({
      normalizedQueryTables: tablesForExecution,
      templateCandidate,
    });

    const status = result.ok ? 200 : 400;
    return res.status(status).json(result);
  } catch (error) {
    console.error("executeBusinessTemplateCandidate error:", error);

    return res.status(500).json({
      ok: false,
      code: "BUSINESS_TEMPLATE_EXECUTE_FAILED",
      message: "업무 템플릿 실행 중 오류가 발생했습니다.",
    });
  }
}

exports.executeBusinessTemplateCandidate = executeBusinessTemplateCandidate;

exports.executeAnalysisCandidate = executeAnalysisCandidate;

exports.downloadGeneratedFile = async (req, res, next) => {
  try {
    const {
      storageKey = "",
      filePath = "",
      displayName = "",
      outputType = "",
    } = req.query || {};

    const safeDisplayName =
      String(displayName || "").trim() ||
      path.basename(String(storageKey || filePath || "download"));

    let buffer = null;

    if (storageKey) {
      if (!assertGeneratedStorageKeyAccess(req, storageKey)) {
        return res.status(403).json({
          ok: false,
          code: "GENERATED_FILE_FORBIDDEN",
          error: "다운로드 권한이 없습니다.",
        });
      }

      buffer = await downloadToBuffer(storageKey);
    } else if (filePath) {
      const safePath = resolveSafeGeneratedLocalPath(filePath);

      if (!safePath || !fs.existsSync(safePath)) {
        return res.status(404).json({
          ok: false,
          code: "GENERATED_FILE_NOT_FOUND",
          error: "생성 파일을 찾을 수 없습니다.",
        });
      }

      buffer = fs.readFileSync(safePath);
    } else {
      return res.status(400).json({
        ok: false,
        code: "DOWNLOAD_TARGET_REQUIRED",
        error: "storageKey 또는 filePath가 필요합니다.",
      });
    }

    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodeDownloadName(safeDisplayName)}`,
    );
    res.setHeader(
      "Content-Type",
      contentTypeForGeneratedFile(safeDisplayName, outputType),
    );

    return res.end(buffer);
  } catch (error) {
    console.error("[automation.downloadGeneratedFile]", error);
    next(error);
  }
};

exports.createSummarySheet = async (req, res, next) => {
  try {
    const {
      queryTablesKey,
      message,
      intent,
      candidate,
      templateCandidate,
      executionResult,
      summarySheetMode = "hybrid",
      includeSourceDataSheet = true,
      formulaOptions = {},
    } = req.body || {};

    if (!queryTablesKey) {
      return res.status(400).json({
        ok: false,
        error: "queryTablesKey가 필요합니다.",
      });
    }
    if (!(await assertTemplateGenerationUsage(req, res))) return;

    const saved = await readJsonObject(queryTablesKey);
    const normalizedQueryTables =
      saved.normalizedQueryTables ||
      buildNormalizedQueryTables(saved.tables || []);

    let queryIntent = intent || null;
    let result = executionResult || null;

    if (!result && templateCandidate?.templateId) {
      result = executeBusinessTemplate({
        normalizedQueryTables,
        templateCandidate,
      });

      queryIntent = {
        ok: true,
        operation: templateCandidate.templateId,
        source: "business-template",
        templateCandidate,
      };
    }

    if (!result && candidate) {
      result = executeAnalysisRecipeCandidate({
        normalizedQueryTables,
        candidate,
      });

      queryIntent = {
        ok: true,
        operation:
          candidate.recipeType || candidate.type || "analysisCandidate",
        source: "analysis-candidate",
        candidate,
      };
    }

    if (!result) {
      queryIntent = intent || parseQueryIntent(message, saved.tables || []);

      if (!queryIntent?.ok) {
        return res.status(400).json({
          ok: false,
          error: queryIntent?.error || "query intent 생성 실패",
          intent: queryIntent,
        });
      }

      result = executeQueryIntent(saved.tables || [], queryIntent);
    }

    result = normalizeExecutedResult(result, templateCandidate);

    if (!result?.ok) {
      return res.status(400).json({
        ok: false,
        error: result?.error || result?.message || "query 실행 실패",
        intent: queryIntent,
        result,
      });
    }

    const chartSpec = Array.isArray(result.sections)
      ? null
      : buildChartSpec(result);

    const workbook = buildSummaryWorkbook({
      fileName: saved.fileName,
      message,
      intent: queryIntent,
      result,
      sourceTables: normalizedQueryTables,
      summarySheetMode,
      includeSourceDataSheet,
      formulaOptions,
    });

    const buffer = workbookToBuffer(workbook);
    const formulaEngineMeta = workbook["!beebeeFormulaEngine"] || {
      prepared: true,
      applied: false,
      mode: summarySheetMode,
      formulaCount: 0,
    };

    const userId = req.user?.id || "local-dev";
    const outputType = OUTPUT_TYPES.SUMMARY_SHEET;
    const outputFileName = buildGeneratedFileName({
      sourceFileName: saved.fileName,
      templateTitle: resultTemplateTitle(
        result,
        templateCandidate,
        getOutputDefaultTitle(outputType),
      ),
      outputType,
    });
    const key = `${SUMMARY_SHEET_STORAGE_PREFIX}/${userId}/${saved.fileHash}/${outputFileName}`;

    const stored = await saveBufferObject(
      key,
      buffer,
      getOutputMimeType(outputType),
    );
    await bumpTemplateGenerationUsage(req);

    return res.json({
      ok: true,
      fileName: outputFileName,
      displayName: outputFileName,
      downloadUrl: buildGeneratedDownloadUrl({
        storageKey: key,
        displayName: outputFileName,
        outputType,
      }),
      sourceFileName: saved.fileName,
      outputType,
      outputLabel: outputTypeLabel(outputType),
      fileHash: saved.fileHash,
      queryTablesKey,
      summarySheetKey: key,
      storageKey: key,
      internalFileKey: key,
      localName: stored.localName,
      gcsName: stored.gcsName,
      sheetNames: workbook.SheetNames || [],
      summarySheetMode,
      includeSourceDataSheet,
      formulaEngine: formulaEngineMeta,
      chartSpec,
      intent: queryIntent,
      result,
    });
  } catch (e) {
    console.error("[automation.createSummarySheet]", e);
    next(e);
  }
};

function writeReportJson({
  fileName,
  message,
  result,
  templateCandidate = null,
}) {
  const outputType = OUTPUT_TYPES.ANALYSIS_REPORT;
  fs.mkdirSync(REPORT_DIR, { recursive: true });

  const report = buildReportSections({
    fileName,
    message,
    result,
  });

  const output = {
    ok: true,
    version: getOutputVersion(outputType),
    outputType,
    outputLabel: outputTypeLabel(outputType),
    generatedAt: new Date().toISOString(),
    source: {
      fileName: fileName || "",
      message: message || "",
    },
    result: {
      operation: result?.operation || "",
      resultType: result?.resultType || "",
      rowCount: Array.isArray(result?.rows)
        ? result.rows.length
        : Array.isArray(result?.sections)
          ? result.sections.reduce(
              (sum, s) =>
                sum +
                (Array.isArray(s.result?.rows) ? s.result.rows.length : 0),
              0,
            )
          : 0,
    },
    report,
  };

  const outputName = buildGeneratedFileName({
    sourceFileName: fileName,
    templateTitle: resultTemplateTitle(
      result,
      templateCandidate,
      getOutputDefaultTitle(outputType),
    ),
    outputType,
  });
  const filePath = path.join(REPORT_DIR, outputName);

  fs.writeFileSync(filePath, JSON.stringify(output, null, 2), "utf-8");

  return {
    ok: true,
    fileName: outputName,
    displayName: outputName,
    filePath,
    outputType,
    outputLabel: outputTypeLabel(outputType),
    report,
  };
}

async function writeReportPpt({
  fileName,
  message,
  result,
  template,
  templateCandidate = null,
}) {
  const outputType = OUTPUT_TYPES.PPT;
  fs.mkdirSync(PPT_DIR, { recursive: true });

  const report = buildReportSections({
    fileName,
    message,
    result,
  });

  const pptx = renderReportPpt(report, { template });

  const outputName = buildGeneratedFileName({
    sourceFileName: fileName,
    templateTitle: resultTemplateTitle(
      result,
      templateCandidate,
      getOutputDefaultTitle(outputType),
    ),
    outputType,
  });
  const filePath = path.join(PPT_DIR, outputName);

  await pptx.writeFile({ fileName: filePath });

  return {
    ok: true,
    fileName: outputName,
    displayName: outputName,
    filePath,
    outputType,
    outputLabel: outputTypeLabel(outputType),
    template: template || "default",
    report,
    slideCount: Array.isArray(report.sections) ? report.sections.length : 0,
  };
}

exports.exportXlsx = async (req, res) => {
  try {
    const {
      queryTablesKey,
      message,
      candidate,
      templateCandidate,
      executionResult,
    } = req.body || {};

    if (!queryTablesKey || !message) {
      return res.status(400).json({
        ok: false,
        code: "MISSING_REQUIRED_FIELDS",
        error: "queryTablesKey와 message가 필요합니다.",
      });
    }

    const saved = await readJsonObject(queryTablesKey);
    const tables = saved.tables || [];

    let intent = null;
    let result = executionResult || null;

    if (!result && templateCandidate?.templateId) {
      result = executeBusinessTemplate({
        normalizedQueryTables:
          saved.normalizedQueryTables ||
          buildNormalizedQueryTables(saved.tables || []),
        templateCandidate,
      });

      intent = {
        ok: true,
        operation: templateCandidate.templateId,
        source: "business-template",
        templateCandidate,
      };
    }

    if (!result && candidate) {
      result = executeAnalysisRecipeCandidate({
        normalizedQueryTables:
          saved.normalizedQueryTables ||
          buildNormalizedQueryTables(saved.tables || []),
        candidate,
      });

      intent = {
        ok: true,
        operation:
          candidate.recipeType || candidate.type || "analysisCandidate",
        source: "analysis-candidate",
        candidate,
      };
    }

    if (!result) {
      intent = parseQueryIntent(message, tables);

      if (!intent.ok) {
        return res.status(400).json({
          ok: false,
          intent,
          code: intent.code,
          error: intent.error || "query intent 생성 실패",
        });
      }

      result = executeQueryIntent(tables, intent);
    }

    result = normalizeExecutedResult(result, templateCandidate);

    if (!result.ok) {
      return res.status(400).json({
        ok: false,
        intent,
        result,
        error: result.error || "query 실행 실패",
      });
    }

    const workbook = buildSummaryWorkbook({
      fileName: saved.fileName || "",
      message,
      intent,
      result,
      sourceTables:
        saved.normalizedQueryTables ||
        buildNormalizedQueryTables(saved.tables || []),
      summarySheetMode: req.body?.summarySheetMode || "hybrid",
      includeSourceDataSheet: req.body?.includeSourceDataSheet !== false,
      formulaOptions: req.body?.formulaOptions || {},
    });

    const buffer = workbookToBuffer(workbook);

    const outputType = OUTPUT_TYPES.SUMMARY_SHEET;
    const fileName = buildGeneratedFileName({
      sourceFileName: saved.fileName || "",
      templateTitle: resultTemplateTitle(
        result,
        templateCandidate,
        getOutputDefaultTitle(outputType),
      ),
      outputType,
    });
    const outputDir = AUTOMATION_DIR;
    fs.mkdirSync(outputDir, { recursive: true });

    const filePath = path.join(outputDir, fileName);
    fs.writeFileSync(filePath, buffer);

    return res.json({
      ok: true,
      fileName,
      displayName: fileName,
      downloadUrl: buildGeneratedDownloadUrl({
        filePath,
        displayName: fileName,
        outputType,
      }),
      filePath,
      outputType,
      outputLabel: outputTypeLabel(outputType),
      sheetNames: workbook.SheetNames || [],
      formulaEngine: workbook["!beebeeFormulaEngine"] || {
        prepared: true,
        applied: false,
        mode: req.body?.summarySheetMode || "hybrid",
        formulaCount: 0,
      },
      result,
    });
  } catch (err) {
    return res.status(500).json({
      ok: false,
      code: "EXPORT_XLSX_FAILED",
      error: err.message,
    });
  }
};

exports.exportReportJson = async (req, res) => {
  try {
    const {
      queryTablesKey,
      message,
      candidate,
      templateCandidate,
      executionResult,
    } = req.body || {};

    if (!queryTablesKey || !message) {
      return res.status(400).json({
        ok: false,
        code: "MISSING_REQUIRED_FIELDS",
        error: "queryTablesKey와 message가 필요합니다.",
      });
    }
    if (!(await assertTemplateGenerationUsage(req, res))) return;

    const saved = await readJsonObject(queryTablesKey);
    const tables = saved.tables || [];

    let intent = null;
    let result = executionResult || null;

    if (!result && templateCandidate?.templateId) {
      result = executeBusinessTemplate({
        normalizedQueryTables:
          saved.normalizedQueryTables ||
          buildNormalizedQueryTables(saved.tables || []),
        templateCandidate,
      });

      intent = {
        ok: true,
        operation: templateCandidate.templateId,
        source: "business-template",
        templateCandidate,
      };
    }

    if (!result && candidate) {
      result = executeAnalysisRecipeCandidate({
        normalizedQueryTables:
          saved.normalizedQueryTables ||
          buildNormalizedQueryTables(saved.tables || []),
        candidate,
      });

      intent = {
        ok: true,
        operation:
          candidate.recipeType || candidate.type || "analysisCandidate",
        source: "analysis-candidate",
        candidate,
      };
    }

    if (!result) {
      intent = parseQueryIntent(message, tables);

      if (!intent.ok) {
        return res.status(400).json({
          ok: false,
          intent,
          code: intent.code,
          error: intent.error || "query intent 생성 실패",
        });
      }

      result = executeQueryIntent(tables, intent);
    }

    result = normalizeExecutedResult(result, templateCandidate);

    if (!result.ok) {
      return res.status(400).json({
        ok: false,
        intent,
        result,
        error: result.error || "query 실행 실패",
      });
    }

    const exported = writeReportJson({
      fileName: saved.fileName || "",
      message,
      result,
      templateCandidate,
    });
    const outputType = exported.outputType || OUTPUT_TYPES.ANALYSIS_REPORT;
    await bumpTemplateGenerationUsage(req);

    return res.json({
      ok: true,
      fileName: exported.fileName,
      displayName: exported.displayName || exported.fileName,
      downloadUrl: buildGeneratedDownloadUrl({
        filePath: exported.filePath,
        displayName: exported.displayName || exported.fileName,
        outputType,
      }),
      filePath: exported.filePath,
      outputType,
      outputLabel: outputTypeLabel(outputType),
      analysisReport: exported.report,
      report: exported.report,
      result,
    });
  } catch (err) {
    return res.status(500).json({
      ok: false,
      code: "EXPORT_REPORT_JSON_FAILED",
      error: err.message,
    });
  }
};

exports.exportAnalysisReport = exports.exportReportJson;

exports.exportPptx = async (req, res) => {
  try {
    const {
      queryTablesKey,
      message,
      template,
      candidate,
      templateCandidate,
      executionResult,
    } = req.body || {};

    if (!queryTablesKey || !message) {
      return res.status(400).json({
        ok: false,
        code: "MISSING_REQUIRED_FIELDS",
        error: "queryTablesKey와 message가 필요합니다.",
      });
    }
    if (!(await assertTemplateGenerationUsage(req, res))) return;

    const saved = await readJsonObject(queryTablesKey);
    const tables = saved.tables || [];

    let intent = null;
    let result = executionResult || null;

    if (!result && templateCandidate?.templateId) {
      result = executeBusinessTemplate({
        normalizedQueryTables:
          saved.normalizedQueryTables ||
          buildNormalizedQueryTables(saved.tables || []),
        templateCandidate,
      });

      intent = {
        ok: true,
        operation: templateCandidate.templateId,
        source: "business-template",
        templateCandidate,
      };
    }

    if (!result && candidate) {
      result = executeAnalysisRecipeCandidate({
        normalizedQueryTables:
          saved.normalizedQueryTables ||
          buildNormalizedQueryTables(saved.tables || []),
        candidate,
      });

      intent = {
        ok: true,
        operation:
          candidate.recipeType || candidate.type || "analysisCandidate",
        source: "analysis-candidate",
        candidate,
      };
    }

    if (!result) {
      intent = parseQueryIntent(message, tables);

      if (!intent.ok) {
        return res.status(400).json({
          ok: false,
          intent,
          code: intent.code,
          error: intent.error || "query intent 생성 실패",
        });
      }

      result = executeQueryIntent(tables, intent);
    }

    result = normalizeExecutedResult(result, templateCandidate);

    if (!result.ok) {
      return res.status(400).json({
        ok: false,
        intent,
        result,
        error: result.error || "query 실행 실패",
      });
    }

    const exported = await writeReportPpt({
      fileName: saved.fileName || "",
      message,
      result,
      template,
      templateCandidate,
    });
    const outputType = exported.outputType || OUTPUT_TYPES.PPT;
    await bumpTemplateGenerationUsage(req);

    return res.json({
      ok: true,
      fileName: exported.fileName,
      displayName: exported.displayName || exported.fileName,
      downloadUrl: buildGeneratedDownloadUrl({
        filePath: exported.filePath,
        displayName: exported.displayName || exported.fileName,
        outputType,
      }),
      filePath: exported.filePath,
      outputType,
      outputLabel: outputTypeLabel(outputType),
      template: exported.template,
      slideCount: exported.slideCount,
      report: exported.report,
      result,
    });
  } catch (err) {
    return res.status(500).json({
      ok: false,
      code: "EXPORT_PPTX_FAILED",
      error: err.message,
    });
  }
};

exports.getAnalysisCandidates = async (req, res, next) => {
  try {
    const { queryTablesKey, fileName } = req.body || {};

    let saved = null;
    let key = queryTablesKey || null;

    if (key) {
      saved = await readJsonObject(key);

      if (saved && hasMojibakeQueryPayload(saved) && saved.fileName) {
        console.warn(
          "[query-tables] mojibake detected. Rebuilding saved key.",
          {
            queryTablesKey: key,
            fileName: saved.fileName,
          },
        );

        const rebuilt = await buildQueryTablesForFile(req, saved.fileName);

        saved = {
          version: "query_tables_v4_text_csv_encoding",
          fileName: saved.fileName,
          fileHash: rebuilt.fileHash,
          sheetStateSig: rebuilt.sheetStateSig,
          tableCount: rebuilt.tables.length,
          createdAt: new Date().toISOString(),
          tables: rebuilt.tables,
          normalizedQueryTables: rebuilt.normalizedQueryTables,
          analysisRecipeCandidates: rebuilt.analysisRecipeCandidates,
          categoryCandidates: rebuilt.categoryCandidates,
          businessTemplateCandidates: rebuilt.businessTemplateCandidates,
        };

        await saveJsonObject(key, saved);
      }
    } else if (fileName) {
      const built = await buildQueryTablesForFile(req, fileName);
      const normalizedQueryTables =
        built.normalizedQueryTables || buildNormalizedQueryTables(built.tables);

      const candidateBundle = built.candidateGeneration
        ? built
        : await generateCandidateBundle({
            normalizedQueryTables,
            fileName,
            source: "analysis-candidates-file",
          });

      const analysisRecipeCandidates =
        candidateBundle.analysisRecipeCandidates || [];

      const candidates = normalizeAnalysisCandidates(analysisRecipeCandidates);

      const categoryCandidates = candidateBundle.categoryCandidates || [];

      const businessTemplateCandidates =
        candidateBundle.businessTemplateCandidates || [];
      const candidateGeneration = candidateBundle.candidateGeneration || null;

      console.log("[analysis-candidates]", {
        source: "file",
        fileName,
        analysisRecipeCandidates: analysisRecipeCandidates.length,
        categoryCandidates: categoryCandidates.length,
        businessTemplateCandidates: businessTemplateCandidates.length,
      });

      return res.json({
        ok: true,
        source: "file",
        fileName,
        fileHash: built.fileHash,
        sheetStateSig: built.sheetStateSig,
        normalizedQueryTables,
        analysisRecipeCandidates,
        candidates,
        categoryCandidates,
        businessTemplateCandidates,
        candidateGeneration,
      });
    }

    if (!saved) {
      return res.status(400).json({
        ok: false,
        code: "QUERY_TABLES_KEY_OR_FILE_NAME_REQUIRED",
        error: "queryTablesKey 또는 fileName이 필요합니다.",
      });
    }

    const normalizedQueryTables =
      saved.normalizedQueryTables ||
      buildNormalizedQueryTables(saved.tables || []);

    const candidateBundle = saved.candidateGeneration
      ? {
          analysisRecipeCandidates: saved.analysisRecipeCandidates || [],
          categoryCandidates: saved.categoryCandidates || [],
          businessTemplateCandidates: saved.businessTemplateCandidates || [],
          candidateGeneration: saved.candidateGeneration,
        }
      : await generateCandidateBundle({
          normalizedQueryTables,
          fileName: saved.fileName,
          source: "analysis-candidates-query-tables",
        });

    const analysisRecipeCandidates =
      candidateBundle.analysisRecipeCandidates || [];

    const candidates = normalizeAnalysisCandidates(analysisRecipeCandidates);
    const categoryCandidates = candidateBundle.categoryCandidates || [];

    const businessTemplateCandidates =
      candidateBundle.businessTemplateCandidates || [];
    const candidateGeneration = candidateBundle.candidateGeneration || null;

    console.log("[analysis-candidates]", {
      source: "query-tables",
      fileName: saved.fileName,
      analysisRecipeCandidates: analysisRecipeCandidates.length,
      categoryCandidates: categoryCandidates.length,
      businessTemplateCandidates: businessTemplateCandidates.length,
    });

    return res.json({
      ok: true,
      source: "query-tables",
      queryTablesKey: key,
      fileName: saved.fileName,
      fileHash: saved.fileHash,
      sheetStateSig: saved.sheetStateSig,
      normalizedQueryTables,
      analysisRecipeCandidates,
      candidates,
      categoryCandidates,
      businessTemplateCandidates,
      candidateGeneration,
    });
  } catch (e) {
    console.error("[automation.getAnalysisCandidates]", e);
    next(e);
  }
};

exports.executeQuery = async (req, res, next) => {
  try {
    const { queryTablesKey, message, intent } = req.body || {};

    if (!queryTablesKey) {
      return res.status(400).json({
        ok: false,
        error: "queryTablesKey가 필요합니다.",
      });
    }

    const saved = await readJsonObject(queryTablesKey);
    const queryIntent = intent || parseQueryIntent(message, saved.tables || []);

    if (!queryIntent?.ok) {
      return res.status(400).json({
        ok: false,
        error: queryIntent?.error || "query intent 생성 실패",
        intent: queryIntent,
      });
    }

    const result = executeQueryIntent(saved.tables || [], queryIntent);

    return res.json({
      ok: true,
      queryTablesKey,
      fileName: saved.fileName,
      fileHash: saved.fileHash,
      intent: queryIntent,
      result,
    });
  } catch (e) {
    console.error("[automation.executeQuery]", e);
    next(e);
  }
};

exports.analyzeQueryIntent = async (req, res, next) => {
  try {
    const { queryTablesKey, message } = req.body || {};

    if (!queryTablesKey || !message) {
      return res.status(400).json({
        ok: false,
        error: "queryTablesKey와 message가 필요합니다.",
      });
    }

    const saved = await readJsonObject(queryTablesKey);
    const intent = parseQueryIntent(message, saved.tables || []);

    return res.json({
      ok: true,
      queryTablesKey,
      fileName: saved.fileName,
      fileHash: saved.fileHash,
      intent,
    });
  } catch (e) {
    console.error("[automation.analyzeQueryIntent]", e);
    next(e);
  }
};

exports.previewQueryTables = async (req, res, next) => {
  try {
    const { fileName } = req.body || {};

    if (!fileName) {
      return res.status(400).json({
        ok: false,
        error: "fileName이 필요합니다.",
      });
    }

    const built = await buildQueryTablesForFile(req, fileName);
    const { fileHash, sheetStateSig, tables } = built;

    const normalizedQueryTables =
      built.normalizedQueryTables || buildNormalizedQueryTables(tables);

    const candidateBundle = built.candidateGeneration
      ? built
      : await generateCandidateBundle({
          normalizedQueryTables,
          fileName,
          source: "preview-query-tables",
        });

    const analysisRecipeCandidates =
      candidateBundle.analysisRecipeCandidates || [];
    const categoryCandidates = candidateBundle.categoryCandidates || [];
    const businessTemplateCandidates =
      candidateBundle.businessTemplateCandidates || [];
    const candidateGeneration = candidateBundle.candidateGeneration || null;

    return res.json({
      ok: true,
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
      normalizedQueryTables,
      analysisRecipeCandidates,
      categoryCandidates,
      businessTemplateCandidates,
      candidateGeneration,
      tables: tables.map((t) => ({
        source: t.source,
        confidence: t.confidence,
        isPrimary: !!t.isPrimary,
        tableId: t.tableId,
        tableName: t.tableName,
        sheetName: t.sheetName,
        isFallback: !!t.isFallback,
        rowCount: t.rowCount,
        columnCount: t.columns.length,
      })),
    });
  } catch (e) {
    console.error("[automation.previewQueryTables]", e);
    next(e);
  }
};

exports.saveQueryTables = async (req, res, next) => {
  try {
    const { fileName } = req.body || {};

    if (!fileName) {
      return res.status(400).json({
        ok: false,
        error: "fileName이 필요합니다.",
      });
    }

    const built = await buildQueryTablesForFile(req, fileName);
    const { fileHash, sheetStateSig, tables } = built;

    const normalizedQueryTables =
      built.normalizedQueryTables || buildNormalizedQueryTables(tables);

    const candidateBundle = built.candidateGeneration
      ? built
      : await generateCandidateBundle({
          normalizedQueryTables,
          fileName,
          source: "save-query-tables",
        });

    const analysisRecipeCandidates =
      candidateBundle.analysisRecipeCandidates || [];

    const categoryCandidates = candidateBundle.categoryCandidates || [];

    const businessTemplateCandidates =
      candidateBundle.businessTemplateCandidates || [];
    const candidateGeneration = candidateBundle.candidateGeneration || null;

    const now = new Date();
    const userId = req.user?.id || "local-dev";
    const rand = crypto.randomBytes(6).toString("hex");

    const key = `query-tables/${userId}/${fileHash}/${Date.now()}_${rand}.json`;

    const payload = {
      version: "query_tables_v4_text_csv_encoding",
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
      createdAt: now.toISOString(),
      tables,
      normalizedQueryTables,
      analysisRecipeCandidates,
      categoryCandidates,
      businessTemplateCandidates,
      candidateGeneration,
    };

    const saved = await saveJsonObject(key, payload);

    return res.json({
      ok: true,
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
      queryTablesKey: key,
      normalizedQueryTables,
      analysisRecipeCandidates,
      categoryCandidates,
      businessTemplateCandidates,
      candidateGeneration,
      localName: saved.localName,
      gcsName: saved.gcsName,
      tables: tables.map((t) => ({
        source: t.source,
        confidence: t.confidence,
        isPrimary: !!t.isPrimary,
        tableId: t.tableId,
        tableName: t.tableName,
        sheetName: t.sheetName,
        isFallback: !!t.isFallback,
        rowCount: t.rowCount,
        columnCount: t.columns.length,
      })),
    });
  } catch (e) {
    console.error("[automation.saveQueryTables]", e);

    if (e.status) {
      return res.status(e.status).json({
        ok: false,
        error: e.message,
        ...(e.payload || {}),
      });
    }

    next(e);
  }
};
