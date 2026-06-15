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
const { getOrBuildAllSheetsData } = require("../utils/sheetPreprocessor");
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
  buildAnalysisRecipeCandidates,
} = require("../automation/analysisRecipeCandidateBuilder");
const {
  executeAnalysisRecipeCandidate,
} = require("../automation/analysisRecipeExecutor");

const REPORT_DIR = path.join(
  process.cwd(),
  ".local_uploads",
  "generated",
  "reports",
);

const PPT_DIR = path.join(process.cwd(), ".local_uploads", "generated", "ppt");

function findUserFile(user, fileName) {
  if (!user || !fileName) return null;
  return user.uploadedFiles?.find((f) => f.originalName === fileName) || null;
}

async function buildQueryTablesForFile(req, fileName) {
  let buffer;

  if (process.env.LOCAL_DEV === "1" && process.env.DEV_BYPASS_AUTH === "1") {
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
  }

  const { fileHash, allSheetsData, sheetStateSig } =
    await getOrBuildAllSheetsData(buffer);

  const workbook = XLSX.read(buffer, {
    type: "buffer",
    cellDates: true,
    cellNF: true,
    cellText: false,
  });

  const tables = buildQueryTablesFromWorkbook(workbook, allSheetsData);

  return { fileHash, sheetStateSig, tables };
}

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
        "automation",
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

exports.executeAnalysisCandidate = executeAnalysisCandidate;

exports.createSummarySheet = async (req, res, next) => {
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
    const chartSpec = buildChartSpec(result);

    if (!result?.ok) {
      return res.status(400).json({
        ok: false,
        error: result?.error || "query 실행 실패",
        intent: queryIntent,
        result,
      });
    }

    const workbook = buildAutomationTemplateWorkbook({
      fileName: saved.fileName,
      message,
      intent: queryIntent,
      result,
      tables: saved.tables || [],
    });

    const buffer = workbookToBuffer(workbook);

    const userId = req.user?.id || "local-dev";
    const rand = crypto.randomBytes(6).toString("hex");
    const key = `summary-sheets/${userId}/${saved.fileHash}/${Date.now()}_${rand}.xlsx`;

    const stored = await saveBufferObject(
      key,
      buffer,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );

    return res.json({
      ok: true,
      fileName: saved.fileName,
      fileHash: saved.fileHash,
      queryTablesKey,
      summarySheetKey: key,
      localName: stored.localName,
      gcsName: stored.gcsName,
      sheetNames: workbook.SheetNames || [],
      chartSpec,
      intent: queryIntent,
      result,
    });
  } catch (e) {
    console.error("[automation.createSummarySheet]", e);
    next(e);
  }
};

function writeReportJson({ fileName, message, result }) {
  fs.mkdirSync(REPORT_DIR, { recursive: true });

  const report = buildReportSections({
    fileName,
    message,
    result,
  });

  const output = {
    ok: true,
    version: "report_export_v1",
    generatedAt: new Date().toISOString(),
    source: {
      fileName: fileName || "",
      message: message || "",
    },
    result: {
      operation: result?.operation || "",
      resultType: result?.resultType || "",
      rowCount: Array.isArray(result?.rows) ? result.rows.length : 0,
    },
    report,
  };

  const outputName = `report_${Date.now()}.json`;
  const filePath = path.join(REPORT_DIR, outputName);

  fs.writeFileSync(filePath, JSON.stringify(output, null, 2), "utf-8");

  return {
    ok: true,
    fileName: outputName,
    filePath,
    report,
  };
}

async function writeReportPpt({ fileName, message, result, template }) {
  fs.mkdirSync(PPT_DIR, { recursive: true });

  const report = buildReportSections({
    fileName,
    message,
    result,
  });

  const pptx = renderReportPpt(report, { template });

  const outputName = `report_${Date.now()}.pptx`;
  const filePath = path.join(PPT_DIR, outputName);

  await pptx.writeFile({ fileName: filePath });

  return {
    ok: true,
    fileName: outputName,
    filePath,
    template: template || "default",
    report,
    slideCount: Array.isArray(report.sections) ? report.sections.length : 0,
  };
}

exports.exportXlsx = async (req, res) => {
  try {
    const { queryTablesKey, message, candidate, executionResult } =
      req.body || {};

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
    });

    const buffer = workbookToBuffer(workbook);

    const fileName = `automation_${Date.now()}.xlsx`;
    const outputDir = path.join(
      process.cwd(),
      ".local_uploads",
      "generated",
      "automation",
    );
    fs.mkdirSync(outputDir, { recursive: true });

    const filePath = path.join(outputDir, fileName);
    fs.writeFileSync(filePath, buffer);

    return res.json({
      ok: true,
      fileName,
      filePath,
      sheetNames: workbook.SheetNames || [],
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
    const { queryTablesKey, message } = req.body || {};

    if (!queryTablesKey || !message) {
      return res.status(400).json({
        ok: false,
        code: "MISSING_REQUIRED_FIELDS",
        error: "queryTablesKey와 message가 필요합니다.",
      });
    }

    const saved = await readJsonObject(queryTablesKey);
    const tables = saved.tables || [];

    const intent = parseQueryIntent(message, tables);

    if (!intent.ok) {
      return res.status(400).json({
        ok: false,
        intent,
        code: intent.code,
        error: intent.error || "query intent 생성 실패",
      });
    }

    const result = executeQueryIntent(tables, intent);

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
    });

    return res.json({
      ok: true,
      fileName: exported.fileName,
      filePath: exported.filePath,
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

exports.exportPptx = async (req, res) => {
  try {
    const { queryTablesKey, message, template } = req.body || {};

    if (!queryTablesKey || !message) {
      return res.status(400).json({
        ok: false,
        code: "MISSING_REQUIRED_FIELDS",
        error: "queryTablesKey와 message가 필요합니다.",
      });
    }

    const saved = await readJsonObject(queryTablesKey);
    const tables = saved.tables || [];

    const intent = parseQueryIntent(message, tables);

    if (!intent.ok) {
      return res.status(400).json({
        ok: false,
        intent,
        code: intent.code,
        error: intent.error || "query intent 생성 실패",
      });
    }

    const result = executeQueryIntent(tables, intent);

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
    });

    return res.json({
      ok: true,
      fileName: exported.fileName,
      filePath: exported.filePath,
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
    } else if (fileName) {
      const built = await buildQueryTablesForFile(req, fileName);
      const normalizedQueryTables = buildNormalizedQueryTables(built.tables);
      const analysisRecipeCandidates = buildAnalysisRecipeCandidates(
        normalizedQueryTables,
      );
      const candidates = normalizeAnalysisCandidates(analysisRecipeCandidates);

      return res.json({
        ok: true,
        source: "file",
        fileName,
        fileHash: built.fileHash,
        sheetStateSig: built.sheetStateSig,
        normalizedQueryTables,
        analysisRecipeCandidates,
        candidates,
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

    const analysisRecipeCandidates =
      saved.analysisRecipeCandidates ||
      buildAnalysisRecipeCandidates(normalizedQueryTables);

    const candidates = normalizeAnalysisCandidates(analysisRecipeCandidates);

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

    const { fileHash, sheetStateSig, tables } = await buildQueryTablesForFile(
      req,
      fileName,
    );

    const normalizedQueryTables = buildNormalizedQueryTables(tables);

    const analysisRecipeCandidates = buildAnalysisRecipeCandidates(
      normalizedQueryTables,
    );

    return res.json({
      ok: true,
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
      normalizedQueryTables,
      analysisRecipeCandidates,
      tables: tables.map((t) => ({
        source: t.source,
        confidence: t.confidence,
        isPrimary: !!t.isPrimary,
        tableId: t.tableId,
        tableName: t.tableName,
        sheetName: t.sheetName,
        isFallback: !!t.isFallback,
        range: t.range,
        dataRange: t.dataRange,
        rowCount: t.rowCount,
        columns: t.columns.map((c) => ({
          key: c.key,
          header: c.header,
          type: c.type,
          sampleValues: c.sampleValues,
          uniqueCount: c.uniqueCount,
          uniqueRatio: c.uniqueRatio,
        })),
        previewRows: t.rows.slice(0, 10),
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

    const { fileHash, sheetStateSig, tables } = await buildQueryTablesForFile(
      req,
      fileName,
    );

    const normalizedQueryTables = buildNormalizedQueryTables(tables);

    const analysisRecipeCandidates = buildAnalysisRecipeCandidates(
      normalizedQueryTables,
    );

    const now = new Date();
    const userId = req.user?.id || "local-dev";
    const rand = crypto.randomBytes(6).toString("hex");

    const key = `query-tables/${userId}/${fileHash}/${Date.now()}_${rand}.json`;

    const payload = {
      version: "query_tables_v1",
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
      createdAt: now.toISOString(),
      tables,
      normalizedQueryTables,
      analysisRecipeCandidates,
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
