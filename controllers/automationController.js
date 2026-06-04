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
  workbookToBuffer,
  buildChartSpec,
} = require("../automation/summarySheetBuilder");

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

    const workbook = buildSummaryWorkbook({
      fileName: saved.fileName,
      message,
      intent: queryIntent,
      result,
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
      chartSpec,
      intent: queryIntent,
      result,
    });
  } catch (e) {
    console.error("[automation.createSummarySheet]", e);
    next(e);
  }
};

exports.exportXlsx = async (req, res) => {
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

    return res.json({
      ok: true,
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
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
    };

    const saved = await saveJsonObject(key, payload);

    return res.json({
      ok: true,
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
      queryTablesKey: key,
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
