const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const User = require("../models/User");
const { downloadToBuffer } = require("../utils/storage");
const { getOrBuildAllSheetsData } = require("../utils/sheetPreprocessor");
const {
  buildQueryTablesFromWorkbook,
} = require("../automation/queryTableBuilder");

function findUserFile(user, fileName) {
  if (!user || !fileName) return null;
  return user.uploadedFiles?.find((f) => f.originalName === fileName) || null;
}

exports.previewQueryTables = async (req, res, next) => {
  try {
    const { fileName } = req.body || {};

    if (!fileName) {
      return res.status(400).json({
        ok: false,
        error: "fileName이 필요합니다.",
      });
    }

    let buffer;

    if (process.env.LOCAL_DEV === "1" && process.env.DEV_BYPASS_AUTH === "1") {
      const localPath = path.join(__dirname, "..", ".local_uploads", fileName);

      if (!fs.existsSync(localPath)) {
        return res.status(404).json({
          ok: false,
          error: "로컬 테스트 파일을 찾을 수 없습니다.",
          path: localPath,
        });
      }

      buffer = fs.readFileSync(localPath);
    } else {
      const user = await User.findById(req.user.id).select("uploadedFiles");
      if (!user) {
        return res.status(404).json({ ok: false, error: "사용자 없음" });
      }

      const fileInfo = findUserFile(user, fileName);
      if (!fileInfo) {
        return res.status(404).json({
          ok: false,
          error: "파일을 찾을 수 없습니다.",
        });
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

    return res.json({
      ok: true,
      fileName,
      fileHash,
      sheetStateSig,
      tableCount: tables.length,
      tables: tables.map((t) => ({
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
