const { bumpUsage, assertCanUse } = require("../services/usageService");
const {
  uploadBufferToGCS,
  downloadToBuffer,
  deleteObject,
  isStorageEnabled,
} = require("../utils/storage");
const XLSX = require("xlsx");
const { readWorkbookFromBuffer } = require("../utils/workbookReader");
const { getOrBuildAllSheetsData } = require("../utils/sheetPreprocessor");
const {
  buildQueryTablesFromWorkbook,
} = require("../automation/queryTableBuilder");
const {
  buildNormalizedQueryTables,
} = require("../automation/normalizedQueryTableBuilder");
const {
  buildAnalysisRecipeCandidates,
} = require("../automation/analysisRecipeCandidateBuilder");
const {
  buildBusinessTemplateCandidates,
} = require("../automation/businessTemplateConfig");
const {
  saveEncryptedQueryJson,
  deleteEncryptedQueryJson,
} = require("../services/encryptedJsonStorageService");
const {
  encryptBuffer,
  decryptBuffer,
} = require("../services/encryptedFileService");

exports.upload = async (req, res) => {
  const saved = await saveToStorage(req);
  if (!saved) return res.status(500).json({ error: "업로드 실패" });

  await bumpUsage(req.user.id, "fileUploads", +1);
  res.json({ ok: true, file: saved });
};

exports.remove = async (req, res) => {
  const removed = await removeFromStorage(req.params.id);
  if (!removed)
    return res.status(404).json({ error: "이미 삭제되었거나 없음" });

  await bumpUsage(req.user.id, "fileUploads", -1);
  res.json({ ok: true });
};

// 1. 파일 목록 조회 API
exports.getFiles = async (req, res, next) => {
  try {
    res.status(200).json(req.user.uploadedFiles);
  } catch (error) {
    next(error);
  }
};

// 2. 파일 업로드 API
exports.uploadFile = async (req, res, next) => {
  try {
    if (!isStorageEnabled()) {
      return res.status(503).json({
        error: "File storage is disabled in local dev",
        code: "STORAGE_DISABLED",
      });
    }

    if (!req.file) {
      return res.status(400).json({ message: "파일이 업로드되지 않았습니다." });
    }

    try {
      await assertCanUse(req.user.id, "fileUploads", 1);
    } catch (e) {
      return res.status(e.status || 429).json({
        error: "Usage limit exceeded",
        code: e.code || "LIMIT_EXCEEDED",
        ...e.meta,
      });
    }

    const user = req.user;
    const originalName = Buffer.from(req.file.originalname, "latin1").toString(
      "utf8",
    );

    const existingFile = user.uploadedFiles.find(
      (f) => f.originalName === originalName,
    );

    if (existingFile) {
      if (existingFile.queryJsonKey) {
        await deleteEncryptedQueryJson(existingFile.queryJsonKey);
      }

      if (existingFile.gcsName) {
        await deleteObject(existingFile.gcsName);
      }

      user.uploadedFiles = user.uploadedFiles.filter(
        (f) => f.originalName !== originalName,
      );
    }

    const encryptedFile = encryptBuffer(req.file.buffer);

    const saved = await uploadBufferToGCS({
      userId: user._id,
      buffer: encryptedFile.buffer,
      originalName,
      metadata: encryptedFile.metadata,
    });

    let queryJsonMeta = null;

    try {
      const { fileHash, allSheetsData, sheetStateSig } =
        await getOrBuildAllSheetsData(req.file.buffer);

      const workbook = readWorkbookFromBuffer(req.file.buffer);

      const tables = buildQueryTablesFromWorkbook(workbook, allSheetsData);
      const normalizedQueryTables = buildNormalizedQueryTables(tables);
      const analysisRecipeCandidates = buildAnalysisRecipeCandidates(
        normalizedQueryTables,
      );
      const businessTemplateCandidates = buildBusinessTemplateCandidates(
        analysisRecipeCandidates,
        { fileName: originalName },
      );

      queryJsonMeta = await saveEncryptedQueryJson({
        userId: String(user._id),
        fileName: originalName,
        payload: {
          version: "query_tables_v4_text_csv_encoding",
          fileName: originalName,
          fileHash,
          sheetStateSig,
          tableCount: tables.length,
          createdAt: new Date().toISOString(),
          tables,
          normalizedQueryTables,
          analysisRecipeCandidates,
          businessTemplateCandidates,
        },
      });
    } catch (error) {
      console.error("[file.upload.autoQueryJson]", error);
    }

    const newFile = {
      originalName,
      gcsName: saved.gcsName,
      localName: saved.localName,
      size: req.file.size,
      encrypted: true,
      encryptionVersion: encryptedFile.metadata.encryptionVersion,
      encryptionIv: encryptedFile.metadata.encryptionIv,
      encryptionTag: encryptedFile.metadata.encryptionTag,
      queryJsonKey: queryJsonMeta?.queryJsonKey || null,
    };
    user.uploadedFiles.push(newFile);
    await user.save();

    await bumpUsage(req.user.id, "fileUploads", 1);
    res.status(201).json(user.uploadedFiles);
  } catch (error) {
    next(error);
  }
};

// 3. 파일 다운로드 API
exports.downloadFile = async (req, res, next) => {
  try {
    if (!isStorageEnabled()) {
      return res.status(503).json({
        error: "File storage is disabled in local dev",
        code: "STORAGE_DISABLED",
      });
    }
    const user = req.user;
    const { originalName } = req.params;

    const fileInfo = user.uploadedFiles.find(
      (f) => f.originalName === originalName,
    );

    if (!fileInfo) {
      return res
        .status(404)
        .json({ message: "파일을 찾을 수 없거나 접근 권한이 없습니다." });
    }

    const encodedFilename = encodeURIComponent(fileInfo.originalName);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodedFilename}`,
    );

    let buffer = await downloadToBuffer(fileInfo.localName || fileInfo.gcsName);

    if (fileInfo.encrypted) {
      buffer = decryptBuffer(buffer, {
        encryptionIv: fileInfo.encryptionIv,
        encryptionTag: fileInfo.encryptionTag,
      });
    }

    res.end(buffer);
  } catch (error) {
    next(error);
  }
};

// 4. 파일 삭제 API
exports.deleteFile = async (req, res, next) => {
  try {
    if (!isStorageEnabled()) {
      return res.status(503).json({
        error: "File storage is disabled in local dev",
        code: "STORAGE_DISABLED",
      });
    }
    const user = req.user;
    const { originalName } = req.params;

    const fileInfo = user.uploadedFiles.find(
      (f) => f.originalName === originalName,
    );

    if (!fileInfo) {
      return res
        .status(404)
        .json({ message: "파일을 찾을 수 없거나 접근 권한이 없습니다." });
    }

    if (fileInfo.queryJsonKey) {
      await deleteEncryptedQueryJson(fileInfo.queryJsonKey);
    }

    if (fileInfo.gcsName) {
      await deleteObject(fileInfo.gcsName);
    }

    user.uploadedFiles = user.uploadedFiles.filter(
      (f) => f.originalName !== originalName,
    );

    await user.save();

    res.status(200).json(user.uploadedFiles);
  } catch (error) {
    next(error);
  }
};
