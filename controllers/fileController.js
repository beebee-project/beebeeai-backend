const { bumpUsage, assertCanUse } = require("../services/usageService");
const {
  uploadBufferToGCS,
  downloadToBuffer,
  deleteObject,
  isStorageEnabled,
} = require("../utils/storage");

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
      await deleteObject(existingFile.localName || existingFile.gcsName);

      user.uploadedFiles = user.uploadedFiles.filter(
        (f) => f.originalName !== originalName,
      );
    }

    const saved = await uploadBufferToGCS({
      userId: user._id,
      buffer: req.file.buffer,
      originalName,
    });

    const newFile = {
      originalName,
      gcsName: saved.gcsName || null,
      localName: saved.localName || null,
      size: req.file.size,
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

    const buffer = await downloadToBuffer(
      fileInfo.localName || fileInfo.gcsName,
    );
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

    await deleteObject(fileInfo.localName || fileInfo.gcsName);

    user.uploadedFiles = user.uploadedFiles.filter(
      (f) => f.originalName !== originalName,
    );
    await user.save();

    res.status(200).json(user.uploadedFiles);
  } catch (error) {
    next(error);
  }
};
