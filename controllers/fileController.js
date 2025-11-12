const { Storage } = require("@google-cloud/storage");
const { v4: uuidv4 } = require("uuid");
const User = require("../models/User");

const storage = new Storage();
const bucket = storage.bucket(process.env.GCS_BUCKET_NAME);
const { bumpUsage } = require("../services/usageService");

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
    if (!req.file) {
      return res.status(400).json({ message: "파일이 업로드되지 않았습니다." });
    }

    const user = req.user;
    const originalName = Buffer.from(req.file.originalname, "latin1").toString(
      "utf8"
    );

    const existingFile = user.uploadedFiles.find(
      (f) => f.originalName === originalName
    );
    if (existingFile) {
      await bucket.file(existingFile.gcsName).delete();

      user.uploadedFiles = user.uploadedFiles.filter(
        (f) => f.originalName !== originalName
      );
    }

    const gcsName = `${user._id}-${uuidv4()}-${originalName}`;
    const file = bucket.file(gcsName);

    const stream = file.createWriteStream({
      metadata: {
        contentType: req.file.mimetype,
      },
    });

    stream.on("error", (err) => {
      next(err);
    });

    stream.on("finish", async () => {
      const newFile = {
        originalName,
        gcsName,
        size: req.file.size,
      };
      user.uploadedFiles.push(newFile);
      await user.save();

      res.status(201).json(user.uploadedFiles);
    });

    stream.end(req.file.buffer);
  } catch (error) {
    next(error);
  }
};

// 3. 파일 다운로드 API
exports.downloadFile = async (req, res, next) => {
  try {
    const user = req.user;
    const { originalName } = req.params;

    const fileInfo = user.uploadedFiles.find(
      (f) => f.originalName === originalName
    );

    if (!fileInfo) {
      return res
        .status(404)
        .json({ message: "파일을 찾을 수 없거나 접근 권한이 없습니다." });
    }

    const encodedFilename = encodeURIComponent(fileInfo.originalName);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodedFilename}`
    );

    bucket.file(fileInfo.gcsName).createReadStream().pipe(res);
  } catch (error) {
    next(error);
  }
};

// 4. 파일 삭제 API
exports.deleteFile = async (req, res, next) => {
  try {
    const user = req.user;
    const { originalName } = req.params;

    const fileInfo = user.uploadedFiles.find(
      (f) => f.originalName === originalName
    );

    if (!fileInfo) {
      return res
        .status(404)
        .json({ message: "파일을 찾을 수 없거나 접근 권한이 없습니다." });
    }

    await bucket.file(fileInfo.gcsName).delete();

    user.uploadedFiles = user.uploadedFiles.filter(
      (f) => f.originalName !== originalName
    );
    await user.save();

    res.status(200).json(user.uploadedFiles);
  } catch (error) {
    next(error);
  }
};
