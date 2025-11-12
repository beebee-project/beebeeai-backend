const crypto = require("crypto");
const path = require("path");
const fs = require("fs");
const { fileTypeFromBuffer } = require("file-type");
const { getBucket } = require("../config/gcs");

function sha256(buf) {
  return crypto.createHash("sha256").update(buf).digest("hex");
}

async function sniffMime(buffer, originalName) {
  try {
    const ft = await fileTypeFromBuffer(buffer);
    if (ft) return { mime: ft.mime, ext: ft.ext };
  } catch (_) {}
  const ext = (originalName.split(".").pop() || "").toLowerCase();
  const map = {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    xls: "application/vnd.ms-excel",
    csv: "text/csv",
  };
  return { mime: map[ext] || "application/octet-stream", ext };
}

function gcsKey({ userId, originalName, hash }) {
  const ts = new Date();
  const yyyy = ts.getUTCFullYear();
  const mm = String(ts.getUTCMonth() + 1).padStart(2, "0");
  const base = path.basename(originalName);
  return `users/${userId}/${yyyy}/${mm}/${hash}-${base}`;
}

async function uploadBufferToGCS({ userId, buffer, originalName }) {
  const { mime } = await sniffMime(buffer, originalName);
  const hash = sha256(buffer);
  const key = gcsKey({ userId, originalName, hash });
  const file = getBucket().file(fileName);

  await file.save(buffer, {
    contentType: mime,
    resumable: false,
    metadata: {
      cacheControl: "private, max-age=0, no-transform",
      metadata: { sha256: hash, originalName },
    },
  });

  return { gcsName: key, mime, hash };
}

async function downloadToBuffer(gcsName) {
  const [buffer] = await bucket.file(gcsName).download();
  return buffer;
}

async function deleteObject(gcsName) {
  await bucket.file(gcsName).delete({ ignoreNotFound: true });
}

async function getSignedUrl(gcsName, { minutes = 5, dispositionName } = {}) {
  const [url] = await bucket.file(gcsName).getSignedUrl({
    version: "v4",
    action: "read",
    expires: Date.now() + minutes * 60 * 1000,
    responseDisposition: `attachment; filename="${
      dispositionName || gcsName.split("/").pop()
    }"`,
  });
  return url;
}

const META_CACHE_DIR = path.join(__dirname, "..", "meta-cache");
if (!fs.existsSync(META_CACHE_DIR)) {
  fs.mkdirSync(META_CACHE_DIR, { recursive: true });
}

function metaCachePath(key) {
  return path.join(META_CACHE_DIR, `${key}.json`);
}

async function readMetaCache(key) {
  const p = metaCachePath(key);
  try {
    if (!fs.existsSync(p)) return null;
    const raw = await fs.promises.readFile(p, "utf8");
    return JSON.parse(raw);
  } catch (e) {
    console.error("readMetaCache error:", e);
    return null;
  }
}

async function writeMetaCache(key, value) {
  const p = metaCachePath(key);
  try {
    await fs.promises.writeFile(p, JSON.stringify(value, null, 2), "utf8");
  } catch (e) {
    console.error("writeMetaCache error:", e);
  }
}

module.exports = {
  uploadBufferToGCS,
  downloadToBuffer,
  deleteObject,
  getSignedUrl,
  sniffMime,
  readMetaCache,
  writeMetaCache,
};
