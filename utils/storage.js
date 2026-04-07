const crypto = require("crypto");
const path = require("path");
const fs = require("fs");
const { Storage } = require("@google-cloud/storage");
const { fileTypeFromBuffer } = require("file-type");

const GCS_ENABLED = Boolean(
  process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON &&
  process.env.GCLOUD_PROJECT &&
  process.env.GCS_BUCKET_NAME,
);

// ==== 서비스 계정 JSON을 임시 파일로 저장하고,
//      GOOGLE_APPLICATION_CREDENTIALS 를 그 파일로 지정 ====

// 키 파일을 저장할 경로 (컨테이너 로컬 디스크, 재시작되면 사라져도 상관 없음)
const KEY_FILE_PATH = path.join(__dirname, "..", "gcs-key.json");

let storage = null;
let bucket = null;

if (GCS_ENABLED) {
  if (!fs.existsSync(KEY_FILE_PATH)) {
    fs.writeFileSync(
      KEY_FILE_PATH,
      process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON,
      "utf8",
    );
  }

  process.env.GOOGLE_APPLICATION_CREDENTIALS = KEY_FILE_PATH;

  storage = new Storage({
    projectId: process.env.GCLOUD_PROJECT,
  });

  bucket = storage.bucket(process.env.GCS_BUCKET_NAME);
} else {
  console.warn("[storage] GCS disabled - missing required env");
}

// ==== 이하 기존 코드 동일 ====

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
  if (!GCS_ENABLED || !bucket) {
    throw new Error("GCS storage is disabled");
  }
  const { mime } = await sniffMime(buffer, originalName);
  const hash = sha256(buffer);
  const key = gcsKey({ userId, originalName, hash });
  const file = bucket.file(key);

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

// =========================================================
// Meta cache (A안): in-memory only, TTL 30 minutes
// - NO disk writes
// - Cleared on restart/redeploy (intended)
// =========================================================
const META_CACHE_TTL_MS = 30 * 60 * 1000; // 30 minutes (fixed per decision)
const META_CACHE_MAX = Number(process.env.META_CACHE_MAX || "500"); // safety cap
const _metaMem = new Map(); // key -> { value, expiresAt, touchedAt }

function _now() {
  return Date.now();
}

function _gcMetaMem() {
  // opportunistic GC: remove expired first, then trim oldest if over capacity
  const now = _now();
  for (const [k, e] of _metaMem.entries()) {
    if (!e || e.expiresAt <= now) _metaMem.delete(k);
  }
  if (_metaMem.size <= META_CACHE_MAX) return;

  // trim by oldest touchedAt
  const entries = Array.from(_metaMem.entries());
  entries.sort((a, b) => (a[1].touchedAt || 0) - (b[1].touchedAt || 0));
  const overflow = _metaMem.size - META_CACHE_MAX;
  for (let i = 0; i < overflow; i++) {
    _metaMem.delete(entries[i][0]);
  }
}

async function readMetaCache(key) {
  if (!GCS_ENABLED || !bucket) return null;
  try {
    const e = _metaMem.get(key);
    if (!e) return null;
    const now = _now();
    if (e.expiresAt <= now) {
      _metaMem.delete(key);
      return null;
    }
    e.touchedAt = now;
    return e.value;
  } catch (e) {
    console.error("readMetaCache(mem) error:", e);
    return null;
  } finally {
    _gcMetaMem();
  }
}

async function writeMetaCache(key, value) {
  if (!GCS_ENABLED || !bucket) return null;
  try {
    const now = _now();
    _metaMem.set(key, {
      value,
      expiresAt: now + META_CACHE_TTL_MS,
      touchedAt: now,
    });
  } catch (e) {
    console.error("writeMetaCache(mem) error:", e);
  } finally {
    _gcMetaMem();
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
  isStorageEnabled: () => GCS_ENABLED,
};
