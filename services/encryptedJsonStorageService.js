const crypto = require("crypto");
const { Storage } = require("@google-cloud/storage");
const { downloadToBuffer } = require("../utils/storage");

const HAS_GCS_ENV = Boolean(
  process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON &&
  process.env.GCLOUD_PROJECT &&
  process.env.GCS_BUCKET_NAME,
);

const storage = HAS_GCS_ENV
  ? new Storage({
      projectId: process.env.GCLOUD_PROJECT,
      credentials: JSON.parse(process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON),
    })
  : null;

const BUCKET_NAME = process.env.GCS_BUCKET_NAME;
const QUERY_JSON_PREFIX =
  process.env.QUERY_JSON_GCS_PREFIX || "query-json/encrypted";

const SECRET =
  process.env.QUERY_JSON_SECRET || process.env.JWT_SECRET || "dev-query-secret";

function getKey() {
  return crypto.createHash("sha256").update(SECRET).digest();
}

function encryptJson(payload) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", getKey(), iv);

  const encrypted = Buffer.concat([
    cipher.update(JSON.stringify(payload), "utf8"),
    cipher.final(),
  ]);

  return {
    iv: iv.toString("base64"),
    tag: cipher.getAuthTag().toString("base64"),
    data: encrypted.toString("base64"),
  };
}

function decryptJson(encrypted = {}) {
  const iv = Buffer.from(encrypted.iv || "", "base64");
  const tag = Buffer.from(encrypted.tag || "", "base64");
  const data = Buffer.from(encrypted.data || "", "base64");

  const decipher = crypto.createDecipheriv("aes-256-gcm", getKey(), iv);
  decipher.setAuthTag(tag);

  const decrypted = Buffer.concat([decipher.update(data), decipher.final()]);

  return JSON.parse(decrypted.toString("utf8"));
}

async function readEncryptedQueryJson(queryJsonKey) {
  if (!queryJsonKey) return null;

  const buffer = await downloadToBuffer(queryJsonKey);
  const body = JSON.parse(buffer.toString("utf8"));

  if (!body?.encrypted) {
    return null;
  }

  return decryptJson(body.encrypted);
}

function buildQueryJsonKey(userId) {
  const now = new Date();
  const yyyy = now.getUTCFullYear();
  const mm = String(now.getUTCMonth() + 1).padStart(2, "0");
  const safeUserId = String(userId || "anonymous").replace(/[^\w.-]/g, "_");
  const id = crypto.randomUUID();

  return `${QUERY_JSON_PREFIX}/${safeUserId}/${yyyy}/${mm}/${id}.json`;
}

async function saveEncryptedQueryJson({ userId, fileName, payload }) {
  if (!storage || !BUCKET_NAME) {
    console.warn("[queryJsonStorage] GCS disabled - skip save");
    return null;
  }

  const bucket = storage.bucket(BUCKET_NAME);
  const key = buildQueryJsonKey(userId);

  const body = {
    version: "encrypted_query_json_v1",
    userId: String(userId || "anonymous"),
    fileName,
    createdAt: new Date().toISOString(),
    encrypted: encryptJson(payload),
  };

  await bucket.file(key).save(JSON.stringify(body), {
    contentType: "application/json; charset=utf-8",
    resumable: false,
    metadata: {
      cacheControl: "private, max-age=0, no-store",
      metadata: {
        fileName: String(fileName || ""),
      },
    },
  });

  return {
    queryJsonKey: key,
  };
}

async function deleteEncryptedQueryJson(queryJsonKey) {
  if (!storage || !BUCKET_NAME || !queryJsonKey) return;

  const bucket = storage.bucket(BUCKET_NAME);
  await bucket.file(queryJsonKey).delete({ ignoreNotFound: true });
}

function normalizeFileNameForCompare(value = "") {
  return String(value || "").trim();
}

function extractEncryptedQueryFileName(body = {}) {
  if (!body || typeof body !== "object") return "";

  if (body.fileName) return body.fileName;

  try {
    const payload = body.encrypted ? decryptJson(body.encrypted) : null;
    return payload?.fileName || "";
  } catch (error) {
    return "";
  }
}

async function deleteEncryptedQueryJsonByFileName({ userId, fileName }) {
  if (!storage || !BUCKET_NAME || !userId || !fileName) return null;

  const safeUserId = String(userId).replace(/[^\w.-]/g, "_");
  const prefix = QUERY_JSON_PREFIX + "/" + safeUserId + "/";
  const targetFileName = normalizeFileNameForCompare(fileName);

  const bucket = storage.bucket(BUCKET_NAME);
  const [files] = await bucket.getFiles({ prefix });

  const deleted = [];
  const errors = [];
  let scanned = 0;

  for (const file of files) {
    scanned += 1;

    try {
      let matched = false;
      const [metadata] = await file.getMetadata();
      const metaFileName = normalizeFileNameForCompare(
        metadata?.metadata?.fileName,
      );

      if (metaFileName && metaFileName === targetFileName) {
        matched = true;
      }

      if (!matched) {
        const [buffer] = await file.download();
        const body = JSON.parse(buffer.toString("utf8"));
        const bodyFileName = normalizeFileNameForCompare(
          extractEncryptedQueryFileName(body),
        );

        matched = bodyFileName === targetFileName;
      }

      if (matched) {
        await file.delete({ ignoreNotFound: true });
        deleted.push(file.name);
      }
    } catch (error) {
      errors.push({
        key: file.name,
        message: error?.message || String(error),
      });
    }
  }

  const result = {
    prefix,
    fileName: targetFileName,
    scanned,
    deleted,
    errors,
  };

  console.log("[queryJson.deleteByFileName]", result);

  return result;
}

module.exports = {
  saveEncryptedQueryJson,
  deleteEncryptedQueryJson,
  deleteEncryptedQueryJsonByFileName,
  readEncryptedQueryJson,
};
