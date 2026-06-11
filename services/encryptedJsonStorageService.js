const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const QUERY_JSON_DIR =
  process.env.QUERY_JSON_DIR ||
  path.join(process.cwd(), ".local_uploads/query-json");

const QUERY_JSON_TTL_MS =
  Number(process.env.QUERY_JSON_TTL_MS) || 1000 * 60 * 30;

const SECRET =
  process.env.QUERY_JSON_SECRET || process.env.JWT_SECRET || "dev-query-secret";

function getKey() {
  return crypto.createHash("sha256").update(SECRET).digest();
}

function ensureDir() {
  fs.mkdirSync(QUERY_JSON_DIR, { recursive: true });
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

function deleteEncryptedQueryJson(queryJsonKey) {
  if (!queryJsonKey) return;

  ensureDir();

  const safeName = path.basename(queryJsonKey);
  const filePath = path.join(QUERY_JSON_DIR, safeName);

  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }
}

function saveEncryptedQueryJson({ userId, fileName, payload }) {
  ensureDir();

  const safeUserId = String(userId || "anonymous").replace(/[^\w.-]/g, "_");
  const id = crypto.randomUUID();
  const filePath = path.join(QUERY_JSON_DIR, `${safeUserId}_${id}.json`);

  const body = {
    version: "encrypted_query_json_v1",
    userId: safeUserId,
    fileName,
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + QUERY_JSON_TTL_MS).toISOString(),
    encrypted: encryptJson(payload),
  };

  fs.writeFileSync(filePath, JSON.stringify(body), "utf8");

  return {
    queryJsonKey: path.basename(filePath),
    queryJsonPath: filePath,
    expiresAt: body.expiresAt,
  };
}

function cleanupExpiredQueryJson() {
  ensureDir();

  const now = Date.now();

  fs.readdirSync(QUERY_JSON_DIR).forEach((name) => {
    const filePath = path.join(QUERY_JSON_DIR, name);

    try {
      const raw = JSON.parse(fs.readFileSync(filePath, "utf8"));
      if (raw.expiresAt && new Date(raw.expiresAt).getTime() <= now) {
        fs.unlinkSync(filePath);
      }
    } catch (_) {
      // 깨진 파일은 정리
      fs.unlinkSync(filePath);
    }
  });
}

module.exports = {
  saveEncryptedQueryJson,
  cleanupExpiredQueryJson,
  deleteEncryptedQueryJson,
};
