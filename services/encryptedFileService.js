const crypto = require("crypto");

const SECRET =
  process.env.FILE_ENCRYPTION_SECRET ||
  process.env.QUERY_JSON_SECRET ||
  process.env.JWT_SECRET ||
  "dev-file-secret";

function getKey() {
  return crypto.createHash("sha256").update(SECRET).digest();
}

function encryptBuffer(buffer) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", getKey(), iv);

  const encrypted = Buffer.concat([cipher.update(buffer), cipher.final()]);

  return {
    buffer: encrypted,
    metadata: {
      encryptionVersion: "file_aes_256_gcm_v1",
      encryptionIv: iv.toString("base64"),
      encryptionTag: cipher.getAuthTag().toString("base64"),
    },
  };
}

function decryptBuffer(buffer, metadata = {}) {
  const iv = Buffer.from(metadata.encryptionIv || "", "base64");
  const tag = Buffer.from(metadata.encryptionTag || "", "base64");

  const decipher = crypto.createDecipheriv("aes-256-gcm", getKey(), iv);
  decipher.setAuthTag(tag);

  return Buffer.concat([decipher.update(buffer), decipher.final()]);
}

module.exports = {
  encryptBuffer,
  decryptBuffer,
};
