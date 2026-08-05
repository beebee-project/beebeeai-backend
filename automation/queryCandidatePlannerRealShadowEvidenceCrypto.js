"use strict";

const crypto = require("crypto");

const ENCRYPTION_VERSION =
  "query_candidate_planner_real_shadow_evidence_aes_256_gcm_v1";

function keyFromSecret(secret) {
  return crypto.createHash("sha256").update(String(secret)).digest();
}

function encryptEvidencePayload(payload, secret) {
  if (String(secret || "").length < 32) {
    const error = new Error("real shadow evidence encryption secret required");
    error.code = "REAL_SHADOW_EVIDENCE_SECRET_REQUIRED";
    throw error;
  }
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", keyFromSecret(secret), iv);
  const ciphertext = Buffer.concat([
    cipher.update(JSON.stringify(payload), "utf8"),
    cipher.final(),
  ]);
  const authTag = cipher.getAuthTag();
  return Object.freeze({
    encryptionVersion: ENCRYPTION_VERSION,
    iv: iv.toString("base64"),
    authTag: authTag.toString("base64"),
    ciphertext: ciphertext.toString("base64"),
  });
}

function decryptEvidencePayload(encrypted, secret) {
  if (encrypted?.encryptionVersion !== ENCRYPTION_VERSION) {
    const error = new Error("unsupported evidence encryption version");
    error.code = "REAL_SHADOW_EVIDENCE_ENCRYPTION_VERSION_INVALID";
    throw error;
  }
  const decipher = crypto.createDecipheriv(
    "aes-256-gcm",
    keyFromSecret(secret),
    Buffer.from(encrypted.iv, "base64"),
  );
  decipher.setAuthTag(Buffer.from(encrypted.authTag, "base64"));
  const plaintext = Buffer.concat([
    decipher.update(Buffer.from(encrypted.ciphertext, "base64")),
    decipher.final(),
  ]).toString("utf8");
  return JSON.parse(plaintext);
}

module.exports = Object.freeze({
  ENCRYPTION_VERSION,
  encryptEvidencePayload,
  decryptEvidencePayload,
});
