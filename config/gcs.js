const fs = require("fs");
const path = require("path");
const { Storage } = require("@google-cloud/storage");

function ensureGcpCredentials() {
  // Railway에 GCS_KEY(서비스 계정 JSON)로 넣어두면 여기서 파일로 저장해 ADC 사용
  const json = process.env.GCS_KEY;
  if (!json) return; // 이미 GOOGLE_APPLICATION_CREDENTIALS를 쓴다면 생략
  const keyPath = path.join("/app", ".gcp-key.json");
  // 문자열/JSON 모두 처리
  const content = typeof json === "string" ? json : JSON.stringify(json);
  fs.writeFileSync(keyPath, content);
  process.env.GOOGLE_APPLICATION_CREDENTIALS = keyPath;
}

function getBucket() {
  const bucketName = (process.env.GCS_BUCKET_NAME || "").trim();
  if (!bucketName) {
    throw new Error(
      "ENV GCS_BUCKET_NAME is empty. Set it in Railway service Variables."
    );
  }
  ensureGcpCredentials();
  const storage = new Storage();
  return storage.bucket(bucketName);
}

module.exports = getBucket();
