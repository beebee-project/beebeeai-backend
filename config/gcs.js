// config/gcs.js
const fs = require("fs");
const path = require("path");
const { Storage } = require("@google-cloud/storage");

function ensureGcpCredentials() {
  const json = process.env.GCS_KEY;
  if (!json) return; // GOOGLE_APPLICATION_CREDENTIALS를 이미 쓰면 생략
  const keyPath = path.join("/app", ".gcp-key.json");
  const content = typeof json === "string" ? json : JSON.stringify(json);
  fs.writeFileSync(keyPath, content);
  process.env.GOOGLE_APPLICATION_CREDENTIALS = keyPath;
}

let _bucket = null;

function getBucket() {
  if (_bucket) return _bucket;

  const name = (process.env.GCS_BUCKET_NAME || "").trim();
  if (!name) {
    // 어떤 키들이 들어왔는지 보려면 아래 라인 잠깐 활성화:
    // console.log('ENV keys:', Object.keys(process.env).filter(k => k.startsWith('GCS')));
    throw new Error(
      "ENV GCS_BUCKET_NAME is empty. Set it in Railway service Variables."
    );
  }

  ensureGcpCredentials();
  const storage = new Storage();
  _bucket = storage.bucket(name);
  return _bucket;
}

module.exports = { getBucket };
