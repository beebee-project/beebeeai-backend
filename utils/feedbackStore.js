// utils/feedbackStore.js
const { Storage } = require("@google-cloud/storage");
const crypto = require("crypto");

const storage = new Storage();

// 업로드용으로 쓰는 버킷을 그대로 사용해도 OK (prefix로 분리)
const BUCKET_NAME = process.env.GCS_BUCKET_NAME;
if (!BUCKET_NAME) throw new Error("Missing env: GCS_BUCKET_NAME");

// 선택: prefix를 바꾸고 싶으면 env로 조절 가능
const FEEDBACK_PREFIX = process.env.GCS_FEEDBACK_PREFIX || "feedback/events";

function ymd(date = new Date()) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

async function appendFeedback(event) {
  const now = new Date();
  const day = ymd(now);
  const ts = now.getTime();
  const rand = crypto.randomBytes(6).toString("hex");

  const payload = {
    ...event,
    ts: event.ts || now.toISOString(),
  };

  const key = `${FEEDBACK_PREFIX}/${day}/${ts}_${rand}.json`;
  const file = storage.bucket(BUCKET_NAME).file(key);

  await file.save(JSON.stringify(payload, null, 2), {
    contentType: "application/json; charset=utf-8",
    resumable: false,
    metadata: { cacheControl: "no-store" },
  });

  return { bucket: BUCKET_NAME, key };
}

module.exports = { appendFeedback };
