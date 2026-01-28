const RequestLog = require("../models/RequestLog");

function safeStr(s, max = 2000) {
  const v = String(s ?? "");
  return v.length > max ? v.slice(0, max) : v;
}

async function writeRequestLog({
  traceId,
  userId,
  route,
  engine,
  status,
  reason,
  isFallback,
  prompt,
  latencyMs,
  debugMeta,
}) {
  try {
    await RequestLog.create({
      traceId,
      userId,
      route,
      engine,
      status,
      reason,
      isFallback: !!isFallback,
      prompt: safeStr(prompt, 4000),
      latencyMs: typeof latencyMs === "number" ? latencyMs : undefined,
      debugMeta:
        debugMeta && typeof debugMeta === "object" ? debugMeta : undefined,
    });
  } catch (e) {
    // 로깅 실패는 사용자 응답을 망치면 안 됨
    console.error("[writeRequestLog] failed:", e.message);
  }
}

module.exports = { writeRequestLog };
