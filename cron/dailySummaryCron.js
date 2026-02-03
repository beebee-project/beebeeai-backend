console.log("[cron-entry] dailySummaryCron loaded");
const cron = require("node-cron");
const mongoose = require("mongoose");
const { computeDailySummary } = require("../services/dailySummaryService");

/**
 * KST 기준 "어제" YYYY-MM-DD
 */
function formatKstYesterday() {
  const now = new Date();
  const kstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const y = kstNow.getUTCFullYear();
  const m = String(kstNow.getUTCMonth() + 1).padStart(2, "0");
  const d = String(kstNow.getUTCDate() - 1).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/**
 * Mongo 기반 락(중복 실행 방지)
 * - 같은 day에 대해 동시에 2번 돌지 않게 막는다.
 * - TTL(만료)로 영구 락 방지
 */
const CronLockSchema = new mongoose.Schema(
  {
    key: { type: String, required: true, unique: true },
    expiresAt: { type: Date, required: true },
  },
  { timestamps: true },
);
CronLockSchema.index({ expiresAt: 1 });
const CronLock =
  mongoose.models.CronLock || mongoose.model("CronLock", CronLockSchema);

async function acquireMongoLock(key, ttlSec = 60 * 60) {
  const now = new Date();
  const nextExpiry = new Date(now.getTime() + ttlSec * 1000);

  // 핵심: 만료된 락이거나 없으면 내가 획득
  // - expiresAt이 now보다 과거면 갱신하며 획득
  // - 없으면 생성하며 획득
  const doc = await CronLock.findOneAndUpdate(
    {
      key,
      $or: [{ expiresAt: { $lte: now } }, { expiresAt: { $exists: false } }],
    },
    { $set: { expiresAt: nextExpiry } },
    { upsert: true, new: true },
  );

  // findOneAndUpdate가 upsert로 생성/갱신되면 doc는 항상 존재
  // 다만 경쟁 상황에서 unique 충돌이 날 수 있으니 호출부에서 catch 처리
  return !!doc;
}

/**
 * 1회 실행용 (Railway Cron에서 사용)
 */
async function runDailySummaryOnce({
  day,
  useLock = true,
  lockTtlSec = 60 * 60,
} = {}) {
  const targetDay = day || formatKstYesterday();
  const lockKey = `dailySummary:${targetDay}`;

  if (useLock) {
    try {
      const ok = await acquireMongoLock(lockKey, lockTtlSec);
      if (!ok) {
        console.log("[cron] lock not acquired, skip:", lockKey);
        return { ok: true, skipped: true, day: targetDay };
      }
    } catch (e) {
      // 경쟁으로 unique 충돌 등 → 이미 누가 락을 잡은 것과 동일 취급
      console.log("[cron] lock contention, skip:", lockKey);
      return { ok: true, skipped: true, day: targetDay };
    }
  }

  console.log("[cron] daily summary start", targetDay);
  const result = await computeDailySummary(targetDay);
  console.log("[cron] daily summary done", targetDay);
  return { ok: true, day: targetDay, result };
}

/**
 * (선택) 내부 node-cron 스케줄러
 * - 웹 서버에서 기본 OFF 권장
 * - 켜야 한다면 RUN_INTERNAL_CRON=1 같은 플래그로만 사용
 */
function startDailySummaryCron() {
  // UTC 18:10 == KST 03:10
  cron.schedule("10 18 * * *", async () => {
    try {
      await runDailySummaryOnce({ useLock: true });
    } catch (e) {
      console.error("[cron] daily summary failed", e);
    }
  });
}

module.exports = {
  formatKstYesterday,
  runDailySummaryOnce,
  startDailySummaryCron,
};
