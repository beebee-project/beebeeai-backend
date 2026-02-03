/**
 * Railway Cron "Start Command"로 쓰는 엔트리 파일
 * 예) node cron/railwayDailySummaryEntry.js
 *
 * 동작:
 * 1) DB 연결
 * 2) daily summary 1회 실행
 * 3) 정상 종료(0) / 실패(1)
 */
const mongoose = require("mongoose");
const connectDB = require("../config/db");
const { runDailySummaryOnce } = require("./dailySummaryCron");

function parseDayArg() {
  // 지원:
  // - env: CRON_DAY=YYYY-MM-DD
  // - argv: --day=YYYY-MM-DD 또는 첫번째 인자 YYYY-MM-DD
  const envDay = process.env.CRON_DAY;
  if (envDay) return envDay;

  const dayFlag = process.argv.find((v) => v.startsWith("--day="));
  if (dayFlag) return dayFlag.split("=")[1];

  const raw = process.argv[2];
  if (raw && /^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return undefined;
}

(async () => {
  try {
    connectDB();
    const day = parseDayArg();
    const out = await runDailySummaryOnce({
      day,
      useLock: true,
      lockTtlSec: 60 * 60, // 1시간
    });

    console.log("[cron-entry] result:", out);
    await mongoose.connection.close();
    process.exit(0);
  } catch (e) {
    console.error("[cron-entry] failed:", e);
    try {
      await mongoose.connection.close();
    } catch (_) {}
    process.exit(1);
  }
})();
