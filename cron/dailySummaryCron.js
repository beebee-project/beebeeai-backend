const cron = require("node-cron");
const { computeDailySummary } = require("../services/dailySummaryService");

function formatKstYesterday() {
  const now = new Date();
  const kstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const y = kstNow.getUTCFullYear();
  const m = String(kstNow.getUTCMonth() + 1).padStart(2, "0");
  const d = String(kstNow.getUTCDate() - 1).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function startDailySummaryCron() {
  // UTC 18:10 == KST 03:10
  cron.schedule("10 18 * * *", async () => {
    try {
      const day = formatKstYesterday();
      console.log("[cron] daily summary start", day);
      await computeDailySummary(day);
      console.log("[cron] daily summary done", day);
    } catch (e) {
      console.error("[cron] daily summary failed", e);
    }
  });
}

module.exports = { startDailySummaryCron };
