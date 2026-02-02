const { computeDailySummary } = require("../services/dailySummaryService");

function formatKstDay(date = new Date()) {
  // "어제(KST)" 기본 실행
  const kst = new Date(date.getTime() + 9 * 3600 * 1000);
  const y = kst.getUTCFullYear();
  const m = String(kst.getUTCMonth() + 1).padStart(2, "0");
  const d = String(kst.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

exports.runDailySummary = async (req, res) => {
  try {
    // date 미지정 시 “어제(KST)”를 요약
    const q = String(req.query.date || "").trim();
    const day =
      q ||
      (() => {
        const now = new Date();
        const yesterday = new Date(now.getTime() - 24 * 3600 * 1000);
        return formatKstDay(yesterday);
      })();

    const result = await computeDailySummary(day, {
      reasonTopN: req.query.reasonTopN,
      validatorTopN: req.query.validatorTopN,
    });

    res.json({ ok: true, day, summary: result });
  } catch (e) {
    console.error("[runDailySummary]", e);
    res.status(500).json({ ok: false, error: "Daily summary cron failed" });
  }
};
