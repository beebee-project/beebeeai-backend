const RequestLog = require("../models/RequestLog");
const DailySummary = require("../models/DailySummary");

function kstDayRange(dayStr) {
  // dayStr: "YYYY-MM-DD" (KST 기준)
  // KST = UTC+9 → from/to를 UTC로 변환
  const [y, m, d] = dayStr.split("-").map(Number);
  const fromKst = new Date(Date.UTC(y, m - 1, d, 0, 0, 0));
  const toKst = new Date(Date.UTC(y, m - 1, d + 1, 0, 0, 0));
  // 위는 UTC 기준 날짜 객체이므로, “KST 하루”를 만들려면 9시간을 빼서 UTC range로 맞춤
  const from = new Date(fromKst.getTime() - 9 * 3600 * 1000);
  const to = new Date(toKst.getTime() - 9 * 3600 * 1000);
  return { from, to };
}

async function computeDailySummary(dayStr, opts = {}) {
  const reasonTopN = Math.min(Number(opts.reasonTopN || 10), 50);
  const validatorTopN = Math.min(Number(opts.validatorTopN || 10), 50);

  const { from, to } = kstDayRange(dayStr);
  const match = { createdAt: { $gte: from, $lt: to } };

  const [
    totals,
    byStatus,
    byEngine,
    reasonTop,
    fallbackCount,
    validatorFailPointsTop,
    validatorKinds,
  ] = await Promise.all([
    RequestLog.countDocuments(match),

    RequestLog.aggregate([
      { $match: match },
      { $group: { _id: "$status", count: { $sum: 1 } } },
    ]),

    RequestLog.aggregate([
      { $match: match },
      { $group: { _id: "$engine", count: { $sum: 1 } } },
      { $sort: { count: -1 } },
    ]),

    RequestLog.aggregate([
      { $match: match },
      { $group: { _id: "$reason", count: { $sum: 1 } } },
      { $sort: { count: -1 } },
      { $limit: reasonTopN },
    ]),

    RequestLog.countDocuments({ ...match, isFallback: true }).catch(() => 0),

    RequestLog.aggregate([
      {
        $match: {
          ...match,
          "debugMeta.validatorFailPoints": { $exists: true, $ne: [] },
        },
      },
      { $unwind: "$debugMeta.validatorFailPoints" },
      { $group: { _id: "$debugMeta.validatorFailPoints", count: { $sum: 1 } } },
      { $sort: { count: -1 } },
      { $limit: validatorTopN },
    ]),

    RequestLog.aggregate([
      {
        $match: {
          ...match,
          "debugMeta.validatorKind": { $exists: true, $ne: null },
        },
      },
      { $group: { _id: "$debugMeta.validatorKind", count: { $sum: 1 } } },
      { $sort: { count: -1 } },
    ]),
  ]);

  const statusMap = byStatus.reduce(
    (a, c) => ((a[c._id ?? "unknown"] = c.count), a),
    {},
  );
  const engineMap = byEngine.reduce(
    (a, c) => ((a[c._id ?? "unknown"] = c.count), a),
    {},
  );
  const validatorKindDist = validatorKinds.reduce(
    (a, c) => ((a[c._id ?? "unknown"] = c.count), a),
    {},
  );

  const payload = {
    day: dayStr,
    range: { from, to },
    totals: {
      all: totals,
      success: statusMap.success || 0,
      fail: statusMap.fail || 0,
      fallback: fallbackCount || 0,
    },
    distributions: {
      status: statusMap,
      engine: engineMap,
      validatorKind: validatorKindDist,
    },
    reasonTop: reasonTop.map((r) => ({
      reason: r._id ?? "unknown",
      count: r.count,
    })),
    validator: {
      failPointsTop: (validatorFailPointsTop || []).map((x) => ({
        code: x._id ?? "unknown",
        count: x.count,
      })),
    },
  };

  await DailySummary.updateOne(
    { day: dayStr },
    { $set: payload },
    { upsert: true },
  );
  return payload;
}

module.exports = { computeDailySummary };
