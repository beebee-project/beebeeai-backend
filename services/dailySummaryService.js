const RequestLog = require("../models/RequestLog");
const DailySummary = require("../models/DailySummary");

function isValidDayStr(dayStr) {
  // YYYY-MM-DD
  return typeof dayStr === "string" && /^\d{4}-\d{2}-\d{2}$/.test(dayStr);
}

function todayKstDayStr() {
  // 현재 시간을 KST로 옮긴 뒤 YYYY-MM-DD로 자름
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 3600 * 1000);
  return kst.toISOString().slice(0, 10);
}

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
  const maxTimeMS = Math.min(Number(opts.maxTimeMS || 1500), 10_000);

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
    ]).option({ maxTimeMS }),

    RequestLog.aggregate([
      { $match: match },
      { $group: { _id: "$engine", count: { $sum: 1 } } },
      { $sort: { count: -1 } },
    ]).option({ maxTimeMS }),

    RequestLog.aggregate([
      { $match: match },
      { $group: { _id: "$reason", count: { $sum: 1 } } },
      { $sort: { count: -1 } },
      { $limit: reasonTopN },
    ]).option({ maxTimeMS }),

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
    ]).option({ maxTimeMS }),

    RequestLog.aggregate([
      {
        $match: {
          ...match,
          "debugMeta.validatorKind": { $exists: true, $ne: null },
        },
      },
      { $group: { _id: "$debugMeta.validatorKind", count: { $sum: 1 } } },
      { $sort: { count: -1 } },
    ]).option({ maxTimeMS }),
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

/**
 * 조회만: 이미 만들어진 요약이 있으면 그대로 반환
 */
async function getDailySummary(dayStr) {
  if (!isValidDayStr(dayStr)) return null;
  return DailySummary.findOne({ day: dayStr }).lean();
}

/**
 * 운영용: (1) day 기본값=오늘(KST) (2) 조회 우선 (3) 없으면 계산 (4) force=1이면 재계산
 */
async function getOrComputeDailySummary(opts = {}) {
  const dayStr = isValidDayStr(opts.day) ? opts.day : todayKstDayStr();
  const force = String(opts.force || "0") === "1";

  if (!force) {
    const existing = await getDailySummary(dayStr);
    if (existing)
      return { ok: true, day: dayStr, cached: true, data: existing };
  }

  try {
    const computed = await computeDailySummary(dayStr, opts);
    return { ok: true, day: dayStr, cached: false, data: computed };
  } catch (e) {
    // 500 방지: 기존 데이터라도 있으면 그걸 반환(스테일 허용)
    const stale = await getDailySummary(dayStr);
    if (stale) {
      return {
        ok: true,
        day: dayStr,
        cached: true,
        partial: true,
        fallbackUsed: true,
        note: "compute failed; returned stale summary",
        data: stale,
      };
    }
    // 그래도 없으면 최소 형태로 내려줌(운영 안정성 우선)
    return {
      ok: true,
      day: dayStr,
      partial: true,
      fallbackUsed: true,
      note: "compute failed; returned empty summary",
      data: {
        day: dayStr,
        range: null,
        totals: { all: 0, success: 0, fail: 0, fallback: 0 },
        distributions: { status: {}, engine: {}, validatorKind: {} },
        reasonTop: [],
        validator: { failPointsTop: [] },
      },
    };
  }
}

module.exports = {
  computeDailySummary,
  getDailySummary,
  getOrComputeDailySummary,
};
