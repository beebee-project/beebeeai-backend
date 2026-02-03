const RequestLog = require("../models/RequestLog");
const { getRecommendation } = require("../utils/reasonRecommendations");
const { getOrComputeDailySummary } = require("../services/dailySummaryService");

function percentile(sortedArr, p) {
  if (!Array.isArray(sortedArr) || sortedArr.length === 0) return null;
  const n = sortedArr.length;
  const idx = Math.floor(p * (n - 1));
  return sortedArr[Math.max(0, Math.min(n - 1, idx))];
}

function computePercentiles(nums) {
  const arr = (nums || [])
    .filter((x) => Number.isFinite(x))
    .sort((a, b) => a - b);
  return {
    p50: percentile(arr, 0.5),
    p95: percentile(arr, 0.95),
    min: arr.length ? arr[0] : null,
    max: arr.length ? arr[arr.length - 1] : null,
    n: arr.length,
  };
}

function extractTimingBuckets(docs) {
  // timingMs 객체에서 숫자 필드들을 모아 key별로 p50/p95 계산
  const buckets = {}; // { key: number[] }
  for (const d of docs || []) {
    const t = d?.debugMeta?.timingMs;
    if (!t || typeof t !== "object") continue;
    for (const [k, v] of Object.entries(t)) {
      if (!Number.isFinite(v)) continue;
      (buckets[k] ||= []).push(v);
    }
  }
  // 과도한 key 폭증 방지: 상위 12개 key만(샘플 내 등장 수 기준)
  const keys = Object.keys(buckets)
    .sort((a, b) => (buckets[b]?.length || 0) - (buckets[a]?.length || 0))
    .slice(0, 12);

  const out = {};
  for (const k of keys) out[k] = computePercentiles(buckets[k]);
  return out;
}

function clampWindowDays(windowRaw) {
  const n = Number(windowRaw || 7);
  if (n === 7 || n === 30) return n;
  return 7;
}

function safeTimezone(tz) {
  // Mongo timezone은 IANA string을 기대. 이상하면 기본값.
  return typeof tz === "string" && tz.length <= 64 ? tz : "Asia/Seoul";
}

function rangeFromWindowDays(windowDays) {
  const to = new Date();
  const from = new Date(to.getTime() - windowDays * 24 * 60 * 60 * 1000);
  return { from, to };
}

exports.getDailySummary = async (req, res) => {
  try {
    const { day, force, reasonTopN, validatorTopN } = req.query;

    // maxTimeMS는 운영에서 튜닝 가능 (일단 기본 1500ms)
    const result = await getOrComputeDailySummary({
      day,
      force,
      reasonTopN,
      validatorTopN,
      maxTimeMS: 1500,
    });

    return res.status(200).json(result);
  } catch (e) {
    console.error("[admin.getDailySummary]", e);
    return res.status(500).json({
      ok: false,
      error: {
        code: "DAILY_SUMMARY_FAILED",
        message: "DailySummary fetch failed",
      },
    });
  }
};

exports.getTrends = async (req, res) => {
  try {
    if (!RequestLog) {
      return res.status(500).json({
        error:
          "RequestLog model not found. Set mongoose.models.RequestLog or import your actual model.",
      });
    }

    const windowDays = clampWindowDays(req.query.window);
    const tz = safeTimezone(req.query.tz || "Asia/Seoul");
    const reasonTopN = Math.min(Number(req.query.reasonTopN || 10), 50);
    const validatorTopN = Math.min(Number(req.query.validatorTopN || 10), 50);
    const limit = Math.min(Number(req.query.limit || 20), 100);
    const maxTimeMS = Math.min(Number(req.query.maxTimeMS || 1500), 10_000);

    const range = rangeFromWindowDays(windowDays);
    const match = { createdAt: { $gte: range.from, $lte: range.to } };

    const [
      series,
      topReasons,
      engineDist,
      fallbackCount,
      validatorFailPointsTop,
      validatorKinds,
      recentFailures,
    ] = await Promise.all([
      // ✅ 일자별 KPI (timezone 반영)
      RequestLog.aggregate([
        { $match: match },
        {
          $group: {
            _id: {
              $dateToString: {
                format: "%Y-%m-%d",
                date: "$createdAt",
                timezone: tz,
              },
            },
            all: { $sum: 1 },
            success: {
              $sum: { $cond: [{ $eq: ["$status", "success"] }, 1, 0] },
            },
            fail: { $sum: { $cond: [{ $eq: ["$status", "fail"] }, 1, 0] } },
            fallback: {
              $sum: { $cond: [{ $eq: ["$isFallback", true] }, 1, 0] },
            },
          },
        },
        { $sort: { _id: 1 } },
      ]).option({ maxTimeMS }),

      // ✅ top reasons
      RequestLog.aggregate([
        { $match: match },
        { $group: { _id: "$reason", count: { $sum: 1 } } },
        { $sort: { count: -1 } },
        { $limit: reasonTopN },
      ]).option({ maxTimeMS }),

      // ✅ engine distribution
      RequestLog.aggregate([
        { $match: match },
        { $group: { _id: "$engine", count: { $sum: 1 } } },
        { $sort: { count: -1 } },
      ]).option({ maxTimeMS }),

      // ✅ fallback 총량(스키마에 isFallback 없을 수도 있으니 방어)
      RequestLog.countDocuments({ ...match, isFallback: true }).catch(() => 0),

      // ✅ validator fail points top
      RequestLog.aggregate([
        {
          $match: {
            ...match,
            "debugMeta.validatorFailPoints": { $exists: true, $ne: [] },
          },
        },
        { $unwind: "$debugMeta.validatorFailPoints" },
        {
          $group: { _id: "$debugMeta.validatorFailPoints", count: { $sum: 1 } },
        },
        { $sort: { count: -1 } },
        { $limit: validatorTopN },
      ]).option({ maxTimeMS }),

      // ✅ validator kind distribution
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

      // ✅ 최근 실패 샘플(운영자가 “왜 터지나” 바로 보게)
      RequestLog.find({ ...match, status: "fail" })
        .sort({ createdAt: -1 })
        .limit(Math.min(limit, 30))
        .select({
          createdAt: 1,
          prompt: 1,
          engine: 1,
          reason: 1,
          isFallback: 1,
          traceId: 1,
          userId: 1,
          debugMeta: 1,
        })
        .lean(),
    ]);

    const totalsAll = series.reduce((a, x) => a + (x.all || 0), 0);
    const totalsSuccess = series.reduce((a, x) => a + (x.success || 0), 0);
    const totalsFail = series.reduce((a, x) => a + (x.fail || 0), 0);

    const engineMap = (engineDist || []).reduce(
      (a, c) => ((a[c._id ?? "unknown"] = c.count), a),
      {},
    );
    const validatorKindMap = (validatorKinds || []).reduce(
      (a, c) => ((a[c._id ?? "unknown"] = c.count), a),
      {},
    );

    res.json({
      windowDays,
      tz,
      range: { from: range.from.toISOString(), to: range.to.toISOString() },
      totals: {
        all: totalsAll,
        success: totalsSuccess,
        fail: totalsFail,
        fallback: fallbackCount || 0,
      },
      series: (series || []).map((x) => ({
        day: x._id,
        all: x.all || 0,
        success: x.success || 0,
        fail: x.fail || 0,
        fallback: x.fallback || 0,
      })),
      distributions: {
        engine: engineMap,
        validatorKind: validatorKindMap,
      },
      topReasons: (topReasons || []).map((r) => ({
        reason: r._id ?? "unknown",
        count: r.count,
        recommendation: getRecommendation ? getRecommendation(r._id) : null,
      })),
      validator: {
        failPointsTop: (validatorFailPointsTop || []).map((x) => ({
          code: x._id ?? "unknown",
          count: x.count,
        })),
      },
      recent: {
        failures: (recentFailures || []).map((x) => ({
          at: x.createdAt,
          engine: x.engine,
          reason: x.reason,
          isFallback: x.isFallback,
          traceId: x.traceId || String(x._id),
          userId: x.userId,
          rawReason: x.debugMeta?.rawReason,
          validatorFailPoints: x.debugMeta?.validatorFailPoints,
          promptPreview: x.prompt ? String(x.prompt).slice(0, 160) : "",
        })),
      },
    });
  } catch (e) {
    console.error("[admin.getTrends]", e);
    res.status(500).json({ error: "Trends query failed" });
  }
};

function parseDateRange(query) {
  // 기본: 최근 7일 (KST 기준으로 딱 맞추려면 timezone 처리가 필요하지만, 1차는 UTC로 가자)
  const now = new Date();
  const to = query.to ? new Date(query.to) : now;
  const from = query.from
    ? new Date(query.from)
    : new Date(to.getTime() - 7 * 24 * 60 * 60 * 1000);

  // invalid 방어
  if (Number.isNaN(from.getTime()) || Number.isNaN(to.getTime())) {
    return null;
  }
  return { from, to };
}

// GET /admin/trace/:traceId
exports.getTraceDetail = async (req, res) => {
  try {
    const { traceId } = req.params;
    if (!traceId) {
      return res.status(400).json({ error: "traceId is required" });
    }

    const log = await RequestLog.findOne({ traceId }).lean();
    if (!log) {
      return res.status(404).json({ error: "Trace not found" });
    }

    res.json({
      traceId: log.traceId,
      createdAt: log.createdAt,
      route: log.route,
      engine: log.engine,

      status: log.status,
      reason: log.reason,
      rawReason: log.debugMeta?.rawReason,
      isFallback: log.isFallback,

      prompt: log.prompt,

      validator: {
        ok: log.debugMeta?.validatorOk,
        kind: log.debugMeta?.validatorKind,
        failPoints: log.debugMeta?.validatorFailPoints || [],
      },

      timingMs: log.debugMeta?.timingMs || {},

      cache: {
        hit: log.debugMeta?.cacheHit,
        intentOp: log.debugMeta?.intentOp,
        intentCacheKey: log.debugMeta?.intentCacheKey,
      },

      extra: Object.fromEntries(
        Object.entries(log.debugMeta || {}).filter(
          ([k]) =>
            ![
              "rawReason",
              "validatorOk",
              "validatorKind",
              "validatorFailPoints",
              "timingMs",
              "cacheHit",
              "intentOp",
              "intentCacheKey",
            ].includes(k),
        ),
      ),
    });
  } catch (e) {
    console.error("[getTraceDetail]", e);
    res.status(500).json({ error: "Trace lookup failed" });
  }
};

exports.getAdminSummary = async (req, res) => {
  try {
    if (!RequestLog) {
      return res.status(500).json({
        error:
          "RequestLog model not found. Set mongoose.models.RequestLog or import your actual model.",
      });
    }

    const range = parseDateRange(req.query);
    if (!range) return res.status(400).json({ error: "Invalid from/to" });

    const limit = Math.min(Number(req.query.limit || 20), 100);
    const reasonTopN = Math.min(Number(req.query.reasonTopN || 10), 50);
    const validatorTopN = Math.min(Number(req.query.validatorTopN || 10), 50);
    const validatorSampleN = Math.min(
      Number(req.query.validatorSampleN || 10),
      50,
    );
    const timingSampleN = Math.min(
      Number(req.query.timingSampleN || 300),
      2000,
    );
    const maxTimeMS = Math.min(Number(req.query.maxTimeMS || 1500), 10_000);

    const match = {
      createdAt: { $gte: range.from, $lte: range.to },
    };

    // ✅ RequestLog 스키마가 다르면 아래 필드명만 맞춰주면 됨:
    // - status: "success" | "fail" (또는 boolean)
    // - engine: "formula" | "officescripts" | "appscript" | "sql" ...
    // - reason: string
    // - isFallback: boolean (없으면 제거)
    // - traceId: string (없으면 _id로 대체)
    const [
      totals,
      byStatus,
      byEngine,
      reasonTop,
      recentFailures,
      recentSuccess,
      validatorFailPointsTop,
      validatorKinds,
      recentValidationFailures,
      cacheHits,
      intentOpDist,
      timingSamples,
    ] = await Promise.all([
      RequestLog.countDocuments(match),

      RequestLog.aggregate([
        { $match: match },
        {
          $group: {
            _id: "$status",
            count: { $sum: 1 },
          },
        },
      ]).option({ maxTimeMS }),

      RequestLog.aggregate([
        { $match: match },
        {
          $group: {
            _id: "$engine",
            count: { $sum: 1 },
          },
        },
        { $sort: { count: -1 } },
      ]).option({ maxTimeMS }),

      RequestLog.aggregate([
        { $match: match },
        {
          $group: {
            _id: "$reason",
            count: { $sum: 1 },
          },
        },
        { $sort: { count: -1 } },
        { $limit: reasonTopN },
      ]).option({ maxTimeMS }),

      RequestLog.find({
        ...match,
        status: "fail",
      })
        .sort({ createdAt: -1 })
        .limit(limit)
        .select({
          createdAt: 1,
          prompt: 1,
          engine: 1,
          reason: 1,
          isFallback: 1,
          traceId: 1,
          userId: 1,
          debugMeta: 1,
        })
        .lean(),

      RequestLog.find({
        ...match,
        status: "success",
      })
        .sort({ createdAt: -1 })
        .limit(Math.min(limit, 10))
        .select({
          createdAt: 1,
          engine: 1,
          traceId: 1,
          userId: 1,
          debugMeta: 1,
        })
        .lean(),

      // ✅ validatorFailPointsTop: debugMeta.validatorFailPoints[]를 unwind 후 topN 집계
      RequestLog.aggregate([
        {
          $match: {
            ...match,
            "debugMeta.validatorFailPoints": { $exists: true, $ne: [] },
          },
        },
        { $unwind: "$debugMeta.validatorFailPoints" },
        {
          $group: {
            _id: "$debugMeta.validatorFailPoints",
            count: { $sum: 1 },
          },
        },
        { $sort: { count: -1 } },
        { $limit: validatorTopN },
      ]).option({ maxTimeMS }),

      // ✅ validatorKinds: 어떤 검증기(kind)에서 많이 터지는지
      RequestLog.aggregate([
        {
          $match: {
            ...match,
            "debugMeta.validatorKind": { $exists: true, $ne: null },
          },
        },
        {
          $group: {
            _id: "$debugMeta.validatorKind",
            count: { $sum: 1 },
          },
        },
        { $sort: { count: -1 } },
      ]).option({ maxTimeMS }),

      // ✅ 최근 검증 실패 샘플
      RequestLog.find({
        ...match,
        "debugMeta.validatorOk": false,
      })
        .sort({ createdAt: -1 })
        .limit(validatorSampleN)
        .select({
          createdAt: 1,
          engine: 1,
          reason: 1,
          traceId: 1,
          prompt: 1,
          debugMeta: 1,
        })
        .lean(),

      // ✅ cacheHit count
      RequestLog.countDocuments({
        ...match,
        "debugMeta.cacheHit": true,
      }).catch(() => 0),

      // ✅ intentOp distribution
      RequestLog.aggregate([
        {
          $match: {
            ...match,
            "debugMeta.intentOp": { $exists: true, $ne: null },
          },
        },
        {
          $group: {
            _id: "$debugMeta.intentOp",
            count: { $sum: 1 },
          },
        },
        { $sort: { count: -1 } },
      ]).option({ maxTimeMS }),

      // ✅ timing 샘플 (p50/p95는 샘플 기반으로 계산)
      RequestLog.find(match)
        .sort({ createdAt: -1 })
        .limit(timingSampleN)
        .select({ createdAt: 1, "debugMeta.timingMs": 1 })
        .lean(),
    ]);

    // 상태 집계 정규화
    const statusMap = byStatus.reduce((acc, cur) => {
      acc[cur._id ?? "unknown"] = cur.count;
      return acc;
    }, {});
    const engineMap = byEngine.reduce((acc, cur) => {
      acc[cur._id ?? "unknown"] = cur.count;
      return acc;
    }, {});
    const reasonList = reasonTop.map((r) => ({
      reason: r._id ?? "unknown",
      count: r.count,
    }));

    const validatorTop = (validatorFailPointsTop || []).map((x) => ({
      code: x._id ?? "unknown",
      count: x.count,
    }));
    const validatorKindDist = (validatorKinds || []).reduce((acc, cur) => {
      acc[cur._id ?? "unknown"] = cur.count;
      return acc;
    }, {});

    // ✅ cache/intentOp 정리
    const cacheHitCount = Number(cacheHits || 0);
    const cacheMissCount = Math.max(0, Number(totals || 0) - cacheHitCount);
    const cacheHitRate = totals ? cacheHitCount / totals : 0;
    const intentOpMap = (intentOpDist || []).reduce((acc, cur) => {
      acc[cur._id ?? "unknown"] = cur.count;
      return acc;
    }, {});

    // ✅ timing p50/p95 (샘플 기반)
    const timingStats = extractTimingBuckets(timingSamples || []);

    // ✅ reason별 대표 실패 샘플 + 추천 액션
    const enrichedReasons = await Promise.all(
      reasonList.map(async ({ reason, count }) => {
        const samples = await RequestLog.find({
          ...match,
          status: "fail",
          reason,
        })
          .sort({ createdAt: -1 })
          .limit(3)
          .select({
            createdAt: 1,
            engine: 1,
            prompt: 1,
            traceId: 1,
            debugMeta: 1,
          })
          .lean();
        return {
          reason,
          count,
          recommendation: getRecommendation(reason),
          samples: samples.map((s) => ({
            at: s.createdAt,
            engine: s.engine,
            traceId: s.traceId || String(s._id),
            rawReason: s.debugMeta?.rawReason,
            promptPreview: s.prompt ? String(s.prompt).slice(0, 160) : "",
          })),
        };
      }),
    );

    // 기본 KPI
    const success = statusMap.success || 0;
    const fail = statusMap.fail || 0;
    const fallbackCount = await RequestLog.countDocuments({
      ...match,
      isFallback: true,
    }).catch(() => 0); // isFallback 없는 스키마면 0 처리

    res.json({
      range: {
        from: range.from.toISOString(),
        to: range.to.toISOString(),
      },
      totals: {
        all: totals,
        success,
        fail,
        fallback: fallbackCount,
      },
      distributions: {
        status: statusMap,
        engine: engineMap,
        validatorKind: validatorKindDist,
      },
      cache: {
        hit: cacheHitCount,
        miss: cacheMissCount,
        hitRate: cacheHitRate,
        intentOp: intentOpMap,
      },
      performance: {
        timingMs: timingStats,
        timingSampleN: (timingSamples || []).length,
      },
      reasonTop: enrichedReasons,
      validator: {
        failPointsTop: validatorTop,
        recentFailures: (recentValidationFailures || []).map((x) => ({
          at: x.createdAt,
          engine: x.engine,
          reason: x.reason,
          traceId: x.traceId || String(x._id),
          rawReason: x.debugMeta?.rawReason,
          validatorKind: x.debugMeta?.validatorKind,
          validatorFailPoints: x.debugMeta?.validatorFailPoints,
          promptPreview: x.prompt ? String(x.prompt).slice(0, 160) : "",
        })),
      },
      recent: {
        failures: recentFailures.map((x) => ({
          at: x.createdAt,
          engine: x.engine,
          reason: x.reason,
          isFallback: x.isFallback,
          traceId: x.traceId || String(x._id),
          promptPreview: x.prompt ? String(x.prompt).slice(0, 160) : "",
          userId: x.userId,
          rawReason: x.debugMeta?.rawReason,
          validatorFailPoints: x.debugMeta?.validatorFailPoints,
        })),
        success: recentSuccess.map((x) => ({
          at: x.createdAt,
          engine: x.engine,
          traceId: x.traceId || String(x._id),
          userId: x.userId,
        })),
      },
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Admin summary failed" });
  }
};
