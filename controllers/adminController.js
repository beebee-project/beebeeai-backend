const RequestLog = require("../models/RequestLog");
const { getRecommendation } = require("../utils/reasonRecommendations");
const DailySummary = require("../models/DailySummary");

function isDayString(s) {
  return typeof s === "string" && /^\d{4}-\d{2}-\d{2}$/.test(s);
}

function toKstDayString(d) {
  // Date -> KST 기준 YYYY-MM-DD
  const kst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
  const y = kst.getUTCFullYear();
  const m = String(kst.getUTCMonth() + 1).padStart(2, "0");
  const day = String(kst.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addDays(dayStr, delta) {
  // dayStr: YYYY-MM-DD (UTC로 파싱해도 날짜 계산은 안정적)
  const dt = new Date(`${dayStr}T00:00:00.000Z`);
  dt.setUTCDate(dt.getUTCDate() + delta);
  return toKstDayString(new Date(dt.getTime() - 9 * 60 * 60 * 1000)); // 역보정 후 KST stringify
}

function parseDayRange(query, defaultDays = 14) {
  // 기본: 최근 defaultDays일 (KST 기준)
  const todayKst = toKstDayString(new Date());
  const to = isDayString(query.to) ? query.to : addDays(todayKst, -1); // 기본: 어제
  const from = isDayString(query.from)
    ? query.from
    : addDays(to, -(defaultDays - 1));

  if (!isDayString(from) || !isDayString(to)) return null;
  if (from > to) return null;
  return { from, to };
}

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
      ]),

      RequestLog.aggregate([
        { $match: match },
        {
          $group: {
            _id: "$engine",
            count: { $sum: 1 },
          },
        },
        { $sort: { count: -1 } },
      ]),

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
      ]),

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
      ]),

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
      ]),

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

// ✅ 일별 DailySummary 리스트
exports.getDailySummaries = async (req, res) => {
  try {
    const range = parseDayRange(req.query, 14);
    if (!range) return res.status(400).json({ error: "Invalid from/to" });

    const limit = Math.min(Number(req.query.limit || 60), 365);

    const items = await DailySummary.find({
      day: { $gte: range.from, $lte: range.to },
    })
      .sort({ day: -1 })
      .limit(limit)
      .lean();

    res.json({
      ok: true,
      range,
      count: items.length,
      items,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Daily summaries fetch failed" });
  }
};

// ✅ 특정 day 한 건
exports.getDailySummaryByDay = async (req, res) => {
  try {
    const { day } = req.params;
    if (!isDayString(day))
      return res.status(400).json({ error: "Invalid day" });

    const doc = await DailySummary.findOne({ day }).lean();
    if (!doc) return res.status(404).json({ error: "Not found" });

    res.json({ ok: true, day, summary: doc });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Daily summary fetch failed" });
  }
};

// ✅ 최근 N일 트렌드(그래프용 시계열 + top stats)
exports.getDailyTrends = async (req, res) => {
  try {
    const days = Math.min(Math.max(Number(req.query.days || 30), 1), 365);
    const range = parseDayRange({ to: req.query.to }, days);
    if (!range) return res.status(400).json({ error: "Invalid range" });

    const docs = await DailySummary.find({
      day: { $gte: range.from, $lte: range.to },
    })
      .sort({ day: 1 })
      .lean();

    // 시계열 (없는 날은 그냥 누락: 프론트에서 gap 처리하거나 여기서 fill 가능)
    const series = docs.map((d) => {
      const totals = d.totals || {};
      const all = Number(totals.all || 0);
      const success = Number(totals.success || 0);
      const fail = Number(totals.fail || 0);
      const fallback = Number(totals.fallback || 0);
      const successRate = all > 0 ? success / all : 0;
      const failRate = all > 0 ? fail / all : 0;
      return {
        day: d.day,
        all,
        success,
        fail,
        fallback,
        successRate,
        failRate,
      };
    });

    // reasonTop / validator.failPointsTop 를 range 전체로 합산 (문서 구조 유연 대응)
    const reasonAgg = new Map();
    const failPointAgg = new Map();

    for (const d of docs) {
      const reasons = Array.isArray(d.reasonTop) ? d.reasonTop : [];
      for (const r of reasons) {
        const key = r?.reason ?? r?._id ?? r?.name ?? null;
        const cnt = Number(r?.count || 0);
        if (!key || !cnt) continue;
        reasonAgg.set(key, (reasonAgg.get(key) || 0) + cnt);
      }

      const fp = d?.validator?.failPointsTop;
      const fps = Array.isArray(fp) ? fp : [];
      for (const x of fps) {
        const key = x?.point ?? x?._id ?? x?.name ?? null;
        const cnt = Number(x?.count || 0);
        if (!key || !cnt) continue;
        failPointAgg.set(key, (failPointAgg.get(key) || 0) + cnt);
      }
    }

    const topReasons = [...reasonAgg.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([reason, count]) => ({ reason, count }));

    const topFailPoints = [...failPointAgg.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([point, count]) => ({ point, count }));

    res.json({
      ok: true,
      range,
      days,
      series,
      top: {
        reasons: topReasons,
        failPoints: topFailPoints,
      },
      count: docs.length,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Daily trends fetch failed" });
  }
};
