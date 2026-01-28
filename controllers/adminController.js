const RequestLog = require("../models/RequestLog");

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
      },
      reasonTop: reasonList,
      recent: {
        failures: recentFailures.map((x) => ({
          at: x.createdAt,
          engine: x.engine,
          reason: x.reason,
          isFallback: x.isFallback,
          traceId: x.traceId || String(x._id),
          promptPreview: x.prompt ? String(x.prompt).slice(0, 160) : "",
          userId: x.userId,
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
