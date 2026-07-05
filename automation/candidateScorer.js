const CANDIDATE_SCORING_VERSION = "candidate_scoring_v1_1";

const TOP_TIER_MAX = 7;
const SECONDARY_TIER_MAX = 30;

let multiSourceCandidateBuilder = null;
try {
  multiSourceCandidateBuilder = require("./multiSourceCandidateBuilder");
} catch (_) {
  multiSourceCandidateBuilder = null;
}

const MULTI_SOURCE_FINAL_INJECTION_VERSION =
  "multi_source_final_bundle_injection_v1";

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const result = [];

  for (const value of values) {
    const text = String(value ?? "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    result.push(text);
  }

  return result;
}

function clampNumber(value, fallback = 0, min = 0, max = 100) {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

function normalizeText(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_\-()\[\]{}.,:/\\]/g, "");
}

function normalizeType(candidate = {}) {
  return String(
    candidate.recipeType ||
      candidate.type ||
      candidate.recipeId ||
      candidate.candidateType ||
      "",
  ).trim();
}

function hasAnyToken(value = "", tokens = []) {
  const text = String(value || "").toLowerCase();
  return tokens.some((token) => text.includes(String(token).toLowerCase()));
}

function hasAnyNormalizedToken(value = "", tokens = []) {
  const text = normalizeText(value);
  return tokens.some((token) => text.includes(normalizeText(token)));
}

function collectCandidateText(candidate = {}) {
  return [
    candidate.title,
    candidate.description,
    candidate.recommendationReason,
    candidate.recipeType,
    candidate.type,
    candidate.recipeId,
    candidate.templateId,
    candidate.sourceSheetName,
    candidate.sourceTableId,
    candidate.candidateId,
    candidate.id,
    candidate.columns?.dimension,
    candidate.columns?.metric,
    candidate.columns?.period,
    candidate.columns?.value,
    ...(candidate.sourceSheetNames || []),
    ...(candidate.sourceTableIds || []),
    ...(candidate.reasonCodes || []),
    ...(candidate.recipeIds || []),
  ]
    .filter(Boolean)
    .join(" ");
}

function collectMetricText(candidate = {}) {
  const columns = candidate.columns || {};
  return [
    columns.metric,
    columns.value,
    columns.amount,
    columns.measure,
    candidate.metricHeader,
    candidate.valueHeader,
    candidate.amountHeader,
    candidate.measureHeader,
    candidate.metricName,
    candidate.valueField,
    candidate.chartHint?.valueField,
    candidate.title,
    candidate.candidateId,
    candidate.recipeId,
    candidate.recipeType,
  ]
    .filter(Boolean)
    .join(" ");
}

function recipeTypeScore(candidate = {}) {
  const type = normalizeType(candidate);
  const recipeIds = asArray(candidate.recipeIds).join(" ");
  const text = `${type} ${recipeIds}`;

  if (/businessTemplate/i.test(candidate.candidateType || "")) return 18;
  if (/multiSource/i.test(candidate.candidateType || "")) return 16;
  if (/dashboard/i.test(candidate.candidateType || "")) return 14;
  if (/time_(sum|avg|growth|trend)|cumulative_sum|wide_time_trend/i.test(text))
    return 13;
  if (/group_(sum|avg|count|summary)|composition_ratio|top_bottom/i.test(text))
    return 11;
  if (/cross_(count|sum)|status_count|category_count/i.test(text)) return 9;
  return 6;
}

function outputTypeScore(candidate = {}) {
  const outputTypes = asArray(candidate.outputTypes);
  if (!outputTypes.length) return 0;
  let score = 0;
  if (outputTypes.includes("summarySheet")) score += 3;
  if (outputTypes.includes("analysisReport")) score += 2;
  if (outputTypes.includes("ppt")) score += 2;
  return Math.min(6, score);
}

function sourceScore(candidate = {}, context = {}) {
  const policy = context.sourceTablePolicy || {};
  const primaryTableId = policy.primarySourceTableId || "";
  const primarySheetName = policy.primarySourceSheetName || "";
  const sourceTableIds = asArray(candidate.sourceTableIds);
  const sourceSheetNames = asArray(candidate.sourceSheetNames);
  const sourceTableId = candidate.sourceTableId || sourceTableIds[0] || "";
  const sourceSheetName =
    candidate.sourceSheetName || sourceSheetNames[0] || "";
  const scope = String(candidate.sourceScope || "").trim();

  let score = 0;
  if (sourceTableId) score += 5;
  if (sourceSheetName) score += 4;
  if (primaryTableId && sourceTableIds.includes(primaryTableId)) score += 3;
  if (primarySheetName && sourceSheetNames.includes(primarySheetName))
    score += 2;
  if (scope === "multiTable") score += 3;
  if (scope === "virtualLinkedTable") score += 3;
  if (sourceTableIds.length > 1) score += 3;

  return Math.min(16, score);
}

function evidenceScore(candidate = {}) {
  const text = collectCandidateText(candidate);
  const matchedCount = Number(candidate.matchedCount || 0);
  let score = 0;

  if (matchedCount > 0) score += Math.min(7, matchedCount * 1.5);
  if (asArray(candidate.matchedHeaders).length)
    score += Math.min(5, asArray(candidate.matchedHeaders).length * 1.5);
  if (asArray(candidate.reasonCodes).length)
    score += Math.min(4, asArray(candidate.reasonCodes).length * 0.8);
  if (candidate.primaryCandidate) score += 3;
  if (asArray(candidate.candidates).length >= 2) score += 3;

  if (hasAnyToken(text, ["월", "연도", "기간", "추이", "growth", "trend"]))
    score += 3;
  if (
    hasAnyToken(text, ["합계", "금액", "매출", "집행", "연봉", "sum", "amount"])
  )
    score += 3;
  if (hasAnyToken(text, ["평균", "avg", "average"])) score += 2;
  if (hasAnyToken(text, ["비중", "구성", "ratio", "composition"])) score += 2;
  if (hasAnyToken(text, ["상위", "하위", "top", "bottom"])) score += 2;

  return Math.min(18, score);
}

const SALES_CONTEXT_TOKENS = [
  "sales_report",
  "매출",
  "판매",
  "순매출",
  "카드매출",
  "거래액",
  "출하액",
  "신용카드",
  "면세점",
];

const SALES_AMOUNT_TOKENS = [
  "순매출액",
  "매출액",
  "판매액",
  "판매금액",
  "거래액",
  "출하액",
  "금액",
  "amount",
  "salesamount",
  "revenue",
];

const SALES_QUANTITY_TOKENS = [
  "매출수량",
  "판매수량",
  "수량",
  "건수",
  "count",
  "quantity",
  "qty",
];

const RESEARCH_CONTEXT_TOKENS = [
  "research_budget_report",
  "연구비",
  "연구개발비",
  "집행금액",
  "정부출연금",
  "민간부담금",
  "과제",
  "rnd",
  "r&d",
  "budget",
  "grant",
];

const RESEARCH_AMOUNT_TOKENS = [
  "총연구비",
  "총 연구비",
  "연구비",
  "정부출연금",
  "민간부담금",
  "집행금액",
  "집행액",
  "현금",
  "현물",
  "금액",
  "budget",
  "grant",
  "expense",
];

const YEAR_LIKE_METRIC_TOKENS = [
  "예산년도",
  "진행년도",
  "집행년도",
  "기준년도",
  "연도구분",
  "연도 구분",
  "연도",
  "년도",
  "year",
];

function hasSalesContext(candidate = {}) {
  return hasAnyNormalizedToken(
    collectCandidateText(candidate),
    SALES_CONTEXT_TOKENS,
  );
}

function hasResearchContext(candidate = {}) {
  return hasAnyNormalizedToken(
    collectCandidateText(candidate),
    RESEARCH_CONTEXT_TOKENS,
  );
}

function isSalesAmountMetric(candidate = {}) {
  return hasAnyNormalizedToken(
    collectMetricText(candidate),
    SALES_AMOUNT_TOKENS,
  );
}

function isSalesQuantityMetric(candidate = {}) {
  return hasAnyNormalizedToken(
    collectMetricText(candidate),
    SALES_QUANTITY_TOKENS,
  );
}

function isResearchAmountMetric(candidate = {}) {
  return hasAnyNormalizedToken(
    collectMetricText(candidate),
    RESEARCH_AMOUNT_TOKENS,
  );
}

function isYearLikeMetric(candidate = {}) {
  const metricText = collectMetricText(candidate);
  const title = String(candidate.title || "");
  const id = String(candidate.candidateId || candidate.id || "");

  if (hasAnyNormalizedToken(metricText, YEAR_LIKE_METRIC_TOKENS)) return true;

  return (
    /(?:기준|별)\s*(예산년도|진행년도|집행년도|기준년도|연도|년도)/.test(
      title,
    ) || /_(예산년도|진행년도|집행년도|기준년도|연도|년도)(?:_|$)/.test(id)
  );
}

function isGenericIndicatorNameMetric(candidate = {}) {
  const metricText = collectMetricText(candidate);
  return hasAnyNormalizedToken(metricText, [
    "지표명",
    "metricname",
    "indicatorname",
  ]);
}

function isGenericIndicatorValueMetric(candidate = {}) {
  const metricText = collectMetricText(candidate);
  return hasAnyNormalizedToken(metricText, [
    "지표값",
    "metricvalue",
    "indicatorvalue",
  ]);
}

function isVirtualCandidate(candidate = {}) {
  const ids = [
    candidate.sourceTableId,
    ...(candidate.sourceTableIds || []),
    candidate.tableId,
    candidate.candidateId,
  ].join(" ");
  return /#(?:WIDE_LONG|CROSS_LONG)(?:_|\b|$)/i.test(ids);
}

function domainAdjustmentScore(candidate = {}) {
  let score = 0;
  const signals = [];
  const penalties = [];
  const salesContext = hasSalesContext(candidate);
  const researchContext = hasResearchContext(candidate);
  const salesAmount = isSalesAmountMetric(candidate);
  const salesQuantity = isSalesQuantityMetric(candidate);
  const researchAmount = isResearchAmountMetric(candidate);
  const yearLikeMetric = isYearLikeMetric(candidate);
  const genericIndicatorName = isGenericIndicatorNameMetric(candidate);
  const genericIndicatorValue = isGenericIndicatorValueMetric(candidate);

  if (/businessTemplate/i.test(candidate.candidateType || "")) {
    score += 5;
    signals.push("BUSINESS_TEMPLATE_PRIORITY");
  }

  if (/multiSource/i.test(candidate.candidateType || "")) {
    score += 3;
    signals.push("MULTI_SOURCE_STRUCTURE_PRIORITY");
  }

  if (salesContext && salesAmount) {
    score += 10;
    signals.push("SALES_AMOUNT_METRIC_BOOST");
  }

  if (salesContext && salesQuantity && !salesAmount) {
    score -= 12;
    penalties.push("SALES_QUANTITY_METRIC_PENALTY");
  }

  if (researchContext && researchAmount && !yearLikeMetric) {
    score += 10;
    signals.push("RESEARCH_AMOUNT_METRIC_BOOST");
  }

  if (researchContext && yearLikeMetric && !researchAmount) {
    score -= 20;
    penalties.push("YEAR_LIKE_METRIC_PENALTY");
  }

  if (genericIndicatorName) {
    score -= 12;
    penalties.push("GENERIC_INDICATOR_NAME_PENALTY");
  }

  if (genericIndicatorValue && !isVirtualCandidate(candidate)) {
    score -= 5;
    penalties.push("GENERIC_INDICATOR_VALUE_PENALTY");
  }

  if (isVirtualCandidate(candidate) && genericIndicatorValue) {
    score += 2;
    signals.push("VIRTUAL_INDICATOR_VALUE_ACCEPTED");
  }

  return {
    score: Math.max(-28, Math.min(18, score)),
    signals,
    penalties,
    traits: {
      salesContext,
      salesAmount,
      salesQuantity,
      researchContext,
      researchAmount,
      yearLikeMetric,
      genericIndicatorName,
      genericIndicatorValue,
      virtualCandidate: isVirtualCandidate(candidate),
    },
  };
}

function penaltyScore(candidate = {}) {
  let penalty = 0;
  const text = collectCandidateText(candidate);

  if (!candidate.title) penalty += 5;
  if (!asArray(candidate.sourceTableIds).length && !candidate.sourceTableId)
    penalty += 10;
  if (!asArray(candidate.sourceSheetNames).length && !candidate.sourceSheetName)
    penalty += 8;
  if (
    !asArray(candidate.recipeIds).length &&
    !candidate.recipeId &&
    !candidate.recipeType
  )
    penalty += 5;
  if (hasAnyToken(text, ["메타", "주석", "출처", "설명", "note", "metadata"]))
    penalty += 4;
  if (candidate.tableUsage?.analysisEligible === false) penalty += 25;

  return penalty;
}

function candidateTypeSortRank(candidate = {}) {
  const type = String(candidate.candidateType || "");
  if (/businessTemplate/i.test(type)) return 0;
  if (/multiSource/i.test(type)) return 1;
  if (/dashboard/i.test(type)) return 2;
  if (/automationCategory|category/i.test(type)) return 3;
  if (/analysisRecipe/i.test(type)) return 4;
  return 9;
}

function domainSortScore(candidate = {}) {
  return (
    Number(candidate.score?.signals?.domainFit || 0) -
    Number(candidate.score?.penalties?.domain || 0)
  );
}

function scoreCandidate(candidate = {}, context = {}) {
  if (!candidate || typeof candidate !== "object") return candidate;

  const confidence = clampNumber(candidate.confidence, 0.5, 0, 1);
  const priority = clampNumber(candidate.priority, 0, -10000, 10000);
  const normalizedPriority = Math.max(0, Math.min(10, priority / 80));
  const domain = domainAdjustmentScore(candidate);
  const signals = {
    confidence: Math.round(confidence * 1000) / 1000,
    priority,
    recipeType: normalizeType(candidate),
    source: sourceScore(candidate, context),
    recipeTypeScore: recipeTypeScore(candidate),
    outputTypes: outputTypeScore(candidate),
    evidence: evidenceScore(candidate),
    normalizedPriority,
    domainFit: domain.score,
    domainTraits: domain.traits,
    versionNotes: [
      "BUSINESS_TEMPLATE_TIE_BREAK",
      "SALES_AMOUNT_OVER_QUANTITY",
      "RESEARCH_YEAR_METRIC_PENALTY",
      "GENERIC_INDICATOR_NAME_PENALTY",
      "REDUCED_SCORE_SATURATION",
    ],
  };
  const penalties = {
    quality: penaltyScore(candidate),
    domain: Math.abs(Math.min(0, domain.score)),
  };
  const baseScore = 18;
  const total = clampNumber(
    baseScore +
      confidence * 14 +
      normalizedPriority +
      signals.recipeTypeScore +
      signals.outputTypes +
      signals.source +
      signals.evidence +
      domain.score -
      penalties.quality,
    0,
    0,
    100,
  );

  const rankScore = Math.round(total * 100) / 100;
  const reasonCodes = unique([
    ...(candidate.reasonCodes || []),
    "SCORED_BY_CANDIDATE_SCORING_V1_1",
    signals.source > 0 ? "HAS_SOURCE_TABLE" : "MISSING_SOURCE_TABLE",
    signals.evidence >= 7 ? "HAS_ANALYSIS_EVIDENCE" : "LOW_ANALYSIS_EVIDENCE",
    ...domain.signals,
    ...domain.penalties,
    penalties.quality > 0 ? "HAS_RANKING_PENALTY" : "",
  ]);

  return {
    ...candidate,
    candidateScoreVersion: CANDIDATE_SCORING_VERSION,
    rankScore,
    rankingTier: candidate.rankingTier || "unranked",
    score: {
      version: CANDIDATE_SCORING_VERSION,
      total: rankScore,
      base: baseScore,
      signals,
      penalties,
    },
    reasonCodes,
  };
}

function rankCandidateList(candidates = [], context = {}) {
  const scored = (Array.isArray(candidates) ? candidates : [])
    .map((candidate) => scoreCandidate(candidate, context))
    .sort((a, b) => {
      const scoreDiff = Number(b.rankScore || 0) - Number(a.rankScore || 0);
      if (Math.abs(scoreDiff) > 0.00001) return scoreDiff;

      const typeDiff = candidateTypeSortRank(a) - candidateTypeSortRank(b);
      if (typeDiff) return typeDiff;

      const domainDiff = domainSortScore(b) - domainSortScore(a);
      if (Math.abs(domainDiff) > 0.00001) return domainDiff;

      const priorityDiff = Number(b.priority || 0) - Number(a.priority || 0);
      if (priorityDiff) return priorityDiff;

      return String(a.candidateId || a.id || "").localeCompare(
        String(b.candidateId || b.id || ""),
      );
    });

  return scored.map((candidate, index) => ({
    ...candidate,
    rank: index + 1,
    rankingTier:
      index < TOP_TIER_MAX
        ? "top"
        : index < SECONDARY_TIER_MAX
          ? "secondary"
          : "longTail",
  }));
}

function flattenCandidateGroups(bundle = {}) {
  return [
    ...(Array.isArray(bundle.businessTemplateCandidates)
      ? bundle.businessTemplateCandidates
      : []),
    ...(Array.isArray(bundle.multiSourceCandidates)
      ? bundle.multiSourceCandidates
      : []),
    ...(Array.isArray(bundle.dashboardCandidates)
      ? bundle.dashboardCandidates
      : []),
    ...(Array.isArray(bundle.categoryCandidates)
      ? bundle.categoryCandidates
      : []),
    ...(Array.isArray(bundle.analysisRecipeCandidates)
      ? bundle.analysisRecipeCandidates
      : []),
  ];
}

function candidateGroupsForMultiSourceFallback(bundle = {}) {
  return [
    ...asArray(bundle.analysisRecipeCandidates),
    ...asArray(bundle.businessTemplateCandidates),
    ...asArray(bundle.dashboardCandidates),
    ...asArray(bundle.categoryCandidates),
    ...asArray(bundle.topCandidates),
    ...asArray(bundle.secondaryCandidates),
  ];
}

function getMultiSourceDiagnostics(candidates = []) {
  return candidates?.diagnostics || candidates?.multiSourceDiagnostics || null;
}

function ensureFinalMultiSourceCandidates(bundle = {}, context = {}) {
  const existing = asArray(bundle.multiSourceCandidates);
  const previousDiagnostics =
    bundle.candidateGeneration?.multiSourceCandidates?.diagnostics || null;

  if (existing.length) {
    return {
      candidates: existing,
      diagnostics: {
        version: MULTI_SOURCE_FINAL_INJECTION_VERSION,
        finalInjectionApplied: false,
        reason: "EXISTING_MULTI_SOURCE_CANDIDATES",
        existingCount: existing.length,
        previousDiagnostics,
      },
    };
  }

  if (!multiSourceCandidateBuilder?.buildMultiSourceCandidates) {
    return {
      candidates: existing,
      diagnostics: {
        version: MULTI_SOURCE_FINAL_INJECTION_VERSION,
        finalInjectionApplied: false,
        reason: "MULTI_SOURCE_BUILDER_UNAVAILABLE",
        existingCount: existing.length,
        previousDiagnostics,
      },
    };
  }

  const candidatePayload = candidateGroupsForMultiSourceFallback(bundle);
  const normalizedQueryTables = [
    ...asArray(bundle.normalizedQueryTables),
    ...asArray(bundle.tables),
    ...asArray(context.normalizedQueryTables),
    ...asArray(context.tables),
  ];
  const sourceTablePolicy =
    context.sourceTablePolicy ||
    bundle.sourceTablePolicy ||
    bundle.candidateGeneration?.sourceTablePolicyRaw ||
    bundle.candidateGeneration?.sourceTablePolicy ||
    bundle.candidateGeneration?.sourceTablePolicySummary ||
    {};

  let rebuilt = [];
  let rebuildError = "";
  try {
    rebuilt = multiSourceCandidateBuilder.buildMultiSourceCandidates({
      normalizedQueryTables,
      sourceTablePolicy,
      analysisRecipeCandidates: candidatePayload,
    });
  } catch (error) {
    rebuildError = error?.message || String(error);
    rebuilt = [];
  }

  const rebuildDiagnostics = getMultiSourceDiagnostics(rebuilt);
  const candidates = asArray(rebuilt);

  return {
    candidates,
    diagnostics: {
      version: MULTI_SOURCE_FINAL_INJECTION_VERSION,
      finalInjectionApplied: true,
      reason: candidates.length
        ? "REBUILT_FROM_FINAL_CANDIDATE_BUNDLE"
        : "REBUILD_RETURNED_EMPTY",
      existingCount: existing.length,
      rebuiltCount: candidates.length,
      rebuildError,
      normalizedQueryTableCount: normalizedQueryTables.length,
      analysisRecipeCandidateCount: asArray(bundle.analysisRecipeCandidates)
        .length,
      candidatePayloadCount: candidatePayload.length,
      previousDiagnostics,
      rebuildDiagnostics,
    },
  };
}

function scoreCandidateBundle(bundle = {}, context = {}) {
  const candidateGeneration = bundle.candidateGeneration || {};
  const sourceTablePolicy =
    context.sourceTablePolicy ||
    candidateGeneration.sourceTablePolicy ||
    candidateGeneration.sourceTablePolicySummary ||
    {};
  const scorerContext = { ...context, sourceTablePolicy };
  const finalMultiSource = ensureFinalMultiSourceCandidates(
    bundle,
    scorerContext,
  );
  const multiSourceGenerationDiagnostics = finalMultiSource.diagnostics;

  const analysisRecipeCandidates = rankCandidateList(
    bundle.analysisRecipeCandidates || [],
    scorerContext,
  );
  const categoryCandidates = rankCandidateList(
    bundle.categoryCandidates || [],
    scorerContext,
  );
  const dashboardCandidates = rankCandidateList(
    bundle.dashboardCandidates || [],
    scorerContext,
  );
  const businessTemplateCandidates = rankCandidateList(
    bundle.businessTemplateCandidates || [],
    scorerContext,
  );
  const multiSourceCandidates = rankCandidateList(
    finalMultiSource.candidates || [],
    scorerContext,
  );

  const combined = rankCandidateList(
    [
      ...businessTemplateCandidates,
      ...multiSourceCandidates,
      ...dashboardCandidates,
      ...categoryCandidates,
      ...analysisRecipeCandidates,
    ],
    scorerContext,
  );
  const topCandidates = combined
    .filter((candidate) => candidate.rankingTier === "top")
    .slice(0, TOP_TIER_MAX);
  const secondaryCandidates = combined
    .filter((candidate) => candidate.rankingTier === "secondary")
    .slice(0, SECONDARY_TIER_MAX - TOP_TIER_MAX);

  return {
    ...bundle,
    analysisRecipeCandidates,
    categoryCandidates,
    dashboardCandidates,
    businessTemplateCandidates,
    multiSourceCandidates,
    topCandidates,
    secondaryCandidates,
    candidateScoring: {
      version: CANDIDATE_SCORING_VERSION,
      applied: true,
      generatedAt: new Date().toISOString(),
      notes: [
        "candidate_scoring_v1_1 reduces score saturation and applies domain-aware metric ranking.",
        "multi_source_final_bundle_injection_v1 rebuilds missing multiSourceCandidates from final candidate payload when needed.",
      ],
      multiSourceDiagnostics: multiSourceGenerationDiagnostics,
      counts: {
        totalCandidates: flattenCandidateGroups({
          analysisRecipeCandidates,
          categoryCandidates,
          dashboardCandidates,
          businessTemplateCandidates,
          multiSourceCandidates,
        }).length,
        analysisRecipeCandidates: analysisRecipeCandidates.length,
        categoryCandidates: categoryCandidates.length,
        dashboardCandidates: dashboardCandidates.length,
        businessTemplateCandidates: businessTemplateCandidates.length,
        multiSourceCandidates: multiSourceCandidates.length,
        topCandidates: topCandidates.length,
        secondaryCandidates: secondaryCandidates.length,
      },
      topCandidateIds: topCandidates.map(
        (candidate) => candidate.candidateId || candidate.id,
      ),
    },
    candidateGeneration: {
      ...candidateGeneration,
      candidateScoring: {
        version: CANDIDATE_SCORING_VERSION,
        applied: true,
      },
      multiSourceCandidates: {
        ...(candidateGeneration.multiSourceCandidates || {}),
        finalInjectionVersion: MULTI_SOURCE_FINAL_INJECTION_VERSION,
        finalInjectionApplied:
          multiSourceGenerationDiagnostics?.finalInjectionApplied === true,
        count: multiSourceCandidates.length,
        diagnostics: multiSourceGenerationDiagnostics,
      },
    },
  };
}

function summarizeCandidateCounts(payload = {}) {
  const scoring =
    payload.candidateScoring ||
    payload.candidateGeneration?.candidateScoring ||
    {};
  return {
    analysisRecipeCandidates: asArray(payload.analysisRecipeCandidates).length,
    categoryCandidates: asArray(payload.categoryCandidates).length,
    dashboardCandidates: asArray(payload.dashboardCandidates).length,
    businessTemplateCandidates: asArray(payload.businessTemplateCandidates)
      .length,
    multiSourceCandidates: asArray(payload.multiSourceCandidates).length,
    topCandidates: asArray(payload.topCandidates).length,
    secondaryCandidates: asArray(payload.secondaryCandidates).length,
    scoringVersion: scoring.version || payload.candidateScoring?.version || "",
  };
}

module.exports = {
  CANDIDATE_SCORING_VERSION,
  MULTI_SOURCE_FINAL_INJECTION_VERSION,
  TOP_TIER_MAX,
  SECONDARY_TIER_MAX,
  scoreCandidate,
  rankCandidateList,
  scoreCandidateBundle,
  summarizeCandidateCounts,
};
