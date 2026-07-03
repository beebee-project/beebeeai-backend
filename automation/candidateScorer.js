const CANDIDATE_SCORING_VERSION = "candidate_scoring_v1";

const TOP_TIER_MAX = 7;
const SECONDARY_TIER_MAX = 30;

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
    ...(candidate.reasonCodes || []),
    ...(candidate.recipeIds || []),
  ]
    .filter(Boolean)
    .join(" ");
}

function recipeTypeScore(candidate = {}) {
  const type = normalizeType(candidate);
  const recipeIds = asArray(candidate.recipeIds).join(" ");
  const text = `${type} ${recipeIds}`;

  if (/businessTemplate/i.test(candidate.candidateType || "")) return 22;
  if (/dashboard/i.test(candidate.candidateType || "")) return 18;
  if (/time_(sum|avg|growth|trend)|cumulative_sum|wide_time_trend/i.test(text))
    return 18;
  if (/group_(sum|avg|count|summary)|composition_ratio|top_bottom/i.test(text))
    return 15;
  if (/cross_(count|sum)|status_count|category_count/i.test(text)) return 12;
  return 8;
}

function outputTypeScore(candidate = {}) {
  const outputTypes = asArray(candidate.outputTypes);
  if (!outputTypes.length) return 0;
  let score = 0;
  if (outputTypes.includes("summarySheet")) score += 4;
  if (outputTypes.includes("analysisReport")) score += 3;
  if (outputTypes.includes("ppt")) score += 3;
  return Math.min(8, score);
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
  if (sourceTableId) score += 8;
  if (sourceSheetName) score += 6;
  if (primaryTableId && sourceTableIds.includes(primaryTableId)) score += 5;
  if (primarySheetName && sourceSheetNames.includes(primarySheetName))
    score += 4;
  if (scope === "multiTable") score += 4;
  if (sourceTableIds.length > 1) score += 4;

  return score;
}

function evidenceScore(candidate = {}) {
  const text = collectCandidateText(candidate);
  const matchedCount = Number(candidate.matchedCount || 0);
  let score = 0;

  if (matchedCount > 0) score += Math.min(10, matchedCount * 2);
  if (asArray(candidate.matchedHeaders).length)
    score += Math.min(8, asArray(candidate.matchedHeaders).length * 2);
  if (asArray(candidate.reasonCodes).length)
    score += Math.min(6, asArray(candidate.reasonCodes).length);
  if (candidate.primaryCandidate) score += 4;
  if (asArray(candidate.candidates).length >= 2) score += 4;

  if (hasAnyToken(text, ["월", "연도", "기간", "추이", "growth", "trend"]))
    score += 4;
  if (
    hasAnyToken(text, ["합계", "금액", "매출", "집행", "연봉", "sum", "amount"])
  )
    score += 4;
  if (hasAnyToken(text, ["평균", "avg", "average"])) score += 3;
  if (hasAnyToken(text, ["비중", "구성", "ratio", "composition"])) score += 3;
  if (hasAnyToken(text, ["상위", "하위", "top", "bottom"])) score += 3;

  return Math.min(22, score);
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

function scoreCandidate(candidate = {}, context = {}) {
  if (!candidate || typeof candidate !== "object") return candidate;

  const confidence = clampNumber(candidate.confidence, 0.5, 0, 1);
  const priority = clampNumber(candidate.priority, 0, -10000, 10000);
  const normalizedPriority = Math.max(0, Math.min(18, priority / 50));
  const signals = {
    confidence: Math.round(confidence * 1000) / 1000,
    priority,
    recipeType: normalizeType(candidate),
    source: sourceScore(candidate, context),
    recipeTypeScore: recipeTypeScore(candidate),
    outputTypes: outputTypeScore(candidate),
    evidence: evidenceScore(candidate),
    normalizedPriority,
  };
  const penalties = {
    quality: penaltyScore(candidate),
  };
  const baseScore = 35;
  const total = clampNumber(
    baseScore +
      confidence * 20 +
      normalizedPriority +
      signals.recipeTypeScore +
      signals.outputTypes +
      signals.source +
      signals.evidence -
      penalties.quality,
    0,
    0,
    100,
  );

  const rankScore = Math.round(total * 100) / 100;
  const reasonCodes = unique([
    ...(candidate.reasonCodes || []),
    "SCORED_BY_CANDIDATE_SCORING_V1",
    signals.source > 0 ? "HAS_SOURCE_TABLE" : "MISSING_SOURCE_TABLE",
    signals.evidence >= 8 ? "HAS_ANALYSIS_EVIDENCE" : "LOW_ANALYSIS_EVIDENCE",
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
      if (scoreDiff) return scoreDiff;
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

function scoreCandidateBundle(bundle = {}, context = {}) {
  const candidateGeneration = bundle.candidateGeneration || {};
  const sourceTablePolicy =
    context.sourceTablePolicy ||
    candidateGeneration.sourceTablePolicy ||
    candidateGeneration.sourceTablePolicySummary ||
    null;
  const scorerContext = { ...context, sourceTablePolicy };

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

  const combined = rankCandidateList(
    [
      ...businessTemplateCandidates,
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
    topCandidates,
    secondaryCandidates,
    candidateScoring: {
      version: CANDIDATE_SCORING_VERSION,
      applied: true,
      generatedAt: new Date().toISOString(),
      counts: {
        totalCandidates: flattenCandidateGroups({
          analysisRecipeCandidates,
          categoryCandidates,
          dashboardCandidates,
          businessTemplateCandidates,
        }).length,
        analysisRecipeCandidates: analysisRecipeCandidates.length,
        categoryCandidates: categoryCandidates.length,
        dashboardCandidates: dashboardCandidates.length,
        businessTemplateCandidates: businessTemplateCandidates.length,
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
    topCandidates: asArray(payload.topCandidates).length,
    secondaryCandidates: asArray(payload.secondaryCandidates).length,
    scoringVersion: scoring.version || payload.candidateScoring?.version || "",
  };
}

module.exports = {
  CANDIDATE_SCORING_VERSION,
  TOP_TIER_MAX,
  SECONDARY_TIER_MAX,
  scoreCandidate,
  rankCandidateList,
  scoreCandidateBundle,
  summarizeCandidateCounts,
};
