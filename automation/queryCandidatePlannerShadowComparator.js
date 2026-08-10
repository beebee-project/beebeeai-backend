const crypto = require("crypto");

const COMPARATOR_VERSION = "query_candidate_planner_api_shadow_comparator_v1";
const COMPARATOR_POLICY_VERSION =
  "api_shadow_candidate_rank_comparison_policy_v1";

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, canonicalize(value[key])]),
  );
}

function sha256(value) {
  const text =
    typeof value === "string" ? value : JSON.stringify(canonicalize(value));
  return crypto.createHash("sha256").update(text).digest("hex");
}

function text(value) {
  return String(value == null ? "" : value).trim();
}

function candidateIdentity(candidate = {}, fallbackIndex = 0) {
  const identity =
    text(candidate.candidateId) ||
    text(candidate.uiCandidateId) ||
    text(candidate.templateId) ||
    text(candidate.recipeId) ||
    text(candidate.id) ||
    text(candidate.key);
  if (identity) return identity;

  const structural = {
    type:
      text(candidate.candidateType) ||
      text(candidate.type) ||
      text(candidate.recipeType),
    tableId: text(candidate.tableId) || text(candidate.sourceTableId),
    operation: text(candidate.operation) || text(candidate.recipeType),
    fallbackIndex,
  };
  return `structural:${sha256(structural)}`;
}

function uniqueCandidates(candidates = []) {
  const seen = new Set();
  const output = [];
  for (const [index, candidate] of candidates.entries()) {
    if (!candidate || typeof candidate !== "object") continue;
    const identity = candidateIdentity(candidate, index);
    if (seen.has(identity)) continue;
    seen.add(identity);
    output.push({ identity, candidate });
  }
  return output;
}

function firstCandidateArray(values = []) {
  for (const value of values) {
    if (Array.isArray(value) && value.length) return value;
  }
  return [];
}

function extractPrimaryCandidates(primaryPayload = {}) {
  const recommended = firstCandidateArray([
    primaryPayload.candidateUiPayload?.recommendedCandidates,
    primaryPayload.topCandidates,
  ]);
  if (recommended.length) return uniqueCandidates(recommended);

  return uniqueCandidates([
    ...(primaryPayload.businessTemplateCandidates || []),
    ...(primaryPayload.analysisRecipeCandidates || []),
    ...(primaryPayload.multiSourceCandidates || []),
    ...(primaryPayload.categoryCandidates || []),
    ...(primaryPayload.dashboardCandidates || []),
    ...(primaryPayload.secondaryCandidates || []),
  ]);
}

function extractShadowCandidates(shadowResolution = {}) {
  const ranked = firstCandidateArray([
    shadowResolution.items,
    shadowResolution.rankedCandidates,
    shadowResolution.reentryItems,
    shadowResolution.rankingResolution?.items,
    shadowResolution.rankingResolution?.rankedCandidates,
    shadowResolution.ranking?.items,
    shadowResolution.shadow?.items,
    shadowResolution.result?.items,
    shadowResolution.result?.rankingResolution?.items,
    shadowResolution.plannerResolution?.acceptedItems,
    shadowResolution.plannerResolution?.items,
  ]);

  const ordered = [...ranked].sort((left, right) => {
    const leftRank = Number(
      left.shadowRank ??
        left.rank ??
        left.ranking?.rank ??
        Number.MAX_SAFE_INTEGER,
    );
    const rightRank = Number(
      right.shadowRank ??
        right.rank ??
        right.ranking?.rank ??
        Number.MAX_SAFE_INTEGER,
    );
    return leftRank - rightRank;
  });
  return uniqueCandidates(ordered);
}

function round(value, digits = 4) {
  const scale = 10 ** digits;
  return Math.round(Number(value || 0) * scale) / scale;
}

function hashedIdentities(entries = [], limit = 20) {
  return entries
    .slice(0, limit)
    .map((entry) => sha256(`candidate-identity:${entry.identity}`));
}

function compareCandidatePlannerShadow({
  primaryPayload = {},
  shadowResolution = {},
  maxCandidates = 20,
} = {}) {
  const primary = extractPrimaryCandidates(primaryPayload).slice(
    0,
    maxCandidates,
  );
  const shadow = extractShadowCandidates(shadowResolution).slice(
    0,
    maxCandidates,
  );

  const primaryIds = primary.map((entry) => entry.identity);
  const shadowIds = shadow.map((entry) => entry.identity);
  const primarySet = new Set(primaryIds);
  const shadowSet = new Set(shadowIds);
  const shared = primaryIds.filter((id) => shadowSet.has(id));
  const primaryOnly = primaryIds.filter((id) => !shadowSet.has(id));
  const shadowOnly = shadowIds.filter((id) => !primarySet.has(id));
  const unionSize = new Set([...primaryIds, ...shadowIds]).size;

  const rankDistances = shared.map((id) =>
    Math.abs(primaryIds.indexOf(id) - shadowIds.indexOf(id)),
  );
  const rankDistanceTotal = rankDistances.reduce(
    (sum, distance) => sum + distance,
    0,
  );
  const rankDenominator =
    Math.max(primaryIds.length, shadowIds.length, 1) *
    Math.max(shared.length, 1);
  const rankAgreement = shared.length
    ? Math.max(0, 1 - rankDistanceTotal / rankDenominator)
    : 0;

  const top3Primary = new Set(primaryIds.slice(0, 3));
  const top3Overlap = shadowIds
    .slice(0, 3)
    .filter((id) => top3Primary.has(id)).length;
  const exactOrder =
    primaryIds.length === shadowIds.length &&
    primaryIds.every((id, index) => shadowIds[index] === id);
  const top1Same = Boolean(
    primaryIds.length && shadowIds.length && primaryIds[0] === shadowIds[0],
  );
  const jaccard = unionSize ? shared.length / unionSize : 1;

  let verdict = "MISMATCH";
  if (!shadow.length) verdict = "NO_SHADOW_CANDIDATES";
  else if (exactOrder) verdict = "MATCH";
  else if (shared.length > 0) verdict = "PARTIAL_MATCH";

  const result = {
    version: COMPARATOR_VERSION,
    policyVersion: COMPARATOR_POLICY_VERSION,
    verdict,
    counts: Object.freeze({
      primary: primary.length,
      shadow: shadow.length,
      shared: shared.length,
      primaryOnly: primaryOnly.length,
      shadowOnly: shadowOnly.length,
    }),
    metrics: Object.freeze({
      exactOrder,
      top1Same,
      top3Overlap,
      jaccard: round(jaccard),
      rankAgreement: round(rankAgreement),
    }),
    fingerprints: Object.freeze({
      primaryOrderSha256: sha256(primaryIds),
      shadowOrderSha256: sha256(shadowIds),
      sharedSetSha256: sha256([...shared].sort()),
      primaryOnlyIdentitySha256: Object.freeze(
        hashedIdentities(primaryOnly.map((identity) => ({ identity }))),
      ),
      shadowOnlyIdentitySha256: Object.freeze(
        hashedIdentities(shadowOnly.map((identity) => ({ identity }))),
      ),
    }),
    privacy: Object.freeze({
      rawIdentifiersIncluded: false,
      rawCandidatePayloadIncluded: false,
      fileNameIncluded: false,
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
    }),
    productionCandidateMerge: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
  };

  return Object.freeze(result);
}

module.exports = Object.freeze({
  COMPARATOR_VERSION,
  COMPARATOR_POLICY_VERSION,
  sha256,
  candidateIdentity,
  extractPrimaryCandidates,
  extractShadowCandidates,
  compareCandidatePlannerShadow,
});
