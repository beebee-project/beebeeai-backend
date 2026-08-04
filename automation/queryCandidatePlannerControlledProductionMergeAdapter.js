"use strict";

const {
  OPERATIONS,
} = require("./queryCandidatePlannerFeatureControl");
const {
  candidateIdentity,
  extractShadowCandidates,
  sha256,
} = require("./queryCandidatePlannerShadowComparator");

const ADAPTER_VERSION =
  "query_candidate_planner_controlled_production_merge_adapter_v1";
const PLAN_VERSION =
  "query_candidate_planner_controlled_production_merge_plan_v1";
const RESULT_VERSION =
  "query_candidate_planner_controlled_production_merge_result_v1";
const PROMOTION_GATE_DECISION_VERSION =
  "query_candidate_planner_controlled_production_promotion_gate_decision_v1";
const DEFAULT_MAX_CANDIDATES = 20;
const MAX_CANDIDATES_LIMIT = 100;

const CANDIDATE_GROUPS = Object.freeze({
  ANALYSIS: "analysisRecipeCandidates",
  BUSINESS_TEMPLATE: "businessTemplateCandidates",
  MULTI_SOURCE: "multiSourceCandidates",
  CATEGORY: "categoryCandidates",
  DASHBOARD: "dashboardCandidates",
});

function text(value, maxLength = 240) {
  return String(value == null ? "" : value)
    .trim()
    .replace(/[\r\n\t]/g, " ")
    .slice(0, maxLength);
}

function finiteNumber(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function boundedCandidateLimit(value) {
  const parsed = Math.floor(finiteNumber(value, DEFAULT_MAX_CANDIDATES));
  return Math.max(1, Math.min(MAX_CANDIDATES_LIMIT, parsed));
}

function canonicalClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function safeSourceTableIds(candidate = {}) {
  const values = Array.isArray(candidate.sourceTableIds)
    ? candidate.sourceTableIds
    : [];
  return Object.freeze(
    values.map((value) => text(value, 160)).filter(Boolean).slice(0, 20),
  );
}

function safeCandidateType(candidate = {}) {
  return text(
    candidate.candidateType || candidate.type || candidate.kind || "analysis",
    80,
  );
}

function safeCandidateStatus(candidate = {}) {
  const value = text(
    candidate.status || candidate.feasibilityStatus || candidate.disposition,
    80,
  );
  if (!value) return "";
  if (/\bREADY\b|PRODUCTION_READY|PROMOTED/i.test(value)) return "";
  return value;
}

function normalizeShadowCandidate(candidate = {}, fallbackIndex = 0) {
  const candidateId = candidateIdentity(candidate, fallbackIndex);
  const candidateType = safeCandidateType(candidate);
  const rank = Math.max(
    1,
    Math.floor(
      finiteNumber(
        candidate.shadowRank ?? candidate.rank ?? candidate.ranking?.rank,
        fallbackIndex + 1,
      ),
    ),
  );
  const status = safeCandidateStatus(candidate);
  const normalized = {
    candidateId,
    uiCandidateId:
      text(candidate.uiCandidateId, 200) ||
      `${candidateType || "candidate"}:${candidateId}`,
    candidateType,
    recipeType: text(candidate.recipeType || candidate.operation, 120),
    recipeId: text(candidate.recipeId || candidate.recipeType, 160),
    operation: text(candidate.operation || candidate.recipeType, 120),
    tableId: text(candidate.tableId || candidate.sourceTableId, 160),
    sourceTableIds: safeSourceTableIds(candidate),
    title: text(candidate.title || candidate.label || candidate.name, 240),
    description: text(candidate.description || candidate.summary, 500),
    rank,
    score: finiteNumber(
      candidate.score?.total ??
        candidate.score?.value ??
        candidate.totalScore ??
        candidate.score,
      0,
    ),
  };
  if (status) normalized.status = status;
  return Object.freeze(normalized);
}

function candidateGroup(candidate = {}) {
  const type = text(candidate.candidateType, 100).toLowerCase();
  if (type.includes("business") || type.includes("template")) {
    return CANDIDATE_GROUPS.BUSINESS_TEMPLATE;
  }
  if (type.includes("multi")) return CANDIDATE_GROUPS.MULTI_SOURCE;
  if (type.includes("category")) return CANDIDATE_GROUPS.CATEGORY;
  if (type.includes("dashboard")) return CANDIDATE_GROUPS.DASHBOARD;
  return CANDIDATE_GROUPS.ANALYSIS;
}

function normalizeShadowCandidateSet(shadowResolution = {}, maxCandidates) {
  const limit = boundedCandidateLimit(maxCandidates);
  const extracted = extractShadowCandidates(shadowResolution);
  const seen = new Set();
  const candidates = [];

  for (const [index, entry] of extracted.entries()) {
    const normalized = normalizeShadowCandidate(entry.candidate, index);
    if (seen.has(normalized.candidateId)) continue;
    seen.add(normalized.candidateId);
    candidates.push(normalized);
    if (candidates.length >= limit) break;
  }

  candidates.sort((left, right) => {
    if (left.rank !== right.rank) return left.rank - right.rank;
    return left.candidateId.localeCompare(right.candidateId);
  });
  return Object.freeze(candidates);
}

function emptyCandidateProjection() {
  return Object.freeze({
    topCandidates: Object.freeze([]),
    recommendedCandidates: Object.freeze([]),
    analysisRecipeCandidates: Object.freeze([]),
    businessTemplateCandidates: Object.freeze([]),
    multiSourceCandidates: Object.freeze([]),
    categoryCandidates: Object.freeze([]),
    dashboardCandidates: Object.freeze([]),
  });
}

function projectCandidateContract(candidates = []) {
  const groups = {
    analysisRecipeCandidates: [],
    businessTemplateCandidates: [],
    multiSourceCandidates: [],
    categoryCandidates: [],
    dashboardCandidates: [],
  };

  for (const candidate of candidates) {
    groups[candidateGroup(candidate)].push(candidate);
  }

  const recommendedCandidates = candidates.map((candidate) =>
    Object.freeze({
      uiCandidateId: candidate.uiCandidateId,
      candidateId: candidate.candidateId,
      candidateType: candidate.candidateType,
      rank: candidate.rank,
      title: candidate.title,
      description: candidate.description,
    }),
  );

  return Object.freeze({
    topCandidates: Object.freeze([...candidates]),
    recommendedCandidates: Object.freeze(recommendedCandidates),
    analysisRecipeCandidates: Object.freeze(groups.analysisRecipeCandidates),
    businessTemplateCandidates: Object.freeze(
      groups.businessTemplateCandidates,
    ),
    multiSourceCandidates: Object.freeze(groups.multiSourceCandidates),
    categoryCandidates: Object.freeze(groups.categoryCandidates),
    dashboardCandidates: Object.freeze(groups.dashboardCandidates),
  });
}

function adapterGuardrails(overrides = {}) {
  return Object.freeze({
    defaultEnabled: false,
    routeWired: false,
    controllerWired: false,
    primaryResponseAuthority: true,
    primaryPayloadMutated: false,
    responsePayloadMutation: false,
    responseHeaderMutation: false,
    responseStatusMutation: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    sourceCandidateStatusMutation: false,
    promotionGateRequired: true,
    failClosed: true,
    ...overrides,
  });
}

function invalidPlan(reason) {
  return Object.freeze({
    version: PLAN_VERSION,
    adapterVersion: ADAPTER_VERSION,
    status: "INVALID_INPUT_FAIL_CLOSED",
    reason,
    candidateProjection: emptyCandidateProjection(),
    counts: Object.freeze({ projected: 0 }),
    fingerprints: Object.freeze({ candidateOrderSha256: sha256([]) }),
    guardrails: adapterGuardrails(),
  });
}

function buildControlledProductionMergePlan({
  primaryPayload = null,
  shadowResolution = null,
  maxCandidates = DEFAULT_MAX_CANDIDATES,
} = {}) {
  if (!isPlainObject(primaryPayload)) {
    return invalidPlan("PRIMARY_PAYLOAD_REQUIRED");
  }
  if (!isPlainObject(shadowResolution)) {
    return invalidPlan("SHADOW_RESOLUTION_REQUIRED");
  }

  const candidates = normalizeShadowCandidateSet(
    shadowResolution,
    maxCandidates,
  );
  const candidateProjection = projectCandidateContract(candidates);
  const candidateOrder = candidates.map((candidate) => candidate.candidateId);
  const status = candidates.length ? "PLAN_READY" : "NO_SHADOW_CANDIDATES";

  return Object.freeze({
    version: PLAN_VERSION,
    adapterVersion: ADAPTER_VERSION,
    status,
    reason: candidates.length
      ? "SHADOW_CANDIDATE_CONTRACT_PROJECTED"
      : "NO_SHADOW_CANDIDATES",
    candidateProjection,
    counts: Object.freeze({
      projected: candidates.length,
      analysis: candidateProjection.analysisRecipeCandidates.length,
      businessTemplate:
        candidateProjection.businessTemplateCandidates.length,
      multiSource: candidateProjection.multiSourceCandidates.length,
      category: candidateProjection.categoryCandidates.length,
      dashboard: candidateProjection.dashboardCandidates.length,
    }),
    fingerprints: Object.freeze({
      candidateOrderSha256: sha256(candidateOrder),
      projectionSha256: sha256(candidateProjection),
      primaryContractSha256: sha256({
        ok: primaryPayload.ok === true,
        source: text(primaryPayload.source),
        topLevelKeys: Object.keys(primaryPayload).sort(),
      }),
    }),
    guardrails: adapterGuardrails(),
  });
}

function evaluatePromotionGateDecision(decision) {
  const checks = Object.freeze({
    objectPresent: isPlainObject(decision),
    version:
      decision?.version === PROMOTION_GATE_DECISION_VERSION,
    allowed: decision?.allowed === true,
    decision: decision?.decision === "ALLOW",
    operation:
      decision?.operation === OPERATIONS.PRODUCTION_CANDIDATE_MERGE,
    failClosed: decision?.failClosed === true,
    adapterVersion: decision?.adapterVersion === ADAPTER_VERSION,
  });
  const valid = Object.values(checks).every(Boolean);
  return Object.freeze({
    valid,
    checks,
    reason: valid
      ? "PROMOTION_GATE_DECISION_VALID"
      : "PROMOTION_GATE_DECISION_REQUIRED",
  });
}

function evaluateControlledProductionMergeAuthorization({
  featureControl = null,
  readinessGate = null,
  promotionGateDecision = null,
} = {}) {
  if (!featureControl || typeof featureControl.evaluate !== "function") {
    return Object.freeze({
      allowed: false,
      reason: "FEATURE_CONTROL_REQUIRED",
      featureDecision: null,
      promotionGate: evaluatePromotionGateDecision(promotionGateDecision),
      failClosed: true,
    });
  }

  const featureDecision = featureControl.evaluate(
    OPERATIONS.PRODUCTION_CANDIDATE_MERGE,
    { readinessGate },
  );
  if (!featureDecision.allowed) {
    return Object.freeze({
      allowed: false,
      reason: featureDecision.reason,
      featureDecision,
      promotionGate: evaluatePromotionGateDecision(promotionGateDecision),
      failClosed: true,
    });
  }

  const promotionGate = evaluatePromotionGateDecision(
    promotionGateDecision,
  );
  if (!promotionGate.valid) {
    return Object.freeze({
      allowed: false,
      reason: promotionGate.reason,
      featureDecision,
      promotionGate,
      failClosed: true,
    });
  }

  return Object.freeze({
    allowed: true,
    reason: "CONTROLLED_PRODUCTION_MERGE_AUTHORIZED",
    featureDecision,
    promotionGate,
    failClosed: true,
  });
}

function applyMergePlanToCopy(primaryPayload, plan) {
  const mergedPayload = canonicalClone(primaryPayload);
  const projection = plan.candidateProjection;

  mergedPayload.topCandidates = canonicalClone(projection.topCandidates);
  mergedPayload.analysisRecipeCandidates = canonicalClone(
    projection.analysisRecipeCandidates,
  );
  mergedPayload.businessTemplateCandidates = canonicalClone(
    projection.businessTemplateCandidates,
  );
  mergedPayload.multiSourceCandidates = canonicalClone(
    projection.multiSourceCandidates,
  );
  mergedPayload.categoryCandidates = canonicalClone(
    projection.categoryCandidates,
  );
  mergedPayload.dashboardCandidates = canonicalClone(
    projection.dashboardCandidates,
  );
  mergedPayload.candidateUiPayload = {
    ...(isPlainObject(mergedPayload.candidateUiPayload)
      ? mergedPayload.candidateUiPayload
      : {}),
    recommendedCandidates: canonicalClone(
      projection.recommendedCandidates,
    ),
  };

  return mergedPayload;
}

function controlledProductionMergeAdapter({
  primaryPayload = null,
  shadowResolution = null,
  featureControl = null,
  readinessGate = null,
  promotionGateDecision = null,
  apply = false,
  maxCandidates = DEFAULT_MAX_CANDIDATES,
} = {}) {
  const plan = buildControlledProductionMergePlan({
    primaryPayload,
    shadowResolution,
    maxCandidates,
  });

  if (plan.status === "INVALID_INPUT_FAIL_CLOSED") {
    return Object.freeze({
      version: RESULT_VERSION,
      adapterVersion: ADAPTER_VERSION,
      status: "BLOCKED",
      reason: plan.reason,
      applied: false,
      authorization: null,
      plan,
      mergedPayload: null,
      guardrails: adapterGuardrails(),
    });
  }

  if (plan.status === "NO_SHADOW_CANDIDATES") {
    return Object.freeze({
      version: RESULT_VERSION,
      adapterVersion: ADAPTER_VERSION,
      status: "BLOCKED",
      reason: "NO_SHADOW_CANDIDATES",
      applied: false,
      authorization: null,
      plan,
      mergedPayload: null,
      guardrails: adapterGuardrails(),
    });
  }

  const authorization = evaluateControlledProductionMergeAuthorization({
    featureControl,
    readinessGate,
    promotionGateDecision,
  });
  if (!authorization.allowed) {
    return Object.freeze({
      version: RESULT_VERSION,
      adapterVersion: ADAPTER_VERSION,
      status: "BLOCKED",
      reason: authorization.reason,
      applied: false,
      authorization,
      plan,
      mergedPayload: null,
      guardrails: adapterGuardrails(),
    });
  }

  if (apply !== true) {
    return Object.freeze({
      version: RESULT_VERSION,
      adapterVersion: ADAPTER_VERSION,
      status: "DRY_RUN_READY",
      reason: "AUTHORIZED_BUT_NOT_APPLIED",
      applied: false,
      authorization,
      plan,
      mergedPayload: null,
      guardrails: adapterGuardrails(),
    });
  }

  const beforeSha256 = sha256(primaryPayload);
  const mergedPayload = applyMergePlanToCopy(primaryPayload, plan);
  const afterOriginalSha256 = sha256(primaryPayload);

  return Object.freeze({
    version: RESULT_VERSION,
    adapterVersion: ADAPTER_VERSION,
    status: "MERGED_COPY_READY",
    reason: "CONTROLLED_MERGE_ADAPTED_TO_COPY",
    applied: true,
    authorization,
    plan,
    mergedPayload,
    primaryPayloadUnchanged: beforeSha256 === afterOriginalSha256,
    guardrails: adapterGuardrails({
      primaryPayloadMutated: false,
      responsePayloadMutation: false,
    }),
  });
}

module.exports = Object.freeze({
  ADAPTER_VERSION,
  PLAN_VERSION,
  RESULT_VERSION,
  PROMOTION_GATE_DECISION_VERSION,
  DEFAULT_MAX_CANDIDATES,
  MAX_CANDIDATES_LIMIT,
  CANDIDATE_GROUPS,
  normalizeShadowCandidate,
  normalizeShadowCandidateSet,
  projectCandidateContract,
  buildControlledProductionMergePlan,
  evaluatePromotionGateDecision,
  evaluateControlledProductionMergeAuthorization,
  controlledProductionMergeAdapter,
});
