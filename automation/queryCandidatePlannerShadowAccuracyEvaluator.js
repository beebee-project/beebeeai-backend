"use strict";

const crypto = require("crypto");
const {
  DATASET_VERSION: ACCURACY_DATASET_VERSION,
  DECISIONS: ACCURACY_DECISIONS,
  canonicalJson,
  evaluateAccuracyDataset,
  findSensitivePaths: findAccuracySensitivePaths,
  sha256,
  validateAccuracyEvaluationDataset,
} = require("./queryCandidatePlannerAccuracyEvaluator");

const EVALUATOR_VERSION =
  "query_candidate_planner_shadow_accuracy_evaluator_v1";
const REPORT_VERSION =
  "query_candidate_planner_shadow_accuracy_evaluation_report_v1";
const OBSERVATION_DATASET_VERSION =
  "query_candidate_planner_shadow_accuracy_observation_dataset_v1";
const THRESHOLD_POLICY_VERSION =
  "query_candidate_planner_shadow_accuracy_threshold_policy_v1";
const CAPTURE_POLICY_VERSION =
  "query_candidate_planner_shadow_accuracy_capture_policy_v1";

const DECISIONS = Object.freeze({
  PASS: "EVALUATION_PASS",
  BLOCKED: "EVALUATION_BLOCKED",
});

const OBSERVATION_STATUSES = Object.freeze({
  COMPLETED: "COMPLETED",
  COMPLETED_SAFE: "COMPLETED_SAFE",
  BLOCKED: "BLOCKED",
  FAILED_SAFE: "FAILED_SAFE",
  TIMEOUT_SAFE: "TIMEOUT_SAFE",
});

const COMPLETED_STATUSES = new Set([
  OBSERVATION_STATUSES.COMPLETED,
  OBSERVATION_STATUSES.COMPLETED_SAFE,
]);

const SHA256_RE = /^[a-f0-9]{64}$/i;
const FORBIDDEN_KEYS = new Set([
  "rows",
  "rawRows",
  "rawData",
  "sampleValues",
  "samples",
  "fileName",
  "originalFileName",
  "originalName",
  "email",
  "userId",
  "tenantId",
  "queryTablesKey",
  "storageKey",
  "cacheSecret",
  "rawPayload",
  "rawPrimaryResponse",
  "rawShadowResolution",
  "providerResponse",
  "prompt",
]);

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function clone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function freezeDeep(value) {
  if (Array.isArray(value)) {
    value.forEach(freezeDeep);
    return Object.freeze(value);
  }
  if (isPlainObject(value) && !Object.isFrozen(value)) {
    Object.values(value).forEach(freezeDeep);
    Object.freeze(value);
  }
  return value;
}

function text(value, maxLength = 200) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function number(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function round(value, digits = 6) {
  if (!Number.isFinite(value)) return 0;
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function safeSha256(value) {
  const normalized = text(value, 64).toLowerCase();
  return SHA256_RE.test(normalized) ? normalized : "";
}

function stableSha256(value) {
  return crypto
    .createHash("sha256")
    .update(typeof value === "string" ? value : canonicalJson(value))
    .digest("hex");
}

function uniqueStrings(values) {
  const seen = new Set();
  const result = [];
  for (const value of Array.isArray(values) ? values : []) {
    const normalized = text(value);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

function findForbiddenPaths(value, basePath = "$") {
  const paths = [];
  if (Array.isArray(value)) {
    value.forEach((entry, index) => {
      paths.push(...findForbiddenPaths(entry, `${basePath}[${index}]`));
    });
    return paths;
  }
  if (!isPlainObject(value)) return paths;
  for (const [key, entry] of Object.entries(value)) {
    const childPath = `${basePath}.${key}`;
    if (FORBIDDEN_KEYS.has(key)) paths.push(childPath);
    paths.push(...findForbiddenPaths(entry, childPath));
  }
  return paths;
}

function normalizeCandidate(candidate = {}, index = 0) {
  const candidateId = text(
    candidate.candidateId ||
      candidate.id ||
      candidate.recipeId ||
      candidate.recipeType,
  );
  const statusText = text(
    candidate.status || candidate.result || candidate.disposition,
  ).toUpperCase();
  const rejected = ["REJECTED", "BLOCKED", "INELIGIBLE", "UNSUPPORTED"].includes(
    statusText,
  );
  return Object.freeze({
    candidateId,
    rank:
      Number.isInteger(candidate.rank) && candidate.rank > 0
        ? candidate.rank
        : index + 1,
    status: rejected ? "REJECTED" : "ACCEPTED",
    productionEligible:
      !rejected && candidate.productionEligible !== false,
  });
}

function candidateLists(shadowResolution = {}) {
  return [
    shadowResolution.plannerResolution?.items,
    shadowResolution.plannerResolution?.candidates,
    shadowResolution.candidateResolution?.items,
    shadowResolution.rankingResolution?.items,
    shadowResolution.items,
    shadowResolution.candidates,
    shadowResolution.topCandidates,
  ].filter(Array.isArray);
}

function extractCandidates(shadowResolution = {}) {
  const seen = new Set();
  const candidates = [];
  for (const list of candidateLists(shadowResolution)) {
    for (const candidate of list) {
      if (!isPlainObject(candidate)) continue;
      const normalized = normalizeCandidate(candidate, candidates.length);
      if (!normalized.candidateId || seen.has(normalized.candidateId)) continue;
      seen.add(normalized.candidateId);
      candidates.push(normalized);
      if (candidates.length >= 100) break;
    }
    if (candidates.length >= 100) break;
  }
  return Object.freeze(
    candidates
      .slice()
      .sort((left, right) =>
        left.rank - right.rank || left.candidateId.localeCompare(right.candidateId),
      )
      .map((candidate, index) =>
        Object.freeze({ ...candidate, rank: index + 1 }),
      ),
  );
}

function extractDomain(shadowResolution = {}) {
  return text(
    shadowResolution.businessDomainProfile?.primaryDomain ||
      shadowResolution.semanticProfile?.primaryDomain ||
      shadowResolution.semanticProfile?.businessDomain ||
      shadowResolution.primaryDomain ||
      shadowResolution.domain ||
      "UNKNOWN",
    100,
  ) || "UNKNOWN";
}

function extractIntent(shadowResolution = {}) {
  return text(
    shadowResolution.businessDomainProfile?.datasetIntent ||
      shadowResolution.semanticProfile?.datasetIntent ||
      shadowResolution.semanticProfile?.intent ||
      shadowResolution.datasetIntent ||
      shadowResolution.intent ||
      "UNKNOWN",
    100,
  ) || "UNKNOWN";
}

function extractFallback(shadowResolution = {}) {
  const fallback = isPlainObject(shadowResolution.fallback)
    ? shadowResolution.fallback
    : isPlainObject(shadowResolution.plannerResolution?.fallback)
      ? shadowResolution.plannerResolution.fallback
      : {};
  return Object.freeze({
    applied:
      shadowResolution.fallbackApplied === true || fallback.applied === true,
    reason: text(
      shadowResolution.fallbackReason || fallback.reason,
      120,
    ),
  });
}

function extractReviewRequired(shadowResolution = {}) {
  return (
    shadowResolution.reviewRequired === true ||
    shadowResolution.businessDomainProfile?.reviewRequired === true ||
    shadowResolution.semanticProfile?.reviewRequired === true ||
    shadowResolution.plannerResolution?.reviewRequired === true
  );
}

function extractUnsupportedRejected(shadowResolution = {}, candidates = []) {
  const status = text(
    shadowResolution.status || shadowResolution.plannerResolution?.status,
  ).toUpperCase();
  if (shadowResolution.unsupportedRejected === true) return true;
  if (["UNSUPPORTED_REJECTED", "REJECTED_UNSUPPORTED"].includes(status)) {
    return true;
  }
  return false;
}

function sanitizeComparison(comparison = null) {
  if (!isPlainObject(comparison)) return null;
  return Object.freeze({
    verdict: text(comparison.verdict || "NOT_AVAILABLE", 60),
    counts: Object.freeze({
      primary: Math.max(0, number(comparison.counts?.primary)),
      shadow: Math.max(0, number(comparison.counts?.shadow)),
      shared: Math.max(0, number(comparison.counts?.shared)),
      primaryOnly: Math.max(0, number(comparison.counts?.primaryOnly)),
      shadowOnly: Math.max(0, number(comparison.counts?.shadowOnly)),
    }),
    metrics: Object.freeze({
      exactOrder: comparison.metrics?.exactOrder === true,
      top1Same: comparison.metrics?.top1Same === true,
      top3Overlap: round(number(comparison.metrics?.top3Overlap)),
      jaccard: round(number(comparison.metrics?.jaccard)),
      rankAgreement: round(number(comparison.metrics?.rankAgreement)),
    }),
    rawIdentifiersIncluded: false,
  });
}

function sanitizeGuardrails(guardrails = {}) {
  return Object.freeze({
    shadowOnly: guardrails.shadowOnly !== false,
    primaryResponseAuthority:
      guardrails.primaryResponseAuthority !== false,
    responsePayloadMutation:
      guardrails.responsePayloadMutation === true,
    responseHeaderMutation:
      guardrails.responseHeaderMutation === true,
    responseStatusMutation:
      guardrails.responseStatusMutation === true,
    productionCandidateMerge:
      guardrails.productionCandidateMerge === true,
    productionReadyAssignment:
      guardrails.productionReadyAssignment === true,
    productionRouteChanged:
      guardrails.productionRouteChanged === true,
  });
}

function sanitizePrivacy(privacy = {}) {
  return Object.freeze({
    rawPrimaryResponseIncluded:
      privacy.rawPrimaryResponseIncluded === true,
    rawShadowResolutionIncluded:
      privacy.rawShadowResolutionIncluded === true,
    rawRowsIncluded: privacy.rawRowsIncluded === true,
    sampleValuesIncluded: privacy.sampleValuesIncluded === true,
    fileNameIncluded: privacy.fileNameIncluded === true,
    originalFileNameIncluded:
      privacy.originalFileNameIncluded === true,
    queryTablesKeyIncluded:
      privacy.queryTablesKeyIncluded === true,
    userIdentityIncluded:
      privacy.userIdentityIncluded === true,
    tenantIdIncluded: privacy.tenantIdIncluded === true,
  });
}

function buildShadowAccuracyObservation({
  caseId,
  apiShadowObservation = {},
  shadowResolution = {},
} = {}) {
  const normalizedCaseId = text(caseId, 160);
  if (!normalizedCaseId) {
    throw new Error("caseId is required");
  }
  const forbiddenPaths = findForbiddenPaths({
    apiShadowObservation,
    shadowResolution,
  });
  if (forbiddenPaths.length > 0) {
    const error = new Error(
      `forbidden sensitive shadow capture input: ${forbiddenPaths[0]}`,
    );
    error.code = "SHADOW_ACCURACY_CAPTURE_PRIVACY_VIOLATION";
    throw error;
  }

  const status = text(
    apiShadowObservation.status || shadowResolution.status || "FAILED_SAFE",
    60,
  ).toUpperCase();
  const candidates = extractCandidates(shadowResolution);
  const fallback = extractFallback(shadowResolution);
  const prediction = Object.freeze({
    candidates,
    domain: extractDomain(shadowResolution),
    intent: extractIntent(shadowResolution),
    fallbackApplied: fallback.applied,
    fallbackReason: fallback.reason,
    unsupportedRejected: extractUnsupportedRejected(
      shadowResolution,
      candidates,
    ),
    reviewRequired: extractReviewRequired(shadowResolution),
  });
  const requestFingerprintSha256 =
    safeSha256(apiShadowObservation.requestFingerprintSha256) ||
    stableSha256({ caseId: normalizedCaseId, status });

  return freezeDeep({
    version: "query_candidate_planner_shadow_accuracy_observation_v1",
    capturePolicyVersion: CAPTURE_POLICY_VERSION,
    observationId: stableSha256({
      caseId: normalizedCaseId,
      requestFingerprintSha256,
      status,
      prediction,
    }),
    caseId: normalizedCaseId,
    requestFingerprintSha256,
    status,
    reason: text(
      apiShadowObservation.reason || shadowResolution.status,
      120,
    ),
    primaryResponseUnchanged:
      apiShadowObservation.primaryResponseUnchanged !== false,
    latencyMs: Math.max(0, number(apiShadowObservation.latencyMs)),
    comparison: sanitizeComparison(apiShadowObservation.comparison),
    shadowPrediction: prediction,
    guardrails: sanitizeGuardrails(apiShadowObservation.guardrails),
    privacy: sanitizePrivacy(apiShadowObservation.privacy),
  });
}

function validateShadowAccuracyObservationDataset(
  dataset,
  accuracyDataset,
) {
  const errors = [];
  if (!isPlainObject(dataset)) {
    return freezeDeep({
      valid: false,
      errors: ["shadow observation dataset must be an object"],
    });
  }
  if (dataset.version !== OBSERVATION_DATASET_VERSION) {
    errors.push(
      `shadow observation dataset version must be ${OBSERVATION_DATASET_VERSION}`,
    );
  }
  if (!text(dataset.datasetId)) errors.push("datasetId is required");
  if (dataset.capturePolicyVersion !== CAPTURE_POLICY_VERSION) {
    errors.push(`capturePolicyVersion must be ${CAPTURE_POLICY_VERSION}`);
  }
  if (dataset.sourceAccuracyDatasetVersion !== ACCURACY_DATASET_VERSION) {
    errors.push(
      `sourceAccuracyDatasetVersion must be ${ACCURACY_DATASET_VERSION}`,
    );
  }
  if (
    text(dataset.sourceAccuracyDatasetId) !==
    text(accuracyDataset?.datasetId)
  ) {
    errors.push("sourceAccuracyDatasetId must match accuracy dataset");
  }
  if (!Array.isArray(dataset.observations) || dataset.observations.length === 0) {
    errors.push("observations must be a non-empty array");
  }

  const forbiddenPaths = findForbiddenPaths(dataset);
  for (const path of forbiddenPaths) {
    errors.push(`shadow observation dataset contains forbidden field: ${path}`);
  }
  for (const path of findAccuracySensitivePaths(dataset)) {
    errors.push(`shadow observation dataset contains sensitive field: ${path}`);
  }

  const knownCases = new Set(
    Array.isArray(accuracyDataset?.cases)
      ? accuracyDataset.cases.map((item) => text(item.caseId))
      : [],
  );
  const seenObservationIds = new Set();
  const seenCaseIds = new Set();
  for (const observation of Array.isArray(dataset.observations)
    ? dataset.observations
    : []) {
    if (!isPlainObject(observation)) {
      errors.push("observation must be an object");
      continue;
    }
    const observationId = text(observation.observationId);
    const caseId = text(observation.caseId);
    const status = text(observation.status).toUpperCase();
    if (!observationId || !SHA256_RE.test(observationId)) {
      errors.push(`${caseId || "unknown"}: valid observationId is required`);
    } else if (seenObservationIds.has(observationId)) {
      errors.push(`duplicate observationId: ${observationId}`);
    } else {
      seenObservationIds.add(observationId);
    }
    if (!caseId) {
      errors.push("observation caseId is required");
    } else if (!knownCases.has(caseId)) {
      errors.push(`unknown observation caseId: ${caseId}`);
    } else if (seenCaseIds.has(caseId)) {
      errors.push(`duplicate observation caseId: ${caseId}`);
    } else {
      seenCaseIds.add(caseId);
    }
    if (!Object.values(OBSERVATION_STATUSES).includes(status)) {
      errors.push(`${caseId}: unsupported observation status ${status}`);
    }
    if (!safeSha256(observation.requestFingerprintSha256)) {
      errors.push(`${caseId}: requestFingerprintSha256 is required`);
    }
    if (typeof observation.primaryResponseUnchanged !== "boolean") {
      errors.push(`${caseId}: primaryResponseUnchanged boolean is required`);
    }
    if (!Number.isFinite(Number(observation.latencyMs)) || Number(observation.latencyMs) < 0) {
      errors.push(`${caseId}: latencyMs must be non-negative`);
    }
    if (COMPLETED_STATUSES.has(status)) {
      if (!isPlainObject(observation.shadowPrediction)) {
        errors.push(`${caseId}: completed observation requires shadowPrediction`);
      } else {
        const prediction = observation.shadowPrediction;
        if (!Array.isArray(prediction.candidates)) {
          errors.push(`${caseId}: shadowPrediction.candidates must be array`);
        }
        if (!text(prediction.domain)) {
          errors.push(`${caseId}: shadowPrediction.domain is required`);
        }
        if (!text(prediction.intent)) {
          errors.push(`${caseId}: shadowPrediction.intent is required`);
        }
        if (typeof prediction.fallbackApplied !== "boolean") {
          errors.push(`${caseId}: fallbackApplied boolean is required`);
        }
        if (typeof prediction.unsupportedRejected !== "boolean") {
          errors.push(`${caseId}: unsupportedRejected boolean is required`);
        }
        if (typeof prediction.reviewRequired !== "boolean") {
          errors.push(`${caseId}: reviewRequired boolean is required`);
        }
        const candidateIds = [];
        const ranks = [];
        for (const [index, candidate] of (
          Array.isArray(prediction.candidates) ? prediction.candidates : []
        ).entries()) {
          const normalized = normalizeCandidate(candidate, index);
          if (!normalized.candidateId) {
            errors.push(`${caseId}: candidateId is required`);
          }
          candidateIds.push(normalized.candidateId);
          ranks.push(normalized.rank);
        }
        if (new Set(candidateIds).size !== candidateIds.length) {
          errors.push(`${caseId}: duplicate candidateId`);
        }
        if (new Set(ranks).size !== ranks.length) {
          errors.push(`${caseId}: duplicate candidate rank`);
        }
      }
    }
    if (!isPlainObject(observation.guardrails)) {
      errors.push(`${caseId}: guardrails are required`);
    }
    if (!isPlainObject(observation.privacy)) {
      errors.push(`${caseId}: privacy declaration is required`);
    }
  }

  return freezeDeep({
    valid: errors.length === 0,
    errors,
    observationCount: Array.isArray(dataset.observations)
      ? dataset.observations.length
      : 0,
    datasetSha256: sha256(dataset),
  });
}

function validateShadowAccuracyThresholdPolicy(policy) {
  const errors = [];
  if (!isPlainObject(policy)) {
    return freezeDeep({
      valid: false,
      errors: ["shadow accuracy threshold policy must be an object"],
    });
  }
  if (policy.version !== THRESHOLD_POLICY_VERSION) {
    errors.push(`threshold policy version must be ${THRESHOLD_POLICY_VERSION}`);
  }
  if (!Number.isInteger(policy.minimumObservationCount) ||
      policy.minimumObservationCount < 1) {
    errors.push("minimumObservationCount must be a positive integer");
  }
  if (typeof policy.requireAllAccuracyCases !== "boolean") {
    errors.push("requireAllAccuracyCases must be boolean");
  }
  if (typeof policy.requireAccuracyEvaluationPass !== "boolean") {
    errors.push("requireAccuracyEvaluationPass must be boolean");
  }
  const rateFields = [
    "minimumCompletedRate",
    "maximumBlockedRate",
    "maximumFailedSafeRate",
    "maximumTimeoutSafeRate",
    "minimumPrimaryResponseUnchangedRate",
    "minimumComparisonCoverage",
    "minimumPredictionCaptureCoverage",
  ];
  for (const field of rateFields) {
    const value = policy[field];
    if (!Number.isFinite(value) || value < 0 || value > 1) {
      errors.push(`${field} must be between 0 and 1`);
    }
  }
  for (const field of [
    "maximumGuardrailViolationCount",
    "maximumPrivacyViolationCount",
  ]) {
    const value = policy[field];
    if (!Number.isInteger(value) || value < 0) {
      errors.push(`${field} must be a non-negative integer`);
    }
  }
  return freezeDeep({ valid: errors.length === 0, errors });
}

function guardrailViolationCount(guardrails = {}) {
  return [
    guardrails.shadowOnly !== true,
    guardrails.primaryResponseAuthority !== true,
    guardrails.responsePayloadMutation === true,
    guardrails.responseHeaderMutation === true,
    guardrails.responseStatusMutation === true,
    guardrails.productionCandidateMerge === true,
    guardrails.productionReadyAssignment === true,
    guardrails.productionRouteChanged === true,
  ].filter(Boolean).length;
}

function privacyViolationCount(privacy = {}) {
  return [
    privacy.rawPrimaryResponseIncluded === true,
    privacy.rawShadowResolutionIncluded === true,
    privacy.rawRowsIncluded === true,
    privacy.sampleValuesIncluded === true,
    privacy.fileNameIncluded === true,
    privacy.originalFileNameIncluded === true,
    privacy.queryTablesKeyIncluded === true,
    privacy.userIdentityIncluded === true,
    privacy.tenantIdIncluded === true,
  ].filter(Boolean).length;
}

function observationsToPredictions(observations = []) {
  return Object.freeze(
    observations
      .filter((observation) =>
        COMPLETED_STATUSES.has(text(observation.status).toUpperCase()),
      )
      .map((observation) => {
        const prediction = observation.shadowPrediction || {};
        return Object.freeze({
          caseId: text(observation.caseId),
          candidates: Object.freeze(
            (Array.isArray(prediction.candidates)
              ? prediction.candidates
              : []
            ).map(normalizeCandidate),
          ),
          domain: text(prediction.domain) || "UNKNOWN",
          intent: text(prediction.intent) || "UNKNOWN",
          fallbackApplied: prediction.fallbackApplied === true,
          fallbackReason: text(prediction.fallbackReason, 120),
          unsupportedRejected:
            prediction.unsupportedRejected === true,
          reviewRequired: prediction.reviewRequired === true,
        });
      }),
  );
}

function summarizeShadowObservations(observations = []) {
  const total = observations.length;
  const statusCounts = {};
  const verdictCounts = {};
  let primaryResponseUnchangedCount = 0;
  let comparisonCount = 0;
  let predictionCaptureCount = 0;
  let guardrailViolations = 0;
  let privacyViolations = 0;
  let latencyTotalMs = 0;

  for (const observation of observations) {
    const status = text(observation.status).toUpperCase() || "UNKNOWN";
    statusCounts[status] = (statusCounts[status] || 0) + 1;
    if (observation.primaryResponseUnchanged === true) {
      primaryResponseUnchangedCount += 1;
    }
    if (isPlainObject(observation.comparison)) {
      comparisonCount += 1;
      const verdict = text(observation.comparison.verdict) || "NOT_AVAILABLE";
      verdictCounts[verdict] = (verdictCounts[verdict] || 0) + 1;
    }
    if (
      COMPLETED_STATUSES.has(status) &&
      isPlainObject(observation.shadowPrediction)
    ) {
      predictionCaptureCount += 1;
    }
    guardrailViolations += guardrailViolationCount(observation.guardrails);
    privacyViolations += privacyViolationCount(observation.privacy);
    latencyTotalMs += Math.max(0, number(observation.latencyMs));
  }

  const completedCount =
    (statusCounts[OBSERVATION_STATUSES.COMPLETED] || 0) +
    (statusCounts[OBSERVATION_STATUSES.COMPLETED_SAFE] || 0);
  const rate = (count) => (total > 0 ? round(count / total) : 0);

  return freezeDeep({
    observationCount: total,
    completedCount,
    blockedCount: statusCounts[OBSERVATION_STATUSES.BLOCKED] || 0,
    failedSafeCount: statusCounts[OBSERVATION_STATUSES.FAILED_SAFE] || 0,
    timeoutSafeCount: statusCounts[OBSERVATION_STATUSES.TIMEOUT_SAFE] || 0,
    comparisonCount,
    predictionCaptureCount,
    primaryResponseUnchangedCount,
    guardrailViolationCount: guardrailViolations,
    privacyViolationCount: privacyViolations,
    completedRate: rate(completedCount),
    blockedRate: rate(statusCounts[OBSERVATION_STATUSES.BLOCKED] || 0),
    failedSafeRate: rate(statusCounts[OBSERVATION_STATUSES.FAILED_SAFE] || 0),
    timeoutSafeRate: rate(statusCounts[OBSERVATION_STATUSES.TIMEOUT_SAFE] || 0),
    primaryResponseUnchangedRate: rate(primaryResponseUnchangedCount),
    comparisonCoverage: rate(comparisonCount),
    predictionCaptureCoverage: rate(predictionCaptureCount),
    averageLatencyMs: total > 0 ? round(latencyTotalMs / total, 3) : 0,
    statusCounts: Object.freeze({ ...statusCounts }),
    verdictCounts: Object.freeze({ ...verdictCounts }),
  });
}

function evaluateShadowThresholds({
  summary,
  policy,
  accuracyReport,
  accuracyCaseCount,
  observedCaseIds,
} = {}) {
  const checks = [];
  const push = (metric, operator, threshold, actual, passed) => {
    checks.push(Object.freeze({ metric, operator, threshold, actual, passed }));
  };
  push(
    "observationCount",
    ">=",
    policy.minimumObservationCount,
    summary.observationCount,
    summary.observationCount >= policy.minimumObservationCount,
  );
  push(
    "completedRate",
    ">=",
    policy.minimumCompletedRate,
    summary.completedRate,
    summary.completedRate >= policy.minimumCompletedRate,
  );
  push(
    "blockedRate",
    "<=",
    policy.maximumBlockedRate,
    summary.blockedRate,
    summary.blockedRate <= policy.maximumBlockedRate,
  );
  push(
    "failedSafeRate",
    "<=",
    policy.maximumFailedSafeRate,
    summary.failedSafeRate,
    summary.failedSafeRate <= policy.maximumFailedSafeRate,
  );
  push(
    "timeoutSafeRate",
    "<=",
    policy.maximumTimeoutSafeRate,
    summary.timeoutSafeRate,
    summary.timeoutSafeRate <= policy.maximumTimeoutSafeRate,
  );
  push(
    "primaryResponseUnchangedRate",
    ">=",
    policy.minimumPrimaryResponseUnchangedRate,
    summary.primaryResponseUnchangedRate,
    summary.primaryResponseUnchangedRate >=
      policy.minimumPrimaryResponseUnchangedRate,
  );
  push(
    "comparisonCoverage",
    ">=",
    policy.minimumComparisonCoverage,
    summary.comparisonCoverage,
    summary.comparisonCoverage >= policy.minimumComparisonCoverage,
  );
  push(
    "predictionCaptureCoverage",
    ">=",
    policy.minimumPredictionCaptureCoverage,
    summary.predictionCaptureCoverage,
    summary.predictionCaptureCoverage >=
      policy.minimumPredictionCaptureCoverage,
  );
  push(
    "guardrailViolationCount",
    "<=",
    policy.maximumGuardrailViolationCount,
    summary.guardrailViolationCount,
    summary.guardrailViolationCount <=
      policy.maximumGuardrailViolationCount,
  );
  push(
    "privacyViolationCount",
    "<=",
    policy.maximumPrivacyViolationCount,
    summary.privacyViolationCount,
    summary.privacyViolationCount <= policy.maximumPrivacyViolationCount,
  );
  const allCasesObserved = observedCaseIds.size === accuracyCaseCount;
  push(
    "accuracyCaseCoverage",
    policy.requireAllAccuracyCases ? "==" : ">=",
    policy.requireAllAccuracyCases ? accuracyCaseCount : 0,
    observedCaseIds.size,
    policy.requireAllAccuracyCases ? allCasesObserved : true,
  );
  push(
    "accuracyEvaluationDecision",
    "==",
    policy.requireAccuracyEvaluationPass
      ? ACCURACY_DECISIONS.PASS
      : "ANY",
    accuracyReport?.decision || "",
    policy.requireAccuracyEvaluationPass
      ? accuracyReport?.decision === ACCURACY_DECISIONS.PASS
      : true,
  );

  return freezeDeep({
    passed: checks.every((check) => check.passed),
    checks,
    failedMetrics: checks
      .filter((check) => !check.passed)
      .map((check) => check.metric),
  });
}

function reportGuardrails() {
  return freezeDeep({
    routeWired: false,
    controllerWired: false,
    internalPreviewWired: false,
    productionGateWired: false,
    promotionDecisionProduced: false,
    productionCandidateMergeApplied: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    providerCallsExecutedByEvaluator: 0,
    rawRowsStored: false,
    rawFileNameStored: false,
    userIdentityStored: false,
    evaluationOnly: true,
    failClosed: true,
  });
}

function blockedReport({
  shadowDataset,
  accuracyDataset,
  shadowPolicy,
  accuracyThresholdPolicy,
  shadowDatasetValidation,
  accuracyDatasetValidation,
  shadowPolicyValidation,
  reason,
} = {}) {
  return freezeDeep({
    version: REPORT_VERSION,
    evaluatorVersion: EVALUATOR_VERSION,
    observationDatasetVersion: shadowDataset?.version || "",
    observationDatasetId: shadowDataset?.datasetId || "",
    observationDatasetSha256: isPlainObject(shadowDataset)
      ? sha256(shadowDataset)
      : "",
    accuracyDatasetVersion: accuracyDataset?.version || "",
    accuracyDatasetId: accuracyDataset?.datasetId || "",
    accuracyDatasetSha256: isPlainObject(accuracyDataset)
      ? sha256(accuracyDataset)
      : "",
    shadowThresholdPolicyVersion: shadowPolicy?.version || "",
    shadowThresholdPolicySha256: isPlainObject(shadowPolicy)
      ? sha256(shadowPolicy)
      : "",
    accuracyThresholdPolicyVersion: accuracyThresholdPolicy?.version || "",
    accuracyThresholdPolicySha256: isPlainObject(accuracyThresholdPolicy)
      ? sha256(accuracyThresholdPolicy)
      : "",
    decision: DECISIONS.BLOCKED,
    reason,
    failClosed: true,
    evaluationOnly: true,
    promotionAuthorized: false,
    observationSummary: null,
    predictionCount: 0,
    accuracyReport: null,
    shadowThresholdEvaluation: null,
    invalid: Object.freeze({
      shadowDatasetErrors: shadowDatasetValidation?.errors || [],
      accuracyDatasetErrors: accuracyDatasetValidation?.errors || [],
      shadowThresholdPolicyErrors: shadowPolicyValidation?.errors || [],
    }),
    guardrails: reportGuardrails(),
  });
}

function evaluateShadowAccuracy({
  shadowObservationDataset,
  accuracyDataset,
  shadowThresholdPolicy,
  accuracyThresholdPolicy,
} = {}) {
  const shadowDatasetSnapshot = clone(shadowObservationDataset);
  const accuracyDatasetSnapshot = clone(accuracyDataset);
  const shadowPolicySnapshot = clone(shadowThresholdPolicy);
  const accuracyPolicySnapshot = clone(accuracyThresholdPolicy);

  const accuracyDatasetValidation = validateAccuracyEvaluationDataset(
    accuracyDatasetSnapshot,
  );
  const shadowDatasetValidation =
    validateShadowAccuracyObservationDataset(
      shadowDatasetSnapshot,
      accuracyDatasetSnapshot,
    );
  const shadowPolicyValidation = validateShadowAccuracyThresholdPolicy(
    shadowPolicySnapshot,
  );

  if (!accuracyDatasetValidation.valid) {
    return blockedReport({
      shadowDataset: shadowDatasetSnapshot,
      accuracyDataset: accuracyDatasetSnapshot,
      shadowPolicy: shadowPolicySnapshot,
      accuracyThresholdPolicy: accuracyPolicySnapshot,
      shadowDatasetValidation,
      accuracyDatasetValidation,
      shadowPolicyValidation,
      reason: "INVALID_ACCURACY_DATASET",
    });
  }
  if (!shadowDatasetValidation.valid) {
    return blockedReport({
      shadowDataset: shadowDatasetSnapshot,
      accuracyDataset: accuracyDatasetSnapshot,
      shadowPolicy: shadowPolicySnapshot,
      accuracyThresholdPolicy: accuracyPolicySnapshot,
      shadowDatasetValidation,
      accuracyDatasetValidation,
      shadowPolicyValidation,
      reason: "INVALID_SHADOW_OBSERVATION_DATASET",
    });
  }
  if (!shadowPolicyValidation.valid) {
    return blockedReport({
      shadowDataset: shadowDatasetSnapshot,
      accuracyDataset: accuracyDatasetSnapshot,
      shadowPolicy: shadowPolicySnapshot,
      accuracyThresholdPolicy: accuracyPolicySnapshot,
      shadowDatasetValidation,
      accuracyDatasetValidation,
      shadowPolicyValidation,
      reason: "INVALID_SHADOW_THRESHOLD_POLICY",
    });
  }

  const observations = shadowDatasetSnapshot.observations;
  const predictions = observationsToPredictions(observations);
  const accuracyReport = evaluateAccuracyDataset({
    dataset: accuracyDatasetSnapshot,
    predictions,
    thresholdPolicy: accuracyPolicySnapshot,
  });
  const summary = summarizeShadowObservations(observations);
  const observedCaseIds = new Set(
    observations.map((observation) => text(observation.caseId)),
  );
  const shadowThresholdEvaluation = evaluateShadowThresholds({
    summary,
    policy: shadowPolicySnapshot,
    accuracyReport,
    accuracyCaseCount: accuracyDatasetSnapshot.cases.length,
    observedCaseIds,
  });
  const passed =
    shadowThresholdEvaluation.passed &&
    (!shadowPolicySnapshot.requireAccuracyEvaluationPass ||
      accuracyReport.decision === ACCURACY_DECISIONS.PASS);

  const reportCore = {
    observationDatasetSha256: sha256(shadowDatasetSnapshot),
    accuracyDatasetSha256: sha256(accuracyDatasetSnapshot),
    shadowThresholdPolicySha256: sha256(shadowPolicySnapshot),
    accuracyThresholdPolicySha256: sha256(accuracyPolicySnapshot),
    observationSummary: summary,
    accuracyReportSha256: accuracyReport.reportSha256 || "",
    shadowThresholdEvaluation,
  };

  return freezeDeep({
    version: REPORT_VERSION,
    evaluatorVersion: EVALUATOR_VERSION,
    observationDatasetVersion: shadowDatasetSnapshot.version,
    observationDatasetId: shadowDatasetSnapshot.datasetId,
    observationDatasetSha256: reportCore.observationDatasetSha256,
    accuracyDatasetVersion: accuracyDatasetSnapshot.version,
    accuracyDatasetId: accuracyDatasetSnapshot.datasetId,
    accuracyDatasetSha256: reportCore.accuracyDatasetSha256,
    shadowThresholdPolicyVersion: shadowPolicySnapshot.version,
    shadowThresholdPolicySha256:
      reportCore.shadowThresholdPolicySha256,
    accuracyThresholdPolicyVersion: accuracyPolicySnapshot.version,
    accuracyThresholdPolicySha256:
      reportCore.accuracyThresholdPolicySha256,
    decision: passed ? DECISIONS.PASS : DECISIONS.BLOCKED,
    reason: passed
      ? "SHADOW_ACCURACY_THRESHOLDS_PASSED"
      : "SHADOW_ACCURACY_THRESHOLDS_NOT_MET",
    failClosed: true,
    evaluationOnly: true,
    promotionAuthorized: false,
    observationSummary: summary,
    predictionCount: predictions.length,
    accuracyReport,
    shadowThresholdEvaluation,
    invalid: Object.freeze({
      shadowDatasetErrors: [],
      accuracyDatasetErrors: [],
      shadowThresholdPolicyErrors: [],
    }),
    reportSha256: stableSha256(reportCore),
    guardrails: reportGuardrails(),
  });
}

module.exports = Object.freeze({
  EVALUATOR_VERSION,
  REPORT_VERSION,
  OBSERVATION_DATASET_VERSION,
  THRESHOLD_POLICY_VERSION,
  CAPTURE_POLICY_VERSION,
  DECISIONS,
  OBSERVATION_STATUSES,
  FORBIDDEN_KEYS,
  buildShadowAccuracyObservation,
  findForbiddenPaths,
  validateShadowAccuracyObservationDataset,
  validateShadowAccuracyThresholdPolicy,
  observationsToPredictions,
  summarizeShadowObservations,
  evaluateShadowThresholds,
  evaluateShadowAccuracy,
});
