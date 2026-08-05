"use strict";

const crypto = require("crypto");

const EVALUATOR_VERSION =
  "query_candidate_planner_accuracy_evaluator_v1";
const REPORT_VERSION =
  "query_candidate_planner_accuracy_evaluation_report_v1";
const DATASET_VERSION =
  "query_candidate_planner_accuracy_evaluation_dataset_v1";
const THRESHOLD_POLICY_VERSION =
  "query_candidate_planner_accuracy_threshold_policy_v1";
const DECISIONS = Object.freeze({
  PASS: "EVALUATION_PASS",
  BLOCKED: "EVALUATION_BLOCKED",
});
const CANDIDATE_STATUSES = Object.freeze({
  ACCEPTED: "ACCEPTED",
  REJECTED: "REJECTED",
});
const SHA256_RE = /^[a-f0-9]{64}$/i;
const SENSITIVE_KEYS = new Set([
  "rows",
  "rawRows",
  "sampleValues",
  "fileName",
  "originalFileName",
  "email",
  "userId",
  "tenantId",
  "queryTablesKey",
  "storageKey",
  "rawPayload",
]);

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, canonicalize(value[key])]),
  );
}

function canonicalJson(value) {
  return JSON.stringify(canonicalize(value));
}

function sha256(value) {
  const serialized = typeof value === "string" ? value : canonicalJson(value);
  return crypto.createHash("sha256").update(serialized).digest("hex");
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

function round(value, digits = 6) {
  if (!Number.isFinite(value)) return 0;
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function mean(values, fallback = 0) {
  const finite = values.filter(Number.isFinite);
  if (finite.length === 0) return fallback;
  return finite.reduce((sum, value) => sum + value, 0) / finite.length;
}

function normalizeString(value) {
  return typeof value === "string" ? value.trim() : "";
}

function uniqueStrings(values) {
  const seen = new Set();
  const result = [];
  for (const value of Array.isArray(values) ? values : []) {
    const normalized = normalizeString(value);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

function findSensitivePaths(value, basePath = "$") {
  const paths = [];
  if (Array.isArray(value)) {
    value.forEach((entry, index) => {
      paths.push(...findSensitivePaths(entry, `${basePath}[${index}]`));
    });
    return paths;
  }
  if (!isPlainObject(value)) return paths;
  for (const [key, entry] of Object.entries(value)) {
    const childPath = `${basePath}.${key}`;
    if (SENSITIVE_KEYS.has(key)) paths.push(childPath);
    paths.push(...findSensitivePaths(entry, childPath));
  }
  return paths;
}

function validateCandidateLabels(labels, caseId, errors) {
  const required = Array.isArray(labels.requiredCandidates)
    ? labels.requiredCandidates
    : [];
  const acceptable = uniqueStrings(labels.acceptableCandidateIds);
  const forbidden = uniqueStrings(labels.forbiddenCandidateIds);
  const preferredTop1 = uniqueStrings(labels.preferredTop1CandidateIds);

  if (!Array.isArray(labels.requiredCandidates)) {
    errors.push(`${caseId}: requiredCandidates must be an array`);
  }
  const requiredIds = [];
  const idealRanks = new Set();
  for (const candidate of required) {
    if (!isPlainObject(candidate)) {
      errors.push(`${caseId}: required candidate must be an object`);
      continue;
    }
    const candidateId = normalizeString(candidate.candidateId);
    if (!candidateId) {
      errors.push(`${caseId}: required candidateId is required`);
      continue;
    }
    requiredIds.push(candidateId);
    if (!Number.isInteger(candidate.idealRank) || candidate.idealRank < 1) {
      errors.push(`${caseId}: ${candidateId} idealRank must be positive integer`);
    } else if (idealRanks.has(candidate.idealRank)) {
      errors.push(`${caseId}: duplicate idealRank ${candidate.idealRank}`);
    } else {
      idealRanks.add(candidate.idealRank);
    }
    if (![1, 2, 3].includes(candidate.relevance)) {
      errors.push(`${caseId}: ${candidateId} relevance must be 1, 2, or 3`);
    }
  }

  if (new Set(requiredIds).size !== requiredIds.length) {
    errors.push(`${caseId}: duplicate required candidateId`);
  }
  if (acceptable.length !== (labels.acceptableCandidateIds || []).length) {
    errors.push(`${caseId}: acceptableCandidateIds must be unique non-empty strings`);
  }
  if (forbidden.length !== (labels.forbiddenCandidateIds || []).length) {
    errors.push(`${caseId}: forbiddenCandidateIds must be unique non-empty strings`);
  }
  if (preferredTop1.length !== (labels.preferredTop1CandidateIds || []).length) {
    errors.push(`${caseId}: preferredTop1CandidateIds must be unique non-empty strings`);
  }

  const requiredSet = new Set(requiredIds);
  const acceptableSet = new Set(acceptable);
  const forbiddenSet = new Set(forbidden);
  for (const id of acceptableSet) {
    if (requiredSet.has(id)) {
      errors.push(`${caseId}: candidate ${id} cannot be required and acceptable`);
    }
  }
  for (const id of forbiddenSet) {
    if (requiredSet.has(id) || acceptableSet.has(id)) {
      errors.push(`${caseId}: forbidden candidate overlaps relevant labels: ${id}`);
    }
  }
  const relevantSet = new Set([...requiredSet, ...acceptableSet]);
  for (const id of preferredTop1) {
    if (!relevantSet.has(id)) {
      errors.push(`${caseId}: preferred top-1 candidate is not relevant: ${id}`);
    }
  }
  if (requiredIds.length > 0 && preferredTop1.length === 0) {
    errors.push(`${caseId}: preferredTop1CandidateIds required for supported case`);
  }
}

function validateAccuracyEvaluationDataset(dataset) {
  const errors = [];
  if (!isPlainObject(dataset)) {
    return freezeDeep({ valid: false, errors: ["dataset must be an object"] });
  }
  if (dataset.version !== DATASET_VERSION) {
    errors.push(`dataset version must be ${DATASET_VERSION}`);
  }
  if (!normalizeString(dataset.datasetId)) errors.push("datasetId is required");
  if (!Array.isArray(dataset.cases) || dataset.cases.length === 0) {
    errors.push("dataset cases must be a non-empty array");
  }
  const sensitivePaths = findSensitivePaths(dataset);
  for (const path of sensitivePaths) {
    errors.push(`dataset contains forbidden sensitive field: ${path}`);
  }

  const caseIds = new Set();
  for (const item of Array.isArray(dataset.cases) ? dataset.cases : []) {
    if (!isPlainObject(item)) {
      errors.push("dataset case must be an object");
      continue;
    }
    const caseId = normalizeString(item.caseId);
    if (!caseId) {
      errors.push("caseId is required");
      continue;
    }
    if (caseIds.has(caseId)) errors.push(`duplicate caseId: ${caseId}`);
    caseIds.add(caseId);
    if (!normalizeString(item.sourceRef)) {
      errors.push(`${caseId}: sourceRef is required`);
    }
    if (!Array.isArray(item.tags) || item.tags.length === 0) {
      errors.push(`${caseId}: tags must be a non-empty array`);
    }
    if (!isPlainObject(item.labels)) {
      errors.push(`${caseId}: labels are required`);
      continue;
    }
    validateCandidateLabels(item.labels, caseId, errors);
    const domain = item.labels.domain;
    const intent = item.labels.intent;
    if (!isPlainObject(domain) || !normalizeString(domain.expected)) {
      errors.push(`${caseId}: domain.expected is required`);
    }
    if (!isPlainObject(intent) || !normalizeString(intent.expected)) {
      errors.push(`${caseId}: intent.expected is required`);
    }
    if (!isPlainObject(item.labels.fallback) ||
        typeof item.labels.fallback.expected !== "boolean") {
      errors.push(`${caseId}: fallback.expected boolean is required`);
    }
    if (!isPlainObject(item.labels.unsupported) ||
        typeof item.labels.unsupported.expectedRejected !== "boolean") {
      errors.push(`${caseId}: unsupported.expectedRejected boolean is required`);
    }
    if (typeof item.labels.reviewRequired !== "boolean") {
      errors.push(`${caseId}: reviewRequired boolean is required`);
    }
    const requiredCount = Array.isArray(item.labels.requiredCandidates)
      ? item.labels.requiredCandidates.length
      : 0;
    if (item.labels.unsupported?.expectedRejected && requiredCount > 0) {
      errors.push(`${caseId}: unsupported case cannot require candidates`);
    }
  }

  return freezeDeep({
    valid: errors.length === 0,
    errors,
    caseCount: caseIds.size,
    datasetSha256: sha256(dataset),
  });
}

function validateThresholdPolicy(policy) {
  const errors = [];
  if (!isPlainObject(policy)) {
    return freezeDeep({ valid: false, errors: ["threshold policy must be an object"] });
  }
  if (policy.version !== THRESHOLD_POLICY_VERSION) {
    errors.push(`threshold policy version must be ${THRESHOLD_POLICY_VERSION}`);
  }
  if (!Number.isInteger(policy.minimumCaseCount) || policy.minimumCaseCount < 1) {
    errors.push("minimumCaseCount must be a positive integer");
  }
  if (typeof policy.requireAllCases !== "boolean") {
    errors.push("requireAllCases must be boolean");
  }
  if (!isPlainObject(policy.metricWeights)) {
    errors.push("metricWeights are required");
  }
  if (!isPlainObject(policy.thresholds)) {
    errors.push("thresholds are required");
  }
  const weights = Object.values(policy.metricWeights || {});
  if (weights.some((value) => !Number.isFinite(value) || value < 0)) {
    errors.push("metricWeights must be non-negative numbers");
  }
  if (weights.length > 0 && Math.abs(weights.reduce((a, b) => a + b, 0) - 1) > 1e-9) {
    errors.push("metricWeights must sum to 1");
  }
  return freezeDeep({ valid: errors.length === 0, errors });
}

function normalizePredictionCandidate(candidate, index) {
  const candidateId = normalizeString(candidate?.candidateId);
  const rank = Number.isInteger(candidate?.rank) && candidate.rank > 0
    ? candidate.rank
    : index + 1;
  const status = candidate?.status === CANDIDATE_STATUSES.REJECTED
    ? CANDIDATE_STATUSES.REJECTED
    : CANDIDATE_STATUSES.ACCEPTED;
  return {
    candidateId,
    rank,
    status,
    productionEligible: candidate?.productionEligible !== false,
  };
}

function validatePredictions(predictions, datasetCaseIds) {
  const errors = [];
  if (!Array.isArray(predictions)) {
    return freezeDeep({ valid: false, errors: ["predictions must be an array"] });
  }
  const seenCases = new Set();
  for (const prediction of predictions) {
    if (!isPlainObject(prediction)) {
      errors.push("prediction must be an object");
      continue;
    }
    const caseId = normalizeString(prediction.caseId);
    if (!caseId) {
      errors.push("prediction caseId is required");
      continue;
    }
    if (seenCases.has(caseId)) errors.push(`duplicate prediction caseId: ${caseId}`);
    seenCases.add(caseId);
    if (!datasetCaseIds.has(caseId)) errors.push(`unknown prediction caseId: ${caseId}`);
    if (!Array.isArray(prediction.candidates)) {
      errors.push(`${caseId}: candidates must be an array`);
      continue;
    }
    const candidateIds = [];
    const ranks = [];
    prediction.candidates.forEach((candidate, index) => {
      const normalized = normalizePredictionCandidate(candidate, index);
      if (!normalized.candidateId) {
        errors.push(`${caseId}: candidateId is required`);
      } else {
        candidateIds.push(normalized.candidateId);
      }
      ranks.push(normalized.rank);
    });
    if (new Set(candidateIds).size !== candidateIds.length) {
      errors.push(`${caseId}: duplicate candidateId in prediction`);
    }
    if (new Set(ranks).size !== ranks.length) {
      errors.push(`${caseId}: duplicate candidate rank`);
    }
    if (!normalizeString(prediction.domain)) errors.push(`${caseId}: domain is required`);
    if (!normalizeString(prediction.intent)) errors.push(`${caseId}: intent is required`);
    if (typeof prediction.fallbackApplied !== "boolean") {
      errors.push(`${caseId}: fallbackApplied boolean is required`);
    }
    if (typeof prediction.unsupportedRejected !== "boolean") {
      errors.push(`${caseId}: unsupportedRejected boolean is required`);
    }
    if (typeof prediction.reviewRequired !== "boolean") {
      errors.push(`${caseId}: reviewRequired boolean is required`);
    }
  }
  const sensitivePaths = findSensitivePaths(predictions);
  for (const path of sensitivePaths) {
    errors.push(`predictions contain forbidden sensitive field: ${path}`);
  }
  return freezeDeep({ valid: errors.length === 0, errors });
}

function promotedCandidates(prediction) {
  return (prediction.candidates || [])
    .map(normalizePredictionCandidate)
    .filter((candidate) =>
      candidate.candidateId &&
      candidate.status !== CANDIDATE_STATUSES.REJECTED &&
      candidate.productionEligible,
    )
    .sort((left, right) => left.rank - right.rank || left.candidateId.localeCompare(right.candidateId));
}

function reciprocalMetric(value) {
  return round(Math.max(0, Math.min(1, value)));
}

function dcg(relevances) {
  return relevances.reduce((sum, relevance, index) => {
    const gain = (2 ** relevance) - 1;
    return sum + gain / Math.log2(index + 2);
  }, 0);
}

function rankingAgreement(labels, candidates) {
  const required = [...(labels.requiredCandidates || [])]
    .sort((a, b) => a.idealRank - b.idealRank);
  if (required.length <= 1) {
    return { applicable: false, value: 1 };
  }
  const positionById = new Map(
    candidates.map((candidate, index) => [candidate.candidateId, index]),
  );
  let concordant = 0;
  let pairCount = 0;
  for (let left = 0; left < required.length; left += 1) {
    for (let right = left + 1; right < required.length; right += 1) {
      pairCount += 1;
      const leftPosition = positionById.get(required[left].candidateId);
      const rightPosition = positionById.get(required[right].candidateId);
      if (Number.isInteger(leftPosition) &&
          Number.isInteger(rightPosition) &&
          leftPosition < rightPosition) {
        concordant += 1;
      }
    }
  }
  return {
    applicable: pairCount > 0,
    value: pairCount > 0 ? reciprocalMetric(concordant / pairCount) : 1,
  };
}

function evaluateCase(item, prediction, policy) {
  const labels = item.labels;
  const candidates = promotedCandidates(prediction);
  const predictedIds = candidates.map((candidate) => candidate.candidateId);
  const predictedSet = new Set(predictedIds);
  const required = [...(labels.requiredCandidates || [])]
    .sort((a, b) => a.idealRank - b.idealRank);
  const requiredIds = required.map((candidate) => candidate.candidateId);
  const requiredSet = new Set(requiredIds);
  const acceptableSet = new Set(labels.acceptableCandidateIds || []);
  const forbiddenSet = new Set(labels.forbiddenCandidateIds || []);
  const relevantSet = new Set([...requiredSet, ...acceptableSet]);

  const relevantHits = predictedIds.filter((id) => relevantSet.has(id)).length;
  const requiredHits = requiredIds.filter((id) => predictedSet.has(id)).length;
  const forbiddenHits = predictedIds.filter((id) => forbiddenSet.has(id)).length;
  const candidatePrecision = candidates.length === 0
    ? (requiredIds.length === 0 ? 1 : 0)
    : relevantHits / candidates.length;
  const candidateRecall = requiredIds.length === 0 ? 1 : requiredHits / requiredIds.length;
  const preferredTop1 = new Set(labels.preferredTop1CandidateIds || []);
  const top1Applicable = preferredTop1.size > 0;
  const top1Accuracy = top1Applicable && candidates[0]
    ? (preferredTop1.has(candidates[0].candidateId) ? 1 : 0)
    : (top1Applicable ? 0 : 1);
  const topK = Number.isInteger(policy.topK) && policy.topK > 0 ? policy.topK : 3;
  const topKIds = new Set(candidates.slice(0, topK).map((candidate) => candidate.candidateId));
  const topKApplicable = requiredIds.length > 0;
  const topKRecall = topKApplicable
    ? requiredIds.filter((id) => topKIds.has(id)).length / requiredIds.length
    : 1;
  const ranking = rankingAgreement(labels, candidates);

  const acceptedDomains = new Set([
    normalizeString(labels.domain.expected),
    ...uniqueStrings(labels.domain.acceptable),
  ]);
  const acceptedIntents = new Set([
    normalizeString(labels.intent.expected),
    ...uniqueStrings(labels.intent.acceptable),
  ]);
  const domainAccuracy = acceptedDomains.has(normalizeString(prediction.domain)) ? 1 : 0;
  const intentAccuracy = acceptedIntents.has(normalizeString(prediction.intent)) ? 1 : 0;

  const fallbackLabel = labels.fallback;
  let fallbackAccuracy = prediction.fallbackApplied === fallbackLabel.expected ? 1 : 0;
  if (fallbackAccuracy === 1 && fallbackLabel.expected) {
    const acceptableReasons = uniqueStrings(fallbackLabel.acceptableReasons);
    if (acceptableReasons.length > 0 &&
        !acceptableReasons.includes(normalizeString(prediction.fallbackReason))) {
      fallbackAccuracy = 0;
    }
  }

  const unsupportedExpected = labels.unsupported.expectedRejected;
  const unsupportedRejectionAccuracy = unsupportedExpected
    ? (prediction.unsupportedRejected && candidates.length === 0 ? 1 : 0)
    : (!prediction.unsupportedRejected ? 1 : 0);
  const reviewDecisionAccuracy = prediction.reviewRequired === labels.reviewRequired ? 1 : 0;
  const falsePromotionRate = candidates.length === 0 ? 0 : forbiddenHits / candidates.length;
  const falsePromotionSafety = 1 - falsePromotionRate;

  const metricValues = {
    candidatePrecision: reciprocalMetric(candidatePrecision),
    candidateRecall: reciprocalMetric(candidateRecall),
    top1Accuracy: reciprocalMetric(top1Accuracy),
    topKRecall: reciprocalMetric(topKRecall),
    rankingAgreement: reciprocalMetric(ranking.value),
    domainAccuracy: reciprocalMetric(domainAccuracy),
    intentAccuracy: reciprocalMetric(intentAccuracy),
    fallbackAccuracy: reciprocalMetric(fallbackAccuracy),
    unsupportedRejectionAccuracy: reciprocalMetric(unsupportedRejectionAccuracy),
    reviewDecisionAccuracy: reciprocalMetric(reviewDecisionAccuracy),
    falsePromotionRate: reciprocalMetric(falsePromotionRate),
    falsePromotionSafety: reciprocalMetric(falsePromotionSafety),
  };
  const weightedScore = round(Object.entries(policy.metricWeights).reduce(
    (sum, [metric, weight]) => sum + (metricValues[metric] || 0) * weight,
    0,
  ));

  return freezeDeep({
    caseId: item.caseId,
    sourceRef: item.sourceRef,
    caseWeight: Number.isFinite(item.caseWeight) ? item.caseWeight : 1,
    predictedCandidateCount: candidates.length,
    requiredCandidateCount: requiredIds.length,
    relevantHitCount: relevantHits,
    requiredHitCount: requiredHits,
    forbiddenHitCount: forbiddenHits,
    metrics: metricValues,
    applicability: {
      top1Accuracy: top1Applicable,
      topKRecall: topKApplicable,
      rankingAgreement: ranking.applicable,
    },
    weightedScore,
    candidateOrderSha256: sha256(predictedIds),
    labelsSha256: sha256(labels),
  });
}

function aggregateCases(caseReports, policy) {
  const aggregate = {};
  const metricNames = [
    "candidatePrecision",
    "candidateRecall",
    "top1Accuracy",
    "topKRecall",
    "rankingAgreement",
    "domainAccuracy",
    "intentAccuracy",
    "fallbackAccuracy",
    "unsupportedRejectionAccuracy",
    "reviewDecisionAccuracy",
    "falsePromotionRate",
    "falsePromotionSafety",
  ];
  for (const metricName of metricNames) {
    const values = [];
    for (const report of caseReports) {
      if (report.applicability[metricName] === false) continue;
      values.push(report.metrics[metricName]);
    }
    aggregate[metricName] = round(mean(values, metricName === "falsePromotionRate" ? 0 : 1));
  }
  aggregate.overallScore = round(Object.entries(policy.metricWeights).reduce(
    (sum, [metric, weight]) => sum + (aggregate[metric] || 0) * weight,
    0,
  ));
  return freezeDeep(aggregate);
}

function evaluateThresholds({ aggregate, caseCount, predictionCount, missingCaseIds, policy }) {
  const checks = [];
  const pushMin = (metric, threshold) => {
    checks.push({
      metric,
      operator: ">=",
      threshold,
      actual: aggregate[metric],
      passed: aggregate[metric] >= threshold,
    });
  };
  const pushMax = (metric, threshold) => {
    checks.push({
      metric,
      operator: "<=",
      threshold,
      actual: aggregate[metric],
      passed: aggregate[metric] <= threshold,
    });
  };
  for (const [metric, threshold] of Object.entries(policy.thresholds.minimum || {})) {
    pushMin(metric, threshold);
  }
  for (const [metric, threshold] of Object.entries(policy.thresholds.maximum || {})) {
    pushMax(metric, threshold);
  }
  checks.push({
    metric: "caseCount",
    operator: ">=",
    threshold: policy.minimumCaseCount,
    actual: caseCount,
    passed: caseCount >= policy.minimumCaseCount,
  });
  checks.push({
    metric: "predictionCoverage",
    operator: policy.requireAllCases ? "==" : ">=",
    threshold: policy.requireAllCases ? caseCount : 0,
    actual: predictionCount,
    passed: policy.requireAllCases ? missingCaseIds.length === 0 : true,
  });
  return freezeDeep({
    passed: checks.every((check) => check.passed),
    checks,
    failedMetrics: checks.filter((check) => !check.passed).map((check) => check.metric),
  });
}

function blockedReport({ dataset, policy, datasetValidation, policyValidation, predictionValidation, reason }) {
  return freezeDeep({
    version: REPORT_VERSION,
    evaluatorVersion: EVALUATOR_VERSION,
    datasetVersion: dataset?.version || "",
    datasetId: dataset?.datasetId || "",
    datasetSha256: isPlainObject(dataset) ? sha256(dataset) : "",
    thresholdPolicyVersion: policy?.version || "",
    thresholdPolicySha256: isPlainObject(policy) ? sha256(policy) : "",
    decision: DECISIONS.BLOCKED,
    reason,
    failClosed: true,
    evaluationOnly: true,
    promotionAuthorized: false,
    caseCount: Array.isArray(dataset?.cases) ? dataset.cases.length : 0,
    predictionCount: 0,
    missingCaseIds: [],
    invalid: {
      datasetErrors: datasetValidation?.errors || [],
      thresholdPolicyErrors: policyValidation?.errors || [],
      predictionErrors: predictionValidation?.errors || [],
    },
    aggregate: null,
    thresholdEvaluation: null,
    cases: [],
    guardrails: guardrails(),
  });
}

function guardrails() {
  return freezeDeep({
    routeWired: false,
    controllerWired: false,
    productionGateWired: false,
    promotionDecisionProduced: false,
    productionCandidateMergeApplied: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    providerCalls: 0,
    rawRowsStored: false,
    rawFileNameStored: false,
    userIdentityStored: false,
    evaluationOnly: true,
    failClosed: true,
  });
}

function evaluateAccuracyDataset({ dataset, predictions, thresholdPolicy } = {}) {
  const datasetSnapshot = clone(dataset);
  const predictionSnapshot = clone(predictions);
  const policySnapshot = clone(thresholdPolicy);
  const datasetValidation = validateAccuracyEvaluationDataset(datasetSnapshot);
  const policyValidation = validateThresholdPolicy(policySnapshot);
  if (!datasetValidation.valid || !policyValidation.valid) {
    return blockedReport({
      dataset: datasetSnapshot,
      policy: policySnapshot,
      datasetValidation,
      policyValidation,
      reason: !datasetValidation.valid
        ? "INVALID_EVALUATION_DATASET"
        : "INVALID_THRESHOLD_POLICY",
    });
  }

  const datasetCaseIds = new Set(datasetSnapshot.cases.map((item) => item.caseId));
  const predictionValidation = validatePredictions(predictionSnapshot, datasetCaseIds);
  if (!predictionValidation.valid) {
    return blockedReport({
      dataset: datasetSnapshot,
      policy: policySnapshot,
      datasetValidation,
      policyValidation,
      predictionValidation,
      reason: "INVALID_PREDICTIONS",
    });
  }

  const predictionByCaseId = new Map(predictionSnapshot.map((item) => [item.caseId, item]));
  const missingCaseIds = datasetSnapshot.cases
    .map((item) => item.caseId)
    .filter((caseId) => !predictionByCaseId.has(caseId));
  const caseReports = datasetSnapshot.cases.map((item) => {
    const prediction = predictionByCaseId.get(item.caseId) || {
      caseId: item.caseId,
      candidates: [],
      domain: "MISSING",
      intent: "MISSING",
      fallbackApplied: false,
      fallbackReason: "",
      unsupportedRejected: false,
      reviewRequired: true,
    };
    return evaluateCase(item, prediction, policySnapshot);
  });
  const aggregate = aggregateCases(caseReports, policySnapshot);
  const thresholdEvaluation = evaluateThresholds({
    aggregate,
    caseCount: datasetSnapshot.cases.length,
    predictionCount: predictionSnapshot.length,
    missingCaseIds,
    policy: policySnapshot,
  });
  const passed = thresholdEvaluation.passed && missingCaseIds.length === 0;

  return freezeDeep({
    version: REPORT_VERSION,
    evaluatorVersion: EVALUATOR_VERSION,
    datasetVersion: datasetSnapshot.version,
    datasetId: datasetSnapshot.datasetId,
    datasetSha256: sha256(datasetSnapshot),
    thresholdPolicyVersion: policySnapshot.version,
    thresholdPolicySha256: sha256(policySnapshot),
    decision: passed ? DECISIONS.PASS : DECISIONS.BLOCKED,
    reason: passed ? "ALL_ACCURACY_THRESHOLDS_PASSED" : "ACCURACY_THRESHOLDS_NOT_MET",
    failClosed: true,
    evaluationOnly: true,
    promotionAuthorized: false,
    caseCount: datasetSnapshot.cases.length,
    predictionCount: predictionSnapshot.length,
    missingCaseIds,
    invalid: {
      datasetErrors: [],
      thresholdPolicyErrors: [],
      predictionErrors: [],
    },
    aggregate,
    thresholdEvaluation,
    cases: caseReports,
    reportSha256: sha256({
      datasetSha256: sha256(datasetSnapshot),
      thresholdPolicySha256: sha256(policySnapshot),
      aggregate,
      thresholdEvaluation,
      cases: caseReports,
    }),
    guardrails: guardrails(),
  });
}

module.exports = Object.freeze({
  EVALUATOR_VERSION,
  REPORT_VERSION,
  DATASET_VERSION,
  THRESHOLD_POLICY_VERSION,
  DECISIONS,
  CANDIDATE_STATUSES,
  canonicalJson,
  sha256,
  findSensitivePaths,
  validateAccuracyEvaluationDataset,
  validateThresholdPolicy,
  validatePredictions,
  evaluateAccuracyDataset,
});
