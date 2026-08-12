const crypto = require("crypto");

const BASELINE_VERSION =
  "query_candidate_planner_real_shadow_evaluation_baseline_v1";
const BASELINE_DECISION = "EVALUATION_BASELINE_PASS";

const DEFAULT_POLICY = Object.freeze({
  version: "query_candidate_planner_real_shadow_evaluation_baseline_policy_v1",
  minimumExecutionCount: 50,
  minimumLifecycleCount: 20,
  minimumCaseCount: 10,
  requireCollectionProtocolComplete: true,
  requireApprovedActualPricing: true,
  requireOperationalEvaluationPass: true,
  requireEvaluationOnly: true,
  requirePromotionBlocked: true,
  maximumPrivacyViolationCount: 0,
  maximumGuardrailViolationCount: 0,
});

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function stableValue(value) {
  if (Array.isArray(value)) return value.map(stableValue);
  if (!isObject(value)) return value;
  const out = {};
  for (const key of Object.keys(value).sort())
    out[key] = stableValue(value[key]);
  return out;
}

function stableStringify(value) {
  return JSON.stringify(stableValue(value));
}

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(stableStringify(value))
    .digest("hex");
}

function walk(value, visitor, path = []) {
  if (Array.isArray(value)) {
    value.forEach((item, index) => walk(item, visitor, path.concat(index)));
    return;
  }
  if (!isObject(value)) return;
  for (const [key, child] of Object.entries(value)) {
    visitor(key, child, path.concat(key));
    if (isObject(child) || Array.isArray(child))
      walk(child, visitor, path.concat(key));
  }
}

function findScalarByKeys(value, keys) {
  const wanted = new Set(keys.map((key) => String(key).toLowerCase()));
  let found;
  walk(value, (key, child) => {
    if (found !== undefined) return;
    if (!wanted.has(String(key).toLowerCase())) return;
    if (
      child === null ||
      ["string", "number", "boolean"].includes(typeof child)
    ) {
      found = child;
    }
  });
  return found;
}

function findNumber(value, keys, fallback = null) {
  const raw = findScalarByKeys(value, keys);
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function findBoolean(value, keys, fallback = null) {
  const raw = findScalarByKeys(value, keys);
  if (raw === true || raw === false) return raw;
  if (String(raw).toLowerCase() === "true") return true;
  if (String(raw).toLowerCase() === "false") return false;
  return fallback;
}

function findString(value, keys, fallback = "") {
  const raw = findScalarByKeys(value, keys);
  return raw == null ? fallback : String(raw).trim();
}

function assertCondition(condition, code, message) {
  if (!condition) {
    const error = new Error(message || code);
    error.code = code;
    throw error;
  }
}

function collectionFacts(summary = {}) {
  const executionCount = findNumber(summary, [
    "actualExecutionCount",
    "executionCount",
    "executionTotal",
    "executions",
  ]);
  const lifecycleCount = findNumber(summary, [
    "actualLifecycleCount",
    "lifecycleCount",
    "lifecycleTotal",
    "lifecycle",
  ]);
  const caseCount = findNumber(
    summary,
    ["readyCaseCount", "registryCaseCount", "caseCount", "casesComplete"],
    10,
  );
  const privacyViolationCount = findNumber(
    summary,
    ["privacyViolationCount", "privacyViolations"],
    0,
  );
  const guardrailViolationCount = findNumber(
    summary,
    ["guardrailViolationCount", "guardrailViolations"],
    0,
  );
  const protocolComplete = findBoolean(
    summary,
    ["collectionProtocolComplete", "protocolComplete"],
    false,
  );
  const readyForPatchF = findBoolean(
    summary,
    ["readyForPatch15_3_2_F", "readyForPatch15_3_2_f"],
    false,
  );
  const from = findString(summary, [
    "from",
    "windowFrom",
    "collectionFrom",
    "startedAt",
  ]);
  const to = findString(summary, [
    "to",
    "windowTo",
    "collectionTo",
    "finalizedAt",
    "endedAt",
  ]);

  return Object.freeze({
    executionCount,
    lifecycleCount,
    caseCount,
    privacyViolationCount,
    guardrailViolationCount,
    protocolComplete,
    readyForPatchF,
    from,
    to,
  });
}

function validateCollectionSummary(summary = {}, policy = DEFAULT_POLICY) {
  const facts = collectionFacts(summary);
  assertCondition(
    facts.readyForPatchF === true,
    "PATCH_E_NOT_READY_FOR_F",
    "Patch E summary must explicitly authorize transition to Patch 15.3.2-F.",
  );
  assertCondition(
    facts.protocolComplete === true,
    "PATCH_E_PROTOCOL_INCOMPLETE",
    "Patch E collection protocol is incomplete.",
  );
  assertCondition(
    Number.isFinite(facts.executionCount) &&
      facts.executionCount >= policy.minimumExecutionCount,
    "PATCH_E_EXECUTION_SAMPLE_TOO_SMALL",
    `Execution count must be >= ${policy.minimumExecutionCount}.`,
  );
  assertCondition(
    Number.isFinite(facts.lifecycleCount) &&
      facts.lifecycleCount >= policy.minimumLifecycleCount,
    "PATCH_E_LIFECYCLE_SAMPLE_TOO_SMALL",
    `Lifecycle count must be >= ${policy.minimumLifecycleCount}.`,
  );
  assertCondition(
    Number.isFinite(facts.caseCount) &&
      facts.caseCount >= policy.minimumCaseCount,
    "PATCH_E_CASE_COVERAGE_TOO_SMALL",
    `Case count must be >= ${policy.minimumCaseCount}.`,
  );
  assertCondition(
    facts.privacyViolationCount <= policy.maximumPrivacyViolationCount,
    "PATCH_E_PRIVACY_VIOLATION",
    "Patch E privacy violations must be zero.",
  );
  assertCondition(
    facts.guardrailViolationCount <= policy.maximumGuardrailViolationCount,
    "PATCH_E_GUARDRAIL_VIOLATION",
    "Patch E guardrail violations must be zero.",
  );
  assertCondition(
    Boolean(facts.to) && Number.isFinite(Date.parse(facts.to)),
    "PATCH_E_EXPORT_UPPER_BOUND_MISSING",
    "Patch E summary must expose a valid export upper bound (`to`).",
  );
  return facts;
}

function validateApprovedActualPricing(policy = {}) {
  assertCondition(
    policy.version === "query_candidate_planner_cost_pricing_policy_v1",
    "PRICING_VERSION_INVALID",
    "Pricing policy version must match Patch 15.1 pricing contract.",
  );
  assertCondition(
    policy.mode === "APPROVED_ACTUAL",
    "ACTUAL_PRICING_NOT_APPROVED",
    "Pricing mode must be APPROVED_ACTUAL.",
  );
  assertCondition(
    policy.rateUnit === "MICROUSD_PER_MILLION_TOKENS",
    "PRICING_RATE_UNIT_INVALID",
    "Pricing rate unit must be MICROUSD_PER_MILLION_TOKENS.",
  );
  assertCondition(
    typeof policy.policyId === "string" &&
      policy.policyId.trim().length >= 8 &&
      !policy.policyId.includes("replace_with"),
    "PRICING_POLICY_ID_INVALID",
    "Pricing policyId must be a real change-history identifier.",
  );
  assertCondition(
    policy.guardrails?.approvedByOperator === true,
    "PRICING_OPERATOR_APPROVAL_REQUIRED",
    "Pricing policy must be approved by the operator.",
  );
  assertCondition(
    policy.guardrails?.productionBillingAuthority === false,
    "PRICING_BILLING_AUTHORITY_FORBIDDEN",
    "Evaluation pricing must not claim production billing authority.",
  );
  assertCondition(
    typeof policy.guardrails?.effectiveAt === "string" &&
      Number.isFinite(Date.parse(policy.guardrails.effectiveAt)),
    "PRICING_EFFECTIVE_AT_INVALID",
    "Pricing effectiveAt must be an ISO date-time.",
  );

  const models = policy.models;
  assertCondition(
    isObject(models) && Object.keys(models).length > 0,
    "PRICING_MODELS_REQUIRED",
    "At least one model price is required.",
  );
  for (const [modelId, rate] of Object.entries(models)) {
    assertCondition(
      modelId.trim().length > 0 && isObject(rate),
      "PRICING_MODEL_ENTRY_INVALID",
      "Every pricing model entry must be valid.",
    );
    const input = Number(rate.inputMicrousdPerMillionTokens);
    const output = Number(rate.outputMicrousdPerMillionTokens);
    assertCondition(
      Number.isFinite(input) &&
        input > 0 &&
        Number.isFinite(output) &&
        output > 0,
      "PRICING_RATE_MUST_BE_POSITIVE",
      `Model ${modelId} must have positive input/output rates.`,
    );
  }
  return Object.freeze({
    policyId: policy.policyId,
    effectiveAt: policy.guardrails.effectiveAt,
    modelCount: Object.keys(models).length,
  });
}

function operationalFacts(report = {}) {
  const decision = String(
    report.decision || findString(report, ["decision"]),
  ).trim();
  const sampleSize = findNumber(report, [
    "sampleSize",
    "executionCount",
    "totalExecutionCount",
  ]);
  const pricingSource = findString(report, [
    "pricingSource",
    "pricingMode",
    "mode",
  ]);
  const privacyViolationCount = findNumber(
    report,
    ["privacyViolationCount", "privacyViolations"],
    0,
  );
  const guardrailViolationCount = findNumber(
    report,
    ["guardrailViolationCount", "guardrailViolations"],
    0,
  );
  const promotionAuthorized = findBoolean(
    report,
    ["promotionAuthorized"],
    false,
  );
  const evaluationOnly = findBoolean(report, ["evaluationOnly"], true);
  const productionCandidateMergeApplied = findBoolean(
    report,
    ["productionCandidateMergeApplied"],
    false,
  );
  const productionReadyAssignment = findBoolean(
    report,
    ["productionReadyAssignment"],
    false,
  );

  return Object.freeze({
    decision,
    sampleSize,
    pricingSource,
    privacyViolationCount,
    guardrailViolationCount,
    promotionAuthorized,
    evaluationOnly,
    productionCandidateMergeApplied,
    productionReadyAssignment,
  });
}

function validOperationalReportVersion(version) {
  const normalized = String(version || "")
    .trim()
    .toLowerCase();
  if (!normalized) return false;
  if (
    normalized ===
    "query_candidate_planner_cost_cache_latency_evaluation_report_v1"
  ) {
    return true;
  }
  return (
    normalized.startsWith("query_candidate_planner_") &&
    (normalized.includes("cost_cache_latency") ||
      normalized.includes("operational")) &&
    normalized.includes("evaluation") &&
    normalized.includes("report") &&
    normalized.includes("v1")
  );
}

function validateOperationalReport(report = {}, policy = DEFAULT_POLICY) {
  assertCondition(
    validOperationalReportVersion(report.version),
    "OPERATIONAL_REPORT_VERSION_INVALID",
    "Operational report must use a compatible Patch 15.1 operational evaluation report v1 contract.",
  );
  const facts = operationalFacts(report);
  assertCondition(
    facts.decision === "EVALUATION_PASS",
    "OPERATIONAL_EVALUATION_BLOCKED",
    "Operational evaluation must be EVALUATION_PASS.",
  );
  if (Number.isFinite(facts.sampleSize)) {
    assertCondition(
      facts.sampleSize >= policy.minimumExecutionCount,
      "OPERATIONAL_SAMPLE_TOO_SMALL",
      `Operational report sample must be >= ${policy.minimumExecutionCount}.`,
    );
  }
  assertCondition(
    facts.evaluationOnly === true,
    "OPERATIONAL_EVALUATION_ONLY_REQUIRED",
    "Operational report must remain evaluationOnly.",
  );
  assertCondition(
    facts.promotionAuthorized === false,
    "OPERATIONAL_PROMOTION_AUTHORIZATION_FORBIDDEN",
    "Operational report must not authorize promotion.",
  );
  assertCondition(
    facts.productionCandidateMergeApplied === false,
    "OPERATIONAL_PRODUCTION_MERGE_FORBIDDEN",
    "Operational report must not apply production merge.",
  );
  assertCondition(
    facts.productionReadyAssignment === false,
    "OPERATIONAL_READY_ASSIGNMENT_FORBIDDEN",
    "Operational report must not assign production readiness.",
  );
  assertCondition(
    facts.privacyViolationCount <= policy.maximumPrivacyViolationCount,
    "OPERATIONAL_PRIVACY_VIOLATION",
    "Operational privacy violations must be zero.",
  );
  assertCondition(
    facts.guardrailViolationCount <= policy.maximumGuardrailViolationCount,
    "OPERATIONAL_GUARDRAIL_VIOLATION",
    "Operational guardrail violations must be zero.",
  );

  const normalizedPricing = facts.pricingSource.toUpperCase();
  if (normalizedPricing) {
    assertCondition(
      normalizedPricing.includes("APPROVED_ACTUAL"),
      "OPERATIONAL_PRICING_NOT_ACTUAL",
      "Operational report must be produced with approved actual pricing.",
    );
  }
  return facts;
}

const METRIC_KEYS = Object.freeze([
  "cacheHitRate",
  "warmCacheHitRate",
  "downloadReuseCacheHitRate",
  "providerCallRate",
  "warmProviderCallRate",
  "reuploadProviderCallRate",
  "overallLatencyP50Ms",
  "overallLatencyP95Ms",
  "overallLatencyP99Ms",
  "warmLatencyP95Ms",
  "cacheHitLatencyP95Ms",
  "timeoutRate",
  "errorRate",
  "totalCostMicrousd",
  "averageCostMicrousd",
  "averageProviderCallCostMicrousd",
  "warmAverageCostMicrousd",
  "cacheAvoidedCostMicrousd",
  "cacheCostAvoidanceRate",
  "monthlyProjectedCostMicrousd",
  "downloadRetentionAccuracy",
  "deleteInvalidationCoverage",
  "reuploadIdentitySeparationAccuracy",
  "staleCacheReuseViolationCount",
]);

function extractOperationalMetrics(report = {}) {
  const metrics = {};
  for (const key of METRIC_KEYS) {
    const value = findNumber(report, [key], null);
    if (value !== null) metrics[key] = value;
  }
  return Object.freeze(metrics);
}

function countExportedRecords(records) {
  if (Array.isArray(records)) return records.length;
  if (Array.isArray(records?.records)) return records.records.length;
  if (Array.isArray(records?.entries)) return records.entries.length;
  return null;
}

function buildRealShadowEvaluationBaseline({
  collectionSummary,
  pricingPolicy,
  operationalReport,
  exportedRecords = null,
  policy = DEFAULT_POLICY,
} = {}) {
  const collection = validateCollectionSummary(collectionSummary, policy);
  const pricing = validateApprovedActualPricing(pricingPolicy);
  const operational = validateOperationalReport(operationalReport, policy);

  const exportedRecordCount = countExportedRecords(exportedRecords);
  if (exportedRecordCount !== null) {
    assertCondition(
      exportedRecordCount >=
        policy.minimumExecutionCount + policy.minimumLifecycleCount,
      "EXPORTED_RECORD_SET_TOO_SMALL",
      "Frozen export must contain at least the Patch E execution+lifecycle minimum.",
    );
  }

  const baseline = {
    version: BASELINE_VERSION,
    decision: BASELINE_DECISION,
    failClosed: true,
    evaluationOnly: true,
    promotionAuthorized: false,
    source: {
      collectionSummarySha256: sha256Json(collectionSummary),
      pricingPolicySha256: sha256Json(pricingPolicy),
      operationalReportSha256: sha256Json(operationalReport),
      exportedRecordsSha256:
        exportedRecords == null ? "" : sha256Json(exportedRecords),
      evaluationWindow: {
        from: collection.from,
        to: collection.to,
      },
    },
    coverage: {
      executionCount: collection.executionCount,
      lifecycleCount: collection.lifecycleCount,
      caseCount: collection.caseCount,
      exportedRecordCount,
      collectionProtocolComplete: collection.protocolComplete,
    },
    pricing: {
      mode: "APPROVED_ACTUAL",
      policyId: pricing.policyId,
      effectiveAt: pricing.effectiveAt,
      modelCount: pricing.modelCount,
    },
    operational: {
      reportVersion: operationalReport.version,
      decision: operational.decision,
      sampleSize: operational.sampleSize,
      pricingSource: operational.pricingSource || "APPROVED_ACTUAL",
      metrics: extractOperationalMetrics(operationalReport),
    },
    guardrails: {
      privacyViolationCount: Math.max(
        collection.privacyViolationCount,
        operational.privacyViolationCount,
      ),
      guardrailViolationCount: Math.max(
        collection.guardrailViolationCount,
        operational.guardrailViolationCount,
      ),
      collectorEnabledByThisOperation: false,
      internalCanaryEnabledByThisOperation: false,
      productionPromotionAuthorized: false,
      productionCandidateMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      providerCallsExecutedByBaselineBuilder: 0,
      privateOutputDoNotCommit: true,
    },
  };

  const baselineSha256 = sha256Json(baseline);
  return Object.freeze({ ...baseline, baselineSha256 });
}

module.exports = Object.freeze({
  BASELINE_VERSION,
  BASELINE_DECISION,
  DEFAULT_POLICY,
  stableStringify,
  sha256Json,
  collectionFacts,
  validateCollectionSummary,
  validateApprovedActualPricing,
  operationalFacts,
  validOperationalReportVersion,
  validateOperationalReport,
  extractOperationalMetrics,
  buildRealShadowEvaluationBaseline,
});
