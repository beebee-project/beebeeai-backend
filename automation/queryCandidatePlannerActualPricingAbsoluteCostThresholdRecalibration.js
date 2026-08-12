const crypto = require("crypto");

const VERSION =
  "query_candidate_planner_actual_pricing_absolute_cost_threshold_recalibration_v1";

const HEADROOM_RATE = 0.2;
const ROUNDING_INCREMENT_MICROUSD = 100;

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (!value || typeof value !== "object") return value;
  const out = {};
  for (const key of Object.keys(value).sort()) out[key] = stable(value[key]);
  return out;
}

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(JSON.stringify(stable(value)))
    .digest("hex");
}

function fail(code, message = code) {
  const error = new Error(message);
  error.code = code;
  throw error;
}

function assert(condition, code, message) {
  if (!condition) fail(code, message);
}

function validatePricingPolicy(policy = {}) {
  assert(
    policy.version === "query_candidate_planner_cost_pricing_policy_v1",
    "PRICING_VERSION_INVALID",
  );
  assert(policy.mode === "APPROVED_ACTUAL", "PRICING_NOT_APPROVED_ACTUAL");
  assert(
    policy.rateUnit === "MICROUSD_PER_MILLION_TOKENS",
    "PRICING_RATE_UNIT_INVALID",
  );
  assert(
    policy.guardrails?.approvedByOperator === true,
    "PRICING_OPERATOR_APPROVAL_REQUIRED",
  );
  assert(
    policy.guardrails?.productionBillingAuthority === false,
    "PRODUCTION_BILLING_AUTHORITY_FORBIDDEN",
  );
}

function validateSourceThresholdPolicy(policy = {}) {
  assert(
    policy.version ===
      "query_candidate_planner_operational_threshold_policy_v1",
    "THRESHOLD_VERSION_INVALID",
  );
  assert(
    policy.thresholds && typeof policy.thresholds === "object",
    "THRESHOLDS_REQUIRED",
  );
  assert(
    Number(policy.thresholds.providerCallRateMax) === 0.4,
    "PROVIDER_CALL_RATE_CONTRACT_DRIFT",
  );
  assert(
    Number(policy.monthlyProjectionExecutions) === 10000,
    "MONTHLY_PROJECTION_CONTRACT_DRIFT",
  );
  assert(
    Number(policy.thresholds.cacheCostAvoidanceRateMin) === 0.59,
    "CACHE_AVOIDANCE_THRESHOLD_DRIFT",
  );
  assert(
    Number(policy.thresholds.warmAverageCostMicrousdMax) === 0,
    "WARM_COST_THRESHOLD_DRIFT",
  );
  assert(
    policy.guardrails?.evaluationOnly === true &&
      policy.guardrails?.promotionAuthorized === false &&
      policy.guardrails?.productionMergeAuthorized === false,
    "THRESHOLD_GUARDRAIL_DRIFT",
  );
}

function calculateProviderCostMicrousd(provider = {}, pricingPolicy = {}) {
  validatePricingPolicy(pricingPolicy);
  assert(provider.called === true, "PROVIDER_CALL_REQUIRED");

  const modelId = String(provider.modelId || "").trim();
  const rates = pricingPolicy.models?.[modelId];
  assert(rates, "MODEL_PRICING_MISSING");

  const inputTokens = Number(provider.inputTokens);
  const outputTokens = Number(provider.outputTokens);
  const inputRate = Number(rates.inputMicrousdPerMillionTokens);
  const outputRate = Number(rates.outputMicrousdPerMillionTokens);

  assert(
    Number.isFinite(inputTokens) &&
      inputTokens >= 0 &&
      Number.isFinite(outputTokens) &&
      outputTokens >= 0,
    "TOKEN_COUNT_INVALID",
  );
  assert(
    Number.isFinite(inputRate) &&
      inputRate > 0 &&
      Number.isFinite(outputRate) &&
      outputRate > 0,
    "PRICING_RATE_INVALID",
  );

  return Math.round(
    (inputTokens * inputRate + outputTokens * outputRate) / 1_000_000,
  );
}

function nearestRank(values, rate) {
  const sorted = [...values].sort((a, b) => a - b);
  const rank = Math.max(1, Math.ceil(sorted.length * rate));
  return sorted[Math.min(sorted.length - 1, rank - 1)];
}

function roundUp(value, increment = ROUNDING_INCREMENT_MICROUSD) {
  return Math.ceil(value / increment) * increment;
}

function providerCostDistribution(dataset = {}, pricingPolicy = {}) {
  const executions = Array.isArray(dataset.executions)
    ? dataset.executions
    : [];
  assert(executions.length >= 25, "EXECUTION_SAMPLE_INSUFFICIENT");

  const providerRows = executions.filter(
    (row) => row.provider?.called === true,
  );
  assert(
    providerRows.length >= 10,
    "PROVIDER_CALL_SAMPLE_INSUFFICIENT",
    "At least 10 provider-call observations are required.",
  );

  const costs = providerRows.map((row) =>
    calculateProviderCostMicrousd(row.provider, pricingPolicy),
  );
  const sorted = [...costs].sort((a, b) => a - b);
  const total = sorted.reduce((sum, value) => sum + value, 0);

  return Object.freeze({
    sampleCount: sorted.length,
    costsMicrousd: Object.freeze(sorted),
    minMicrousd: sorted[0],
    averageMicrousd: Math.round(total / sorted.length),
    p50Microusd: nearestRank(sorted, 0.5),
    p95Microusd: nearestRank(sorted, 0.95),
    maxMicrousd: sorted[sorted.length - 1],
    totalMicrousd: total,
  });
}

function deriveThresholds({
  dataset,
  pricingPolicy,
  sourceThresholdPolicy,
  headroomRate = HEADROOM_RATE,
  roundingIncrementMicrousd = ROUNDING_INCREMENT_MICROUSD,
} = {}) {
  validatePricingPolicy(pricingPolicy);
  validateSourceThresholdPolicy(sourceThresholdPolicy);

  assert(
    Number.isFinite(headroomRate) && headroomRate >= 0.1 && headroomRate <= 0.5,
    "HEADROOM_RATE_INVALID",
  );
  assert(
    Number.isInteger(roundingIncrementMicrousd) &&
      roundingIncrementMicrousd > 0,
    "ROUNDING_INCREMENT_INVALID",
  );

  const distribution = providerCostDistribution(dataset, pricingPolicy);

  const rawProviderCeilingMicrousd =
    distribution.maxMicrousd * (1 + headroomRate);

  const providerCallAverageCostMicrousdMax = roundUp(
    rawProviderCeilingMicrousd,
    roundingIncrementMicrousd,
  );

  const providerCallRateMax = Number(
    sourceThresholdPolicy.thresholds.providerCallRateMax,
  );

  const averageCostMicrousdMax = roundUp(
    providerCallAverageCostMicrousdMax * providerCallRateMax,
    roundingIncrementMicrousd,
  );

  const monthlyProjectionExecutions = Number(
    sourceThresholdPolicy.monthlyProjectionExecutions,
  );

  const monthlyProjectedCostMicrousdMax =
    averageCostMicrousdMax * monthlyProjectionExecutions;

  return Object.freeze({
    version: VERSION,
    methodology:
      "MAX_OBSERVED_PROVIDER_COST_PLUS_20_PERCENT_HEADROOM_WITH_PRESERVED_CALL_RATE",
    headroomRate,
    roundingIncrementMicrousd,
    distribution,
    rawProviderCeilingMicrousd,
    providerCallAverageCostMicrousdMax,
    providerCallRateMax,
    averageCostMicrousdMax,
    monthlyProjectionExecutions,
    monthlyProjectedCostMicrousdMax,
  });
}

function buildPrivateThresholdPolicy({
  dataset,
  pricingPolicy,
  sourceThresholdPolicy,
} = {}) {
  const sourceSnapshot = JSON.stringify(sourceThresholdPolicy);
  const derived = deriveThresholds({
    dataset,
    pricingPolicy,
    sourceThresholdPolicy,
  });

  const policy = clone(sourceThresholdPolicy);

  // Keep version/policyId/schema compatible with the existing evaluator.
  policy.thresholds.averageCostMicrousdMax = derived.averageCostMicrousdMax;
  policy.thresholds.providerCallAverageCostMicrousdMax =
    derived.providerCallAverageCostMicrousdMax;
  policy.thresholds.monthlyProjectedCostMicrousdMax =
    derived.monthlyProjectedCostMicrousdMax;

  assert(
    JSON.stringify(sourceThresholdPolicy) === sourceSnapshot,
    "SOURCE_THRESHOLD_POLICY_MUTATED",
  );

  const changed = Object.keys(sourceThresholdPolicy.thresholds).filter(
    (key) => sourceThresholdPolicy.thresholds[key] !== policy.thresholds[key],
  );

  assert(
    changed.length === 3 &&
      changed.includes("averageCostMicrousdMax") &&
      changed.includes("providerCallAverageCostMicrousdMax") &&
      changed.includes("monthlyProjectedCostMicrousdMax"),
    "THRESHOLD_CHANGE_SCOPE_INVALID",
  );

  assert(
    policy.thresholds.cacheCostAvoidanceRateMin === 0.59,
    "CACHE_AVOIDANCE_THRESHOLD_CHANGED",
  );

  return Object.freeze({
    policy: Object.freeze(policy),
    evidence: Object.freeze({
      version: VERSION,
      sourceThresholdPolicySha256: sha256Json(sourceThresholdPolicy),
      pricingPolicySha256: sha256Json(pricingPolicy),
      datasetSha256: sha256Json(dataset),
      methodology: derived.methodology,
      headroomRate: derived.headroomRate,
      roundingIncrementMicrousd: derived.roundingIncrementMicrousd,
      providerCostDistribution: derived.distribution,
      rawProviderCeilingMicrousd: derived.rawProviderCeilingMicrousd,
      derivedThresholds: {
        providerCallAverageCostMicrousdMax:
          derived.providerCallAverageCostMicrousdMax,
        averageCostMicrousdMax: derived.averageCostMicrousdMax,
        monthlyProjectedCostMicrousdMax:
          derived.monthlyProjectedCostMicrousdMax,
      },
      preservedContracts: {
        providerCallRateMax: derived.providerCallRateMax,
        monthlyProjectionExecutions: derived.monthlyProjectionExecutions,
        cacheCostAvoidanceRateMin: 0.59,
        warmAverageCostMicrousdMax: 0,
      },
      changedThresholds: Object.freeze([...changed]),
      sourceThresholdPolicyModified: false,
      evaluatorModified: false,
      providerCallsExecutedByRecalibration: 0,
      actualOperationalTelemetry: false,
      productionPromotionAuthorized: false,
      privateOutputDoNotCommit: true,
    }),
  });
}

function assessEvaluation(report = {}, thresholdPolicy = {}) {
  const results = Array.isArray(report.thresholdResults)
    ? report.thresholdResults
    : [];

  const failed = results.filter((item) => item?.passed === false);

  const absoluteMetrics = [
    "cost.averagePerExecutionMicrousd",
    "cost.averagePerProviderCallMicrousd",
    "cost.monthlyProjectedCostMicrousd",
  ];

  const absolute = absoluteMetrics.map((metric) => {
    const result = results.find((item) => item?.metric === metric);
    assert(result, "ABSOLUTE_COST_RESULT_MISSING", metric);
    return result;
  });

  const cacheAvoidance = results.find(
    (item) => item?.metric === "cost.cacheCostAvoidanceRate",
  );
  assert(cacheAvoidance, "CACHE_AVOIDANCE_RESULT_MISSING");

  const decision =
    report.decision === "EVALUATION_PASS" &&
    failed.length === 0 &&
    absolute.every((item) => item.passed === true) &&
    cacheAvoidance.passed === true
      ? "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS"
      : "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_REVIEW_REQUIRED";

  return Object.freeze({
    version:
      "query_candidate_planner_actual_pricing_absolute_cost_threshold_recalibration_assessment_v1",
    decision,
    operationalDecision: String(report.decision || ""),
    failedCheckCount: failed.length,
    absoluteCostPassCount: absolute.filter((item) => item.passed === true)
      .length,
    absoluteCostFailureCount: absolute.filter((item) => item.passed !== true)
      .length,
    cacheCostAvoidancePassed: cacheAvoidance.passed === true,
    cacheCostAvoidanceActual: Number(cacheAvoidance.actual),
    cacheCostAvoidanceThreshold: Number(cacheAvoidance.threshold),
    thresholds: {
      averageCostMicrousdMax: Number(
        thresholdPolicy.thresholds?.averageCostMicrousdMax,
      ),
      providerCallAverageCostMicrousdMax: Number(
        thresholdPolicy.thresholds?.providerCallAverageCostMicrousdMax,
      ),
      monthlyProjectedCostMicrousdMax: Number(
        thresholdPolicy.thresholds?.monthlyProjectedCostMicrousdMax,
      ),
      cacheCostAvoidanceRateMin: Number(
        thresholdPolicy.thresholds?.cacheCostAvoidanceRateMin,
      ),
    },
    sourceThresholdPolicyModified: false,
    evaluatorModified: false,
    providerCallsExecutedByAssessment: 0,
    productionPromotionAuthorized: false,
  });
}

module.exports = Object.freeze({
  VERSION,
  HEADROOM_RATE,
  ROUNDING_INCREMENT_MICROUSD,
  calculateProviderCostMicrousd,
  providerCostDistribution,
  deriveThresholds,
  buildPrivateThresholdPolicy,
  assessEvaluation,
});
