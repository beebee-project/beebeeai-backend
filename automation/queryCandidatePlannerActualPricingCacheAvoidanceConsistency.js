const crypto = require("crypto");

const CONSISTENCY_VERSION =
  "query_candidate_planner_actual_pricing_cache_avoidance_consistency_v1";

const EXPECTED_INPUT_VERSION =
  "query_candidate_planner_e_x_compatibility_canonical_evaluation_input_v1";

const EXPECTED_BENCHMARK_MODE =
  "CANONICAL_BENCHMARK_WITH_APPROVED_ACTUAL_PRICING";

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function stableValue(value) {
  if (Array.isArray(value)) return value.map(stableValue);
  if (!isObject(value)) return value;
  const out = {};
  for (const key of Object.keys(value).sort()) {
    out[key] = stableValue(value[key]);
  }
  return out;
}

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(JSON.stringify(stableValue(value)))
    .digest("hex");
}

function assertCondition(condition, code, message) {
  if (!condition) {
    const error = new Error(message || code);
    error.code = code;
    throw error;
  }
}

function validatePricingPolicy(policy = {}) {
  assertCondition(
    policy.version === "query_candidate_planner_cost_pricing_policy_v1",
    "PRICING_VERSION_INVALID",
    "Patch F-1 approved pricing policy is required.",
  );
  assertCondition(
    policy.mode === "APPROVED_ACTUAL",
    "PRICING_NOT_APPROVED_ACTUAL",
    "Pricing mode must be APPROVED_ACTUAL.",
  );
  assertCondition(
    policy.rateUnit === "MICROUSD_PER_MILLION_TOKENS",
    "PRICING_RATE_UNIT_INVALID",
    "Pricing rate unit must be MICROUSD_PER_MILLION_TOKENS.",
  );
  assertCondition(
    policy.guardrails?.approvedByOperator === true,
    "PRICING_OPERATOR_APPROVAL_REQUIRED",
    "Operator-approved pricing is required.",
  );
  assertCondition(
    policy.guardrails?.productionBillingAuthority === false,
    "PRICING_BILLING_AUTHORITY_FORBIDDEN",
    "Evaluation pricing must not claim production billing authority.",
  );
  return true;
}

function calculateProviderCostMicrousd(provider = {}, pricingPolicy = {}) {
  validatePricingPolicy(pricingPolicy);

  assertCondition(
    provider.called === true,
    "PROVIDER_CALL_REQUIRED",
    "Provider-called execution is required for token cost calculation.",
  );

  const modelId = String(provider.modelId || "").trim();
  assertCondition(
    modelId.length > 0,
    "MODEL_ID_REQUIRED",
    "Provider modelId is required.",
  );

  const rates = pricingPolicy.models?.[modelId];
  assertCondition(
    rates && typeof rates === "object",
    "MODEL_PRICING_MISSING",
    `Approved pricing is missing for model ${modelId}.`,
  );

  const inputTokens = Number(provider.inputTokens);
  const outputTokens = Number(provider.outputTokens);
  const inputRate = Number(rates.inputMicrousdPerMillionTokens);
  const outputRate = Number(rates.outputMicrousdPerMillionTokens);

  assertCondition(
    Number.isFinite(inputTokens) &&
      inputTokens >= 0 &&
      Number.isFinite(outputTokens) &&
      outputTokens >= 0,
    "TOKEN_COUNT_INVALID",
    "Provider token counts must be non-negative finite numbers.",
  );
  assertCondition(
    Number.isFinite(inputRate) &&
      inputRate > 0 &&
      Number.isFinite(outputRate) &&
      outputRate > 0,
    "MODEL_PRICING_INVALID",
    "Approved model pricing must be positive.",
  );

  return Math.round(
    (inputTokens * inputRate + outputTokens * outputRate) / 1_000_000,
  );
}

function validateCanonicalInput(canonicalInput = {}) {
  assertCondition(
    canonicalInput.version === EXPECTED_INPUT_VERSION,
    "CANONICAL_INPUT_VERSION_INVALID",
    "Patch 15.3.2-F.1 canonical input is required.",
  );
  assertCondition(
    canonicalInput.benchmarkMode === EXPECTED_BENCHMARK_MODE,
    "CANONICAL_INPUT_MODE_INVALID",
    "Canonical input benchmark mode is invalid.",
  );
  assertCondition(
    Array.isArray(canonicalInput.dataset?.executions) &&
      canonicalInput.dataset.executions.length >= 25,
    "CANONICAL_EXECUTIONS_INSUFFICIENT",
    "Canonical evaluation input requires at least 25 executions.",
  );
  assertCondition(
    canonicalInput.guardrails?.actualOperationalTelemetry === false,
    "ACTUAL_TELEMETRY_FORBIDDEN",
    "This compatibility patch is not actual operational telemetry.",
  );
  assertCondition(
    canonicalInput.guardrails?.productionPromotionAuthorized === false,
    "PROMOTION_AUTHORIZATION_FORBIDDEN",
    "Canonical input must not authorize production promotion.",
  );
  return true;
}

function buildColdCostMap(executions, pricingPolicy) {
  const coldByScenario = new Map();

  for (const execution of executions) {
    if (String(execution.phase || "").toUpperCase() !== "COLD") continue;

    const scenarioId = String(execution.scenarioId || "").trim();
    assertCondition(
      scenarioId.length > 0,
      "COLD_SCENARIO_ID_REQUIRED",
      "COLD execution scenarioId is required.",
    );
    assertCondition(
      !coldByScenario.has(scenarioId),
      "DUPLICATE_COLD_EXECUTION",
      `Exactly one COLD execution is required per scenario: ${scenarioId}.`,
    );

    const costMicrousd = calculateProviderCostMicrousd(
      execution.provider,
      pricingPolicy,
    );

    coldByScenario.set(
      scenarioId,
      Object.freeze({
        scenarioId,
        executionId: String(execution.executionId || ""),
        modelId: String(execution.provider.modelId || ""),
        inputTokens: Number(execution.provider.inputTokens),
        outputTokens: Number(execution.provider.outputTokens),
        costMicrousd,
      }),
    );
  }

  assertCondition(
    coldByScenario.size > 0,
    "COLD_EXECUTIONS_REQUIRED",
    "At least one COLD execution is required.",
  );

  return coldByScenario;
}

function deriveActualPricingConsistentInput({
  canonicalInput,
  pricingPolicy,
} = {}) {
  validateCanonicalInput(canonicalInput);
  validatePricingPolicy(pricingPolicy);

  const sourceSnapshot = JSON.stringify(canonicalInput);
  const derived = clone(canonicalInput);
  const executions = derived.dataset.executions;
  const coldCostMap = buildColdCostMap(executions, pricingPolicy);

  const sourceExecutions = canonicalInput.dataset.executions;
  const sourceScenarioIds = new Set(
    sourceExecutions.map((execution) => String(execution.scenarioId || "")),
  );

  assertCondition(
    coldCostMap.size === sourceScenarioIds.size,
    "COLD_COVERAGE_INCOMPLETE",
    "Every canonical scenario must contain exactly one priced COLD execution.",
  );

  let repricedExecutionCount = 0;
  let repricedCacheHitCount = 0;
  let avoidedByCacheMicrousd = 0;

  for (const execution of executions) {
    const scenarioId = String(execution.scenarioId || "").trim();
    const cold = coldCostMap.get(scenarioId);
    assertCondition(
      cold,
      "SCENARIO_COLD_COST_MISSING",
      `Missing actual-pricing COLD cost for scenario ${scenarioId}.`,
    );

    execution.expectedColdCostMicrousd = cold.costMicrousd;
    repricedExecutionCount += 1;

    if (execution.cache?.hit === true) {
      avoidedByCacheMicrousd += cold.costMicrousd;
      repricedCacheHitCount += 1;
    }
  }

  let providerCostMicrousd = 0;
  let providerCallCount = 0;
  for (const execution of executions) {
    if (execution.provider?.called !== true) continue;
    providerCostMicrousd += calculateProviderCostMicrousd(
      execution.provider,
      pricingPolicy,
    );
    providerCallCount += 1;
  }

  const denominator = providerCostMicrousd + avoidedByCacheMicrousd;
  const cacheCostAvoidanceRate =
    denominator > 0
      ? Number((avoidedByCacheMicrousd / denominator).toFixed(6))
      : 0;

  derived.dataset.compatibility = derived.dataset.compatibility || {};

  derived.dataset.compatibility.actualPricingCacheAvoidanceConsistency = {
    version: CONSISTENCY_VERSION,
    method: "SCENARIO_COLD_TOKEN_COST_REPRICED_WITH_APPROVED_ACTUAL_POLICY",
    sourceCanonicalInputSha256: sha256Json(canonicalInput),
    pricingPolicySha256: sha256Json(pricingPolicy),
    scenarioCount: coldCostMap.size,
    repricedExecutionCount,
    repricedCacheHitCount,
    providerCallCount,
    providerCostMicrousd,
    avoidedByCacheMicrousd,
    cacheCostAvoidanceRate,
    originalThresholdPolicyModified: false,
    evaluatorModified: false,
    providerCallsExecutedByConsistencyPreparation: 0,
    actualOperationalTelemetry: false,
    productionPromotionAuthorized: false,
  };

  derived.dataset.benchmarkMode =
    "CANONICAL_BENCHMARK_WITH_APPROVED_ACTUAL_PRICING_AND_COST_CONSISTENCY";

  assertCondition(
    JSON.stringify(canonicalInput) === sourceSnapshot,
    "SOURCE_CANONICAL_INPUT_MUTATED",
    "Source canonical input was mutated.",
  );

  return Object.freeze({
    version: EXPECTED_INPUT_VERSION,
    benchmarkMode: EXPECTED_BENCHMARK_MODE,
    dataset: derived.dataset,
    source: {
      ...(derived.source || {}),
      preConsistencyCanonicalInputSha256: sha256Json(canonicalInput),
      pricingPolicySha256: sha256Json(pricingPolicy),
    },
    guardrails: {
      ...(derived.guardrails || {}),
      actualOperationalTelemetry: false,
      productionPromotionAuthorized: false,
      productionCandidateMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      providerCallsExecutedByThisOperation: 0,
    },
    costConsistency: {
      version: CONSISTENCY_VERSION,
      scenarioCount: coldCostMap.size,
      repricedExecutionCount,
      repricedCacheHitCount,
      providerCallCount,
      providerCostMicrousd,
      avoidedByCacheMicrousd,
      cacheCostAvoidanceRate,
      thresholdPolicyModified: false,
      evaluatorModified: false,
    },
  });
}

function findThresholdResult(report, metric) {
  const results = Array.isArray(report?.thresholdResults)
    ? report.thresholdResults
    : [];
  return results.find((item) => item?.metric === metric) || null;
}

function assessCostConsistencyReevaluation(report = {}) {
  const targetMetrics = [
    "cost.averagePerExecutionMicrousd",
    "cost.averagePerProviderCallMicrousd",
    "cost.monthlyProjectedCostMicrousd",
    "cost.cacheCostAvoidanceRate",
  ];

  const byMetric = {};
  for (const metric of targetMetrics) {
    const result = findThresholdResult(report, metric);
    assertCondition(
      result,
      "COST_THRESHOLD_RESULT_MISSING",
      `Missing threshold result: ${metric}.`,
    );
    byMetric[metric] = result;
  }

  const absoluteCostMetrics = targetMetrics.slice(0, 3);
  const absoluteFailures = absoluteCostMetrics.filter(
    (metric) => byMetric[metric].passed !== true,
  );

  const cacheAvoidance = byMetric["cost.cacheCostAvoidanceRate"];

  const nonCostFailures = (report.thresholdResults || []).filter(
    (item) =>
      item?.passed === false && !String(item.metric || "").startsWith("cost."),
  );

  const decision =
    cacheAvoidance.passed === true &&
    absoluteFailures.length === 3 &&
    nonCostFailures.length === 0
      ? "CACHE_AVOIDANCE_PRICING_CONSISTENCY_PASS_ABSOLUTE_COST_RECALIBRATION_REQUIRED"
      : "COST_CONSISTENCY_REEVALUATION_REVIEW_REQUIRED";

  return Object.freeze({
    version:
      "query_candidate_planner_actual_pricing_cost_consistency_reevaluation_assessment_v1",
    decision,
    operationalDecision: String(report.decision || ""),
    targetCostResults: byMetric,
    absoluteCostFailureCount: absoluteFailures.length,
    absoluteCostFailures: absoluteFailures,
    cacheCostAvoidancePassed: cacheAvoidance.passed === true,
    cacheCostAvoidanceActual: Number(cacheAvoidance.actual),
    cacheCostAvoidanceThreshold: Number(cacheAvoidance.threshold),
    nonCostFailureCount: nonCostFailures.length,
    thresholdPolicyModified: false,
    evaluatorModified: false,
    providerCallsExecutedByAssessment: 0,
    productionPromotionAuthorized: false,
  });
}

module.exports = Object.freeze({
  CONSISTENCY_VERSION,
  calculateProviderCostMicrousd,
  validateCanonicalInput,
  buildColdCostMap,
  deriveActualPricingConsistentInput,
  findThresholdResult,
  assessCostConsistencyReevaluation,
});
