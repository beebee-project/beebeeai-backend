const crypto = require("crypto");

const INPUT_VERSION =
  "query_candidate_planner_e_x_compatibility_canonical_evaluation_input_v1";
const BENCHMARK_MODE = "CANONICAL_BENCHMARK_WITH_APPROVED_ACTUAL_PRICING";

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
    "Pricing version must match Patch 15.1.",
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
    "Operator approval is required.",
  );
  assertCondition(
    policy.guardrails?.productionBillingAuthority === false,
    "PRICING_BILLING_AUTHORITY_FORBIDDEN",
    "Evaluation pricing cannot claim production billing authority.",
  );
  assertCondition(
    isObject(policy.models) && Object.keys(policy.models).length > 0,
    "PRICING_MODELS_REQUIRED",
    "At least one model pricing entry is required.",
  );
  for (const [modelId, rates] of Object.entries(policy.models)) {
    const input = Number(rates?.inputMicrousdPerMillionTokens);
    const output = Number(rates?.outputMicrousdPerMillionTokens);
    assertCondition(
      modelId &&
        Number.isFinite(input) &&
        input > 0 &&
        Number.isFinite(output) &&
        output > 0,
      "PRICING_RATE_INVALID",
      `Positive input/output pricing is required for ${modelId}.`,
    );
  }
  return true;
}

function validateLiveParityReadiness(readiness = {}) {
  const origin = readiness.origin || {};
  const replay = readiness.replay || {};
  const parity = readiness.parityAudit || {};
  const gate = readiness.readinessGate || {};
  const gateGuardrails = gate.guardrails || {};

  assertCondition(
    origin.status === "SHADOW_COMPLETED",
    "LIVE_PARITY_ORIGIN_STATUS_INVALID",
    "Live parity origin must be SHADOW_COMPLETED.",
  );
  assertCondition(
    origin.invocationStatus === "CALLED",
    "LIVE_PARITY_ORIGIN_NOT_CALLED",
    "Live parity origin must be CALLED.",
  );
  assertCondition(
    Number(origin.providerCallCount) === 1,
    "LIVE_PARITY_ORIGIN_PROVIDER_CALL_COUNT_INVALID",
    "Live parity origin must contain exactly one provider call.",
  );
  assertCondition(
    replay.status === "SHADOW_COMPLETED",
    "LIVE_PARITY_REPLAY_STATUS_INVALID",
    "Live parity replay must be SHADOW_COMPLETED.",
  );
  assertCondition(
    replay.invocationStatus === "CACHE_HIT",
    "LIVE_PARITY_REPLAY_NOT_CACHE_HIT",
    "Live parity replay must be CACHE_HIT.",
  );
  assertCondition(
    Number(replay.providerCallCount) === 0,
    "LIVE_PARITY_REPLAY_PROVIDER_CALL_FORBIDDEN",
    "Live parity replay must execute zero provider calls.",
  );
  assertCondition(
    replay.cache?.plannerResolution?.source === "L3_SEMANTIC",
    "LIVE_PARITY_L3_SOURCE_INVALID",
    "Planner replay source must be L3_SEMANTIC.",
  );
  assertCondition(
    replay.cache?.reentry?.source === "L4_REENTRY",
    "LIVE_PARITY_L4_SOURCE_INVALID",
    "Re-entry replay source must be L4_REENTRY.",
  );
  assertCondition(
    parity.valid === true,
    "LIVE_PARITY_AUDIT_INVALID",
    "Live parity audit must be valid.",
  );
  assertCondition(
    Number(parity.observedProviderCallCount) === 1,
    "LIVE_PARITY_OBSERVED_PROVIDER_CALL_COUNT_INVALID",
    "Parity audit must observe exactly one provider call.",
  );
  assertCondition(
    Number(parity.persistentFiles?.encryptedFileCount) >= 3,
    "LIVE_PARITY_ENCRYPTED_CACHE_EVIDENCE_INSUFFICIENT",
    "At least three encrypted persistent cache files are required.",
  );
  assertCondition(
    Number(parity.persistentFiles?.plaintextFileCount) === 0,
    "LIVE_PARITY_PLAINTEXT_CACHE_FORBIDDEN",
    "Plaintext persistent cache files are forbidden.",
  );
  assertCondition(
    gate.eligible === true,
    "LIVE_PARITY_READINESS_NOT_ELIGIBLE",
    "Patch 13.3 readiness evidence must be eligible.",
  );
  assertCondition(
    gateGuardrails.productionPromotionAllowed === false,
    "LIVE_PARITY_PROMOTION_AUTHORIZATION_FORBIDDEN",
    "Patch 13.3 evidence must not authorize production promotion.",
  );
  assertCondition(
    gateGuardrails.productionRouteAutoWired === false,
    "LIVE_PARITY_ROUTE_AUTOWIRE_FORBIDDEN",
    "Patch 13.3 evidence must not auto-wire production route.",
  );

  return Object.freeze({
    originProviderCalls: 1,
    replayProviderCalls: 0,
    plannerCacheSource: "L3_SEMANTIC",
    reentryCacheSource: "L4_REENTRY",
    parityValid: true,
    encryptedPersistentFileCount: Number(
      parity.persistentFiles?.encryptedFileCount,
    ),
    plaintextPersistentFileCount: 0,
    readinessEligible: true,
    readinessDecision: String(gate.decision || ""),
  });
}

function validateCanonicalDataset(dataset = {}) {
  assertCondition(
    dataset.version ===
      "query_candidate_planner_operational_evaluation_dataset_v1",
    "CANONICAL_DATASET_VERSION_INVALID",
    "Canonical dataset version must match Patch 15.1.",
  );
  assertCondition(
    dataset.datasetId === "beebeeai_query_candidate_cost_cache_latency_core_v1",
    "CANONICAL_DATASET_ID_INVALID",
    "Canonical dataset ID must match Patch 15.1.",
  );
  assertCondition(
    dataset.benchmarkMode === "SYNTHETIC_DETERMINISTIC_NO_PROVIDER_CALLS",
    "CANONICAL_DATASET_MODE_INVALID",
    "Source dataset must remain the Patch 15.1 deterministic benchmark.",
  );
  assertCondition(
    Array.isArray(dataset.executions) && dataset.executions.length >= 25,
    "CANONICAL_EXECUTION_SAMPLE_TOO_SMALL",
    "Patch 15.1 canonical dataset requires at least 25 executions.",
  );
  assertCondition(
    Array.isArray(dataset.lifecycleEvents) &&
      dataset.lifecycleEvents.length >= 15,
    "CANONICAL_LIFECYCLE_SAMPLE_TOO_SMALL",
    "Patch 15.1 canonical dataset requires at least 15 lifecycle events.",
  );
  assertCondition(
    dataset.guardrails?.providerCallsExecuted === 0,
    "CANONICAL_DATASET_PROVIDER_CALL_GUARDRAIL_INVALID",
    "Canonical dataset must execute zero provider calls.",
  );
  assertCondition(
    dataset.guardrails?.productionTraffic === false,
    "CANONICAL_DATASET_PRODUCTION_TRAFFIC_FORBIDDEN",
    "Canonical dataset must not be production traffic.",
  );
  return true;
}

function deriveCanonicalEvaluationDataset({
  dataset,
  pricingPolicy,
  liveParityReadiness,
} = {}) {
  validateCanonicalDataset(dataset);
  validatePricingPolicy(pricingPolicy);
  const liveParity = validateLiveParityReadiness(liveParityReadiness);

  const sourceSnapshot = stableStringify(dataset);
  const derived = clone(dataset);
  let strippedObservedCostCount = 0;
  let providerCalledExecutionCount = 0;

  for (const execution of derived.executions) {
    if (execution?.provider?.called === true) {
      providerCalledExecutionCount += 1;
      if (
        Object.prototype.hasOwnProperty.call(
          execution.provider,
          "observedCostMicrousd",
        )
      ) {
        delete execution.provider.observedCostMicrousd;
        strippedObservedCostCount += 1;
      }
    }
  }

  assertCondition(
    providerCalledExecutionCount > 0,
    "CANONICAL_PROVIDER_CALL_ROWS_REQUIRED",
    "Canonical dataset must contain provider-called benchmark rows.",
  );

  assertCondition(
    stableStringify(dataset) === sourceSnapshot,
    "CANONICAL_SOURCE_DATASET_MUTATED",
    "Source Patch 15.1 dataset was mutated.",
  );

  derived.benchmarkMode = BENCHMARK_MODE;
  derived.compatibility = {
    version: INPUT_VERSION,
    predecessorMode: "PATCH_E_X_EXCLUDED_NO_COLLECTION_SUMMARY",
    sourceDatasetMode: "SYNTHETIC_DETERMINISTIC_NO_PROVIDER_CALLS",
    pricingMode: "APPROVED_ACTUAL",
    actualPricingAppliedByEvaluator: true,
    observedSyntheticCostRemovedFromProviderRows: true,
    strippedObservedCostCount,
    providerCalledExecutionCount,
    actualLiveProviderParityEvidence: true,
    actualOperationalTelemetry: false,
    currentProductionTrafficMeasurement: false,
    canaryPromotionEvidence: false,
    productionPromotionEvidence: false,
    collectorRequired: false,
    providerCallsExecutedByPreparation: 0,
    sourceDatasetSha256: sha256Json(dataset),
    pricingPolicySha256: sha256Json(pricingPolicy),
    liveParityReadinessSha256: sha256Json(liveParityReadiness),
    liveParity,
  };

  return Object.freeze({
    version: INPUT_VERSION,
    benchmarkMode: BENCHMARK_MODE,
    dataset: derived,
    source: Object.freeze({
      sourceDatasetSha256: sha256Json(dataset),
      pricingPolicySha256: sha256Json(pricingPolicy),
      liveParityReadinessSha256: sha256Json(liveParityReadiness),
    }),
    guardrails: Object.freeze({
      evaluationOnly: true,
      actualOperationalTelemetry: false,
      productionPromotionAuthorized: false,
      productionCandidateMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      collectorEnabledByThisOperation: false,
      internalCanaryEnabledByThisOperation: false,
      providerCallsExecutedByThisOperation: 0,
    }),
  });
}

module.exports = Object.freeze({
  INPUT_VERSION,
  BENCHMARK_MODE,
  stableStringify,
  sha256Json,
  validatePricingPolicy,
  validateLiveParityReadiness,
  validateCanonicalDataset,
  deriveCanonicalEvaluationDataset,
});
