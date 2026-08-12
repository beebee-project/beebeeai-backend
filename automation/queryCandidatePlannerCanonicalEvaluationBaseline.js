const crypto = require("crypto");

const BASELINE_VERSION =
  "query_candidate_planner_e_x_compatibility_evaluation_baseline_v1";

const COST_CHECK_PATTERNS = Object.freeze([
  /cost/i,
  /microusd/i,
  /pricing/i,
  /monthly.*projection/i,
]);

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

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(JSON.stringify(stableValue(value)))
    .digest("hex");
}

function flattenFailures(value, path = [], output = []) {
  if (Array.isArray(value)) {
    value.forEach((child, index) =>
      flattenFailures(child, path.concat(String(index)), output),
    );
    return output;
  }
  if (!isObject(value)) return output;

  const pass =
    value.pass === false ||
    value.passed === false ||
    String(value.status || "").toUpperCase() === "FAIL" ||
    String(value.decision || "").toUpperCase() === "BLOCKED";

  if (pass) {
    output.push({
      path: path.join("."),
      name: String(
        value.name ||
          value.metric ||
          value.check ||
          value.id ||
          value.reason ||
          path.at(-1) ||
          "",
      ),
      reason: String(value.reason || value.message || ""),
    });
  }

  for (const [key, child] of Object.entries(value)) {
    if (isObject(child) || Array.isArray(child)) {
      flattenFailures(child, path.concat(key), output);
    }
  }
  return output;
}

function isCostFailure(item = {}) {
  const text = `${item.path} ${item.name} ${item.reason}`;
  return COST_CHECK_PATTERNS.some((pattern) => pattern.test(text));
}

function buildCanonicalEvaluationBaseline({
  canonicalInput,
  pricingPolicy,
  liveParityReadiness,
  thresholdPolicy,
  operationalReport,
  evaluatorIdentity = {},
} = {}) {
  if (
    canonicalInput?.version !==
    "query_candidate_planner_e_x_compatibility_canonical_evaluation_input_v1"
  ) {
    throw new Error("Invalid canonical compatibility input.");
  }
  if (pricingPolicy?.mode !== "APPROVED_ACTUAL") {
    throw new Error("APPROVED_ACTUAL pricing is required.");
  }
  if (
    thresholdPolicy?.version !==
    "query_candidate_planner_operational_threshold_policy_v1"
  ) {
    throw new Error("Patch 15.1 threshold policy is required.");
  }
  if (
    operationalReport?.version !==
    "query_candidate_planner_cost_cache_latency_evaluation_report_v1"
  ) {
    throw new Error("Patch 15.1 operational report is required.");
  }

  const failures = flattenFailures(operationalReport);
  const costFailures = failures.filter(isCostFailure);
  const nonCostFailures = failures.filter((item) => !isCostFailure(item));
  const reportDecision = String(operationalReport.decision || "");

  let decision = "CANONICAL_EVALUATION_BASELINE_CAPTURED";
  if (reportDecision === "EVALUATION_PASS") {
    decision = "CANONICAL_EVALUATION_BASELINE_PASS";
  } else if (
    reportDecision === "EVALUATION_BLOCKED" &&
    costFailures.length > 0 &&
    nonCostFailures.length === 0
  ) {
    decision = "CANONICAL_EVALUATION_BASELINE_COST_RECALIBRATION_REQUIRED";
  }

  const baseline = {
    version: BASELINE_VERSION,
    decision,
    methodology: {
      mode: "CANONICAL_BENCHMARK_WITH_APPROVED_ACTUAL_PRICING",
      actualPricing: true,
      actualLiveProviderParityEvidence: true,
      actualOperationalTelemetry: false,
      currentProductionTrafficMeasurement: false,
      canonicalLatencyAndLifecycleBenchmark: true,
      patchECollectionSummaryRequired: false,
      patchECollectionSummaryUsed: false,
      patchEXCompatibility: true,
    },
    evaluator: {
      version:
        operationalReport.evaluatorVersion ||
        "query_candidate_planner_cost_cache_latency_evaluator_v1",
      worktreeSha256: String(evaluatorIdentity.worktreeSha256 || ""),
      headSha256: String(evaluatorIdentity.headSha256 || ""),
      worktreeEqualsHead: evaluatorIdentity.worktreeEqualsHead === true,
      providerCallsExecutedByEvaluator: 0,
    },
    operationalEvaluation: {
      reportDecision,
      reportSha256: sha256Json(operationalReport),
      failedCheckCount: failures.length,
      costFailureCount: costFailures.length,
      nonCostFailureCount: nonCostFailures.length,
      costThresholdRecalibrationRequired:
        reportDecision !== "EVALUATION_PASS" && costFailures.length > 0,
      report: operationalReport,
    },
    source: {
      canonicalInputSha256: sha256Json(canonicalInput),
      sourceDatasetSha256: canonicalInput.source?.sourceDatasetSha256 || "",
      pricingPolicySha256: sha256Json(pricingPolicy),
      liveParityReadinessSha256: sha256Json(liveParityReadiness),
      thresholdPolicySha256: sha256Json(thresholdPolicy),
    },
    guardrails: {
      evaluationOnly: true,
      promotionAuthorized: false,
      canaryPromotionEvidence: false,
      productionPromotionEvidence: false,
      productionCandidateMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      collectorEnabledByThisOperation: false,
      internalCanaryEnabledByThisOperation: false,
      providerCallsExecutedByThisOperation: 0,
      privateOutputDoNotCommit: true,
    },
  };

  return Object.freeze({
    ...baseline,
    baselineSha256: sha256Json(baseline),
  });
}

module.exports = Object.freeze({
  BASELINE_VERSION,
  flattenFailures,
  isCostFailure,
  buildCanonicalEvaluationBaseline,
});
