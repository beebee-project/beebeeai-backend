const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const BUNDLE_VERSION =
  "query_candidate_planner_internal_canary_evidence_candidate_v1";

const REVIEW_DECISION = "ELIGIBLE_FOR_INTERNAL_ALLOWLIST_CANARY_REVIEW";

const SCOPE = "INTERNAL_ALLOWLIST_CANARY_ONLY";

const EXPECTED = Object.freeze({
  evaluatorSha256:
    "2461A48972A8F771E6D49911D70079009E62148658C17EEF986CA3E01972208D",
  pricingFileSha256:
    "0E81028C30E535CB7F047CD4F05FB3FD96D5345636F9E0ED055DD4A6E765FA0D",
  sourceThresholdFileSha256:
    "4F266D07D4C3DE049FE83167788D5B1DABD28717450A76C3607F0FBFA3DA4E1D",
  historicalReadinessFileSha256:
    "33B70E7B4278CBC7E6F66D10CC6AA0F8FA7219A46E553EAD70612494E654F7D5",
  finalBaselineSha256:
    "0c59e08cead5a81d84abd4159aedd34d21666898d6d637c58aed7616ab62730f",
  thresholds: Object.freeze({
    averageCostMicrousdMax: 2600,
    providerCallAverageCostMicrousdMax: 6500,
    monthlyProjectedCostMicrousdMax: 26000000,
    cacheCostAvoidanceRateMin: 0.59,
    providerCallRateMax: 0.4,
    warmAverageCostMicrousdMax: 0,
  }),
  costConsistency: Object.freeze({
    providerCostMicrousd: 48900,
    avoidedByCacheMicrousd: 72600,
    cacheCostAvoidanceRate: 0.597531,
  }),
  headroom: Object.freeze({
    providerCostMaxMicrousd: 5380,
    headroomRate: 0.2,
    rawProviderCeilingMicrousd: 6456,
    providerCallAverageCostMicrousdMax: 6500,
    averageCostMicrousdMax: 2600,
    monthlyProjectedCostMicrousdMax: 26000000,
  }),
});

function fail(code, message = code) {
  const error = new Error(message);
  error.code = code;
  throw error;
}

function assert(condition, code, message) {
  if (!condition) fail(code, message);
}

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (!isObject(value)) return value;
  const out = {};
  for (const key of Object.keys(value).sort()) {
    out[key] = stable(value[key]);
  }
  return out;
}

function canonicalJson(value) {
  return JSON.stringify(stable(value));
}

function sha256Buffer(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex").toUpperCase();
}

function sha256File(file) {
  return sha256Buffer(fs.readFileSync(path.resolve(file)));
}

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(canonicalJson(value))
    .digest("hex")
    .toUpperCase();
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(path.resolve(file), "utf8"));
}

function basename(file) {
  return path.basename(path.resolve(file));
}

function gitHeadFileSha256(file) {
  const absolute = path.resolve(file);
  const relative = path.relative(process.cwd(), absolute).replace(/\\/g, "/");
  let data;
  try {
    data = execFileSync("git", ["show", `HEAD:${relative}`], {
      encoding: null,
    });
  } catch {
    fail(
      "GIT_HEAD_FILE_UNRESOLVED",
      `Unable to resolve Git HEAD file: ${relative}`,
    );
  }
  return sha256Buffer(data);
}

function evaluatorIdentity(
  evaluatorFile = "automation/queryCandidatePlannerCostCacheLatencyEvaluator.js",
) {
  const worktreeSha256 = sha256File(evaluatorFile);
  const headSha256 = gitHeadFileSha256(evaluatorFile);
  return Object.freeze({
    filename: basename(evaluatorFile),
    worktreeSha256,
    headSha256,
    worktreeEqualsHead: worktreeSha256 === headSha256,
  });
}

function validatePricing(pricing, pricingFileSha256) {
  assert(
    pricingFileSha256 === EXPECTED.pricingFileSha256,
    "APPROVED_PRICING_FILE_DRIFT",
  );
  assert(
    pricing.version === "query_candidate_planner_cost_pricing_policy_v1",
    "APPROVED_PRICING_VERSION_INVALID",
  );
  assert(pricing.mode === "APPROVED_ACTUAL", "APPROVED_PRICING_MODE_INVALID");
  assert(
    pricing.rateUnit === "MICROUSD_PER_MILLION_TOKENS",
    "APPROVED_PRICING_RATE_UNIT_INVALID",
  );
  assert(
    pricing.guardrails?.approvedByOperator === true,
    "APPROVED_PRICING_OPERATOR_APPROVAL_REQUIRED",
  );
  assert(
    pricing.guardrails?.productionBillingAuthority === false,
    "PRODUCTION_BILLING_AUTHORITY_FORBIDDEN",
  );

  const rates = pricing.models?.semantic_profiler_default;
  assert(rates, "APPROVED_TERRA_RATES_MISSING");
  assert(
    Number(rates.inputMicrousdPerMillionTokens) === 2000000,
    "APPROVED_TERRA_INPUT_RATE_DRIFT",
  );
  assert(
    Number(rates.outputMicrousdPerMillionTokens) === 12000000,
    "APPROVED_TERRA_OUTPUT_RATE_DRIFT",
  );
}

function validateHistoricalReadiness(readiness, fileSha256) {
  assert(
    fileSha256 === EXPECTED.historicalReadinessFileSha256,
    "HISTORICAL_READINESS_FILE_DRIFT",
  );
  assert(
    readiness.version ===
      "query_candidate_planner_live_cache_parity_readiness_evidence_v1",
    "HISTORICAL_READINESS_VERSION_INVALID",
  );
  assert(
    readiness.origin?.invocationStatus === "CALLED" &&
      Number(readiness.origin?.providerCallCount) === 1,
    "HISTORICAL_ORIGIN_EVIDENCE_INVALID",
  );
  assert(
    readiness.replay?.invocationStatus === "CACHE_HIT" &&
      Number(readiness.replay?.providerCallCount) === 0,
    "HISTORICAL_REPLAY_EVIDENCE_INVALID",
  );
  assert(
    readiness.replay?.cache?.plannerResolution?.source === "L3_SEMANTIC" &&
      readiness.replay?.cache?.reentry?.source === "L4_REENTRY",
    "HISTORICAL_CACHE_SOURCE_INVALID",
  );
  assert(readiness.parityAudit?.valid === true, "HISTORICAL_PARITY_INVALID");
  assert(
    Number(readiness.parityAudit?.persistentFiles?.encryptedFileCount) === 3 &&
      Number(readiness.parityAudit?.persistentFiles?.plaintextFileCount) === 0,
    "HISTORICAL_PERSISTENCE_SECURITY_INVALID",
  );
  assert(
    readiness.readinessGate?.eligible === true,
    "HISTORICAL_READINESS_NOT_ELIGIBLE",
  );
  assert(
    readiness.recovery?.historicalEvidenceOnly === true &&
      readiness.recovery?.currentOperationalTelemetry === false &&
      Number(readiness.recovery?.providerCallsExecutedByRecovery) === 0,
    "HISTORICAL_RECOVERY_GUARDRAIL_INVALID",
  );
  assert(
    readiness.recovery?.productionPromotionAuthorized === false,
    "HISTORICAL_PROMOTION_AUTHORIZATION_FORBIDDEN",
  );
}

function validateSourceThresholdPolicy(policy, fileSha256) {
  assert(
    fileSha256 === EXPECTED.sourceThresholdFileSha256,
    "SOURCE_THRESHOLD_FILE_DRIFT",
  );
  assert(
    policy.version ===
      "query_candidate_planner_operational_threshold_policy_v1",
    "SOURCE_THRESHOLD_VERSION_INVALID",
  );
  assert(
    Number(policy.thresholds?.providerCallRateMax) === 0.4 &&
      Number(policy.thresholds?.cacheCostAvoidanceRateMin) === 0.59 &&
      Number(policy.thresholds?.warmAverageCostMicrousdMax) === 0,
    "SOURCE_THRESHOLD_PRESERVED_CONTRACT_DRIFT",
  );
  assert(
    policy.guardrails?.evaluationOnly === true &&
      policy.guardrails?.promotionAuthorized === false &&
      policy.guardrails?.productionMergeAuthorized === false,
    "SOURCE_THRESHOLD_GUARDRAIL_INVALID",
  );
}

function validateConsistentInput(input) {
  assert(
    input.version ===
      "query_candidate_planner_e_x_compatibility_canonical_evaluation_input_v1",
    "CONSISTENT_INPUT_VERSION_INVALID",
  );
  assert(
    input.costConsistency?.providerCostMicrousd ===
      EXPECTED.costConsistency.providerCostMicrousd,
    "PROVIDER_COST_EVIDENCE_DRIFT",
  );
  assert(
    input.costConsistency?.avoidedByCacheMicrousd ===
      EXPECTED.costConsistency.avoidedByCacheMicrousd,
    "AVOIDED_CACHE_COST_EVIDENCE_DRIFT",
  );
  assert(
    Number(input.costConsistency?.cacheCostAvoidanceRate) ===
      EXPECTED.costConsistency.cacheCostAvoidanceRate,
    "CACHE_AVOIDANCE_RATE_EVIDENCE_DRIFT",
  );
  assert(
    input.guardrails?.actualOperationalTelemetry === false &&
      input.guardrails?.productionPromotionAuthorized === false &&
      Number(input.guardrails?.providerCallsExecutedByThisOperation) === 0,
    "CONSISTENT_INPUT_GUARDRAIL_INVALID",
  );
}

function validateRecalibratedThresholdPolicy(policy) {
  const t = policy.thresholds || {};
  for (const [key, expected] of Object.entries(EXPECTED.thresholds)) {
    assert(
      Number(t[key]) === expected,
      "RECALIBRATED_THRESHOLD_DRIFT",
      `${key} expected ${expected} but found ${t[key]}`,
    );
  }

  assert(
    policy.guardrails?.evaluationOnly === true &&
      policy.guardrails?.promotionAuthorized === false &&
      policy.guardrails?.productionMergeAuthorized === false,
    "RECALIBRATED_THRESHOLD_GUARDRAIL_INVALID",
  );
}

function validateRecalibrationEvidence(evidence) {
  assert(
    evidence.version ===
      "query_candidate_planner_actual_pricing_absolute_cost_threshold_recalibration_v1",
    "RECALIBRATION_EVIDENCE_VERSION_INVALID",
  );

  const d = evidence.providerCostDistribution || {};
  const t = evidence.derivedThresholds || {};

  assert(
    Number(d.sampleCount) === 10 &&
      Number(d.averageMicrousd) === 4890 &&
      Number(d.p95Microusd) === 5380 &&
      Number(d.maxMicrousd) === EXPECTED.headroom.providerCostMaxMicrousd,
    "PROVIDER_COST_DISTRIBUTION_DRIFT",
  );

  assert(
    Number(evidence.headroomRate) === EXPECTED.headroom.headroomRate &&
      Number(evidence.rawProviderCeilingMicrousd) ===
        EXPECTED.headroom.rawProviderCeilingMicrousd,
    "HEADROOM_EVIDENCE_DRIFT",
  );

  assert(
    Number(t.providerCallAverageCostMicrousdMax) ===
      EXPECTED.headroom.providerCallAverageCostMicrousdMax &&
      Number(t.averageCostMicrousdMax) ===
        EXPECTED.headroom.averageCostMicrousdMax &&
      Number(t.monthlyProjectedCostMicrousdMax) ===
        EXPECTED.headroom.monthlyProjectedCostMicrousdMax,
    "DERIVED_THRESHOLD_EVIDENCE_DRIFT",
  );

  assert(
    evidence.sourceThresholdPolicyModified === false &&
      evidence.evaluatorModified === false &&
      Number(evidence.providerCallsExecutedByRecalibration) === 0 &&
      evidence.actualOperationalTelemetry === false &&
      evidence.productionPromotionAuthorized === false,
    "RECALIBRATION_EVIDENCE_GUARDRAIL_INVALID",
  );
}

function validateOperationalReport(report) {
  assert(
    report.decision === "EVALUATION_PASS",
    "OPERATIONAL_EVALUATION_NOT_PASS",
  );

  const failed = Array.isArray(report.thresholdResults)
    ? report.thresholdResults.filter((item) => item?.passed === false)
    : [];

  assert(failed.length === 0, "OPERATIONAL_FAILED_CHECKS_PRESENT");

  const costAvoidance = (report.thresholdResults || []).find(
    (item) => item?.metric === "cost.cacheCostAvoidanceRate",
  );

  assert(costAvoidance?.passed === true, "CACHE_AVOIDANCE_THRESHOLD_NOT_PASS");
  assert(
    Number(costAvoidance.actual) ===
      EXPECTED.costConsistency.cacheCostAvoidanceRate &&
      Number(costAvoidance.threshold) ===
        EXPECTED.thresholds.cacheCostAvoidanceRateMin,
    "CACHE_AVOIDANCE_THRESHOLD_RESULT_DRIFT",
  );
}

function validateAssessment(assessment) {
  assert(
    assessment.decision === "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS",
    "ABSOLUTE_COST_ASSESSMENT_NOT_PASS",
  );
  assert(
    Number(assessment.failedCheckCount) === 0 &&
      Number(assessment.absoluteCostPassCount) === 3 &&
      Number(assessment.absoluteCostFailureCount) === 0 &&
      assessment.cacheCostAvoidancePassed === true,
    "ABSOLUTE_COST_ASSESSMENT_RESULT_INVALID",
  );
  assert(
    assessment.sourceThresholdPolicyModified === false &&
      assessment.evaluatorModified === false &&
      Number(assessment.providerCallsExecutedByAssessment) === 0 &&
      assessment.productionPromotionAuthorized === false,
    "ABSOLUTE_COST_ASSESSMENT_GUARDRAIL_INVALID",
  );
}

function validateBaseline(baseline) {
  assert(
    baseline.baselineSha256 === EXPECTED.finalBaselineSha256,
    "FINAL_BASELINE_SHA_DRIFT",
  );
  assert(
    baseline.methodology?.actualOperationalTelemetry === false,
    "BASELINE_ACTUAL_TELEMETRY_CLAIM_FORBIDDEN",
  );
  assert(
    baseline.guardrails?.promotionAuthorized === false,
    "BASELINE_PROMOTION_AUTHORIZATION_FORBIDDEN",
  );
}

function buildCandidate({
  pricingFile,
  readinessFile,
  consistentInputFile,
  sourceThresholdFile,
  recalibratedThresholdFile,
  recalibrationEvidenceFile,
  operationalReportFile,
  assessmentFile,
  baselineFile,
  evaluatorFile = "automation/queryCandidatePlannerCostCacheLatencyEvaluator.js",
} = {}) {
  const required = {
    pricingFile,
    readinessFile,
    consistentInputFile,
    sourceThresholdFile,
    recalibratedThresholdFile,
    recalibrationEvidenceFile,
    operationalReportFile,
    assessmentFile,
    baselineFile,
  };

  for (const [name, file] of Object.entries(required)) {
    assert(
      typeof file === "string" && file.trim().length > 0,
      "REQUIRED_INPUT_PATH_MISSING",
      name,
    );
    assert(
      fs.existsSync(path.resolve(file)),
      "REQUIRED_INPUT_FILE_MISSING",
      file,
    );
  }

  const hashes = Object.fromEntries(
    Object.entries(required).map(([name, file]) => [name, sha256File(file)]),
  );

  const pricing = readJson(pricingFile);
  const readiness = readJson(readinessFile);
  const consistentInput = readJson(consistentInputFile);
  const sourceThreshold = readJson(sourceThresholdFile);
  const recalibratedThreshold = readJson(recalibratedThresholdFile);
  const recalibrationEvidence = readJson(recalibrationEvidenceFile);
  const operationalReport = readJson(operationalReportFile);
  const assessment = readJson(assessmentFile);
  const baseline = readJson(baselineFile);

  const evaluator = evaluatorIdentity(evaluatorFile);

  assert(evaluator.worktreeEqualsHead === true, "EVALUATOR_WORKTREE_DRIFT");
  assert(
    evaluator.worktreeSha256 === EXPECTED.evaluatorSha256,
    "EVALUATOR_SHA_DRIFT",
  );

  validatePricing(pricing, hashes.pricingFile);
  validateHistoricalReadiness(readiness, hashes.readinessFile);
  validateConsistentInput(consistentInput);
  validateSourceThresholdPolicy(sourceThreshold, hashes.sourceThresholdFile);
  validateRecalibratedThresholdPolicy(recalibratedThreshold);
  validateRecalibrationEvidence(recalibrationEvidence);
  validateOperationalReport(operationalReport);
  validateAssessment(assessment);
  validateBaseline(baseline);

  const candidatePayload = {
    version: BUNDLE_VERSION,
    scope: SCOPE,
    sourcePatch: "15.3.2-F.1.4",
    purpose: "SANITIZED_EVIDENCE_FOR_MANUAL_INTERNAL_ALLOWLIST_CANARY_REVIEW",
    evaluation: {
      operationalDecision: operationalReport.decision,
      assessmentDecision: assessment.decision,
      failedCheckCount: Number(assessment.failedCheckCount),
      absoluteCostPassCount: Number(assessment.absoluteCostPassCount),
      absoluteCostFailureCount: Number(assessment.absoluteCostFailureCount),
      cacheCostAvoidancePassed: assessment.cacheCostAvoidancePassed === true,
      cacheCostAvoidanceActual: Number(assessment.cacheCostAvoidanceActual),
      cacheCostAvoidanceThreshold: Number(
        assessment.cacheCostAvoidanceThreshold,
      ),
    },
    thresholds: {
      averageCostMicrousdMax: Number(
        recalibratedThreshold.thresholds.averageCostMicrousdMax,
      ),
      providerCallAverageCostMicrousdMax: Number(
        recalibratedThreshold.thresholds.providerCallAverageCostMicrousdMax,
      ),
      monthlyProjectedCostMicrousdMax: Number(
        recalibratedThreshold.thresholds.monthlyProjectedCostMicrousdMax,
      ),
      cacheCostAvoidanceRateMin: Number(
        recalibratedThreshold.thresholds.cacheCostAvoidanceRateMin,
      ),
      providerCallRateMax: Number(
        recalibratedThreshold.thresholds.providerCallRateMax,
      ),
      warmAverageCostMicrousdMax: Number(
        recalibratedThreshold.thresholds.warmAverageCostMicrousdMax,
      ),
    },
    pricing: {
      mode: pricing.mode,
      rateUnit: pricing.rateUnit,
      modelKey: "semantic_profiler_default",
      inputMicrousdPerMillionTokens: Number(
        pricing.models.semantic_profiler_default.inputMicrousdPerMillionTokens,
      ),
      outputMicrousdPerMillionTokens: Number(
        pricing.models.semantic_profiler_default.outputMicrousdPerMillionTokens,
      ),
      approvedByOperator: true,
      productionBillingAuthority: false,
    },
    costConsistency: {
      providerCostMicrousd: Number(
        consistentInput.costConsistency.providerCostMicrousd,
      ),
      avoidedByCacheMicrousd: Number(
        consistentInput.costConsistency.avoidedByCacheMicrousd,
      ),
      cacheCostAvoidanceRate: Number(
        consistentInput.costConsistency.cacheCostAvoidanceRate,
      ),
      providerCostSampleCount: Number(
        recalibrationEvidence.providerCostDistribution.sampleCount,
      ),
      providerCostAverageMicrousd: Number(
        recalibrationEvidence.providerCostDistribution.averageMicrousd,
      ),
      providerCostP95Microusd: Number(
        recalibrationEvidence.providerCostDistribution.p95Microusd,
      ),
      providerCostMaxMicrousd: Number(
        recalibrationEvidence.providerCostDistribution.maxMicrousd,
      ),
      headroomRate: Number(recalibrationEvidence.headroomRate),
      rawProviderCeilingMicrousd: Number(
        recalibrationEvidence.rawProviderCeilingMicrousd,
      ),
    },
    liveParity: {
      historicalLiveProviderEvidence: true,
      originProviderCallCount: Number(readiness.origin.providerCallCount),
      replayProviderCallCount: Number(readiness.replay.providerCallCount),
      plannerResolutionSource: readiness.replay.cache.plannerResolution.source,
      reentrySource: readiness.replay.cache.reentry.source,
      parityValid: readiness.parityAudit.valid === true,
      readinessEligible: readiness.readinessGate.eligible === true,
      encryptedPersistentFileCount: Number(
        readiness.parityAudit.persistentFiles.encryptedFileCount,
      ),
      plaintextPersistentFileCount: Number(
        readiness.parityAudit.persistentFiles.plaintextFileCount,
      ),
      currentOperationalTelemetry: false,
    },
    integrity: {
      evaluator: evaluator,
      finalBaselineSha256: baseline.baselineSha256,
      inputFiles: {
        pricingPolicy: {
          filename: basename(pricingFile),
          sha256: hashes.pricingFile,
        },
        historicalReadiness: {
          filename: basename(readinessFile),
          sha256: hashes.readinessFile,
        },
        actualPricingConsistentInput: {
          filename: basename(consistentInputFile),
          sha256: hashes.consistentInputFile,
        },
        sourceThresholdPolicy: {
          filename: basename(sourceThresholdFile),
          sha256: hashes.sourceThresholdFile,
        },
        recalibratedThresholdPolicy: {
          filename: basename(recalibratedThresholdFile),
          sha256: hashes.recalibratedThresholdFile,
        },
        recalibrationEvidence: {
          filename: basename(recalibrationEvidenceFile),
          sha256: hashes.recalibrationEvidenceFile,
        },
        operationalReport: {
          filename: basename(operationalReportFile),
          sha256: hashes.operationalReportFile,
        },
        assessment: {
          filename: basename(assessmentFile),
          sha256: hashes.assessmentFile,
        },
        finalBaseline: {
          filename: basename(baselineFile),
          sha256: hashes.baselineFile,
        },
      },
    },
    eligibility: {
      decision: REVIEW_DECISION,
      internalCanaryReviewEligible: true,
      manualOperatorApprovalRequired: true,
      immutableApprovalBundleHashBindingRequired: true,
      immutableAllowlistHashBindingRequired: true,
      internalCanaryAuthorized: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
      productionMergeAuthorized: false,
    },
    methodology: {
      actualPricing: true,
      actualHistoricalLiveProviderParityEvidence: true,
      actualOperationalTelemetry: false,
      canonicalBenchmarkEvidence: true,
      internalCanaryEvidence: false,
      productionPromotionEvidence: false,
    },
    guardrails: {
      evaluationOnly: true,
      noGateMutation: true,
      noEnvironmentMutation: true,
      noRouteMutation: true,
      noFeatureFlagMutation: true,
      noKillSwitchMutation: true,
      noAllowlistMutation: true,
      providerCallsExecutedByBundleBuilder: 0,
      privateOutputDoNotCommit: true,
    },
    sanitization: {
      responseIdsIncluded: false,
      rawExecutionsIncluded: false,
      rawRowsIncluded: false,
      tokenUsageRowsIncluded: false,
      immutableAccountIdsIncluded: false,
      allowlistSubjectsIncluded: false,
      environmentValuesIncluded: false,
    },
  };

  const candidatePayloadSha256 = sha256Json(candidatePayload);

  const bundle = Object.freeze({
    ...candidatePayload,
    candidatePayloadSha256,
  });

  const serialized = JSON.stringify(bundle);
  const forbiddenKeys = [
    '"responseId"',
    '"inputTokens"',
    '"outputTokens"',
    '"executions"',
    '"rawRows"',
    '"immutableAccountId"',
    '"allowlistSubjects"',
  ];
  for (const key of forbiddenKeys) {
    assert(
      !serialized.includes(key),
      "SANITIZATION_FAILURE",
      `Forbidden field leaked into bundle: ${key}`,
    );
  }

  return bundle;
}

function verifyCandidate(bundle = {}) {
  assert(bundle.version === BUNDLE_VERSION, "BUNDLE_VERSION_INVALID");
  assert(bundle.scope === SCOPE, "BUNDLE_SCOPE_INVALID");
  assert(
    bundle.evaluation?.operationalDecision === "EVALUATION_PASS",
    "BUNDLE_EVALUATION_NOT_PASS",
  );
  assert(
    bundle.evaluation?.assessmentDecision ===
      "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS",
    "BUNDLE_ASSESSMENT_NOT_PASS",
  );
  assert(
    Number(bundle.evaluation?.failedCheckCount) === 0 &&
      Number(bundle.evaluation?.absoluteCostFailureCount) === 0 &&
      bundle.evaluation?.cacheCostAvoidancePassed === true,
    "BUNDLE_EVALUATION_RESULT_INVALID",
  );
  assert(
    bundle.eligibility?.decision === REVIEW_DECISION &&
      bundle.eligibility?.internalCanaryReviewEligible === true,
    "BUNDLE_REVIEW_ELIGIBILITY_INVALID",
  );
  assert(
    bundle.eligibility?.internalCanaryAuthorized === false &&
      bundle.eligibility?.percentageRolloutAuthorized === false &&
      bundle.eligibility?.productionPromotionAuthorized === false &&
      bundle.eligibility?.productionMergeAuthorized === false,
    "BUNDLE_AUTHORIZATION_BOUNDARY_INVALID",
  );
  assert(
    bundle.methodology?.actualOperationalTelemetry === false &&
      bundle.methodology?.internalCanaryEvidence === false &&
      bundle.methodology?.productionPromotionEvidence === false,
    "BUNDLE_METHODOLOGY_BOUNDARY_INVALID",
  );
  assert(
    bundle.guardrails?.noGateMutation === true &&
      bundle.guardrails?.noEnvironmentMutation === true &&
      bundle.guardrails?.noRouteMutation === true &&
      Number(bundle.guardrails?.providerCallsExecutedByBundleBuilder) === 0,
    "BUNDLE_GUARDRAIL_INVALID",
  );

  const copy = JSON.parse(JSON.stringify(bundle));
  const observed = copy.candidatePayloadSha256;
  delete copy.candidatePayloadSha256;
  assert(observed === sha256Json(copy), "BUNDLE_PAYLOAD_SHA_INVALID");

  return true;
}

module.exports = Object.freeze({
  BUNDLE_VERSION,
  REVIEW_DECISION,
  SCOPE,
  EXPECTED,
  canonicalJson,
  sha256File,
  sha256Json,
  evaluatorIdentity,
  buildCandidate,
  verifyCandidate,
});
