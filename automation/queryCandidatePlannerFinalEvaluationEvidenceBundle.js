const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const BUNDLE_VERSION =
  "query_candidate_planner_final_evaluation_evidence_bundle_v1";
const SCOPE = "INTERNAL_ALLOWLIST_CANARY_BOOTSTRAP_READINESS_ONLY";
const DECISION = "READY_FOR_15_3_3_INTERNAL_ALLOWLIST_CANARY_BOOTSTRAP";

const EXPECTED = Object.freeze({
  finalBaselineSha256:
    "0C59E08CEAD5A81D84ABD4159AEDD34D21666898D6D637C58AED7616AB62730F",
  evaluatorSha256:
    "2461A48972A8F771E6D49911D70079009E62148658C17EEF986CA3E01972208D",
  candidatePayloadSha256:
    "928F6A6E0AA8683D63A5A2CB62199FA460EB84494B119EB7E171000843D484EA",
  candidateFileSha256:
    "ED5A92C1F5809E6BF48A20C30046418F5F90152B87027DFFAD95E71DB77A534A",
  rotationPlanFileSha256:
    "E2B77CE84DA908130E4C35779A9243E9A66598F70B943B496E8470CCA496C17A",
  allowlistSha256:
    "35D88A2074548BB9A6DB6BD3415CEE3CD2024BE9896AE6EC23260DB9B859AB95",
  approvalReceiptPayloadSha256:
    "4F5BA14C79CCB0ADD5DED476335729E59546A244A301FDF12897B14A7A09EF81",
  approvalReceiptFileSha256:
    "3367FC2ECEE9FE21D943C26221701FFA04697E75344D0462FC222E7CF3767FE6",
  approvalBindingGateSha256:
    "ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7",
  composedServiceSha256:
    "1A61F219ADF49BD863B84C5B8C4DB02158E901E7EDA864AC551656A4A7E75C8F",
});

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (!isObject(value)) return value;
  const out = {};
  for (const key of Object.keys(value).sort()) out[key] = stable(value[key]);
  return out;
}

function canonicalJson(value) {
  return JSON.stringify(stable(value));
}

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(canonicalJson(value))
    .digest("hex")
    .toUpperCase();
}

function sha256File(file) {
  return crypto
    .createHash("sha256")
    .update(fs.readFileSync(path.resolve(file)))
    .digest("hex")
    .toUpperCase();
}

function normalizeSha256(value) {
  const normalized = String(value || "")
    .trim()
    .toUpperCase();
  return /^[A-F0-9]{64}$/.test(normalized) && !/^0{64}$/.test(normalized)
    ? normalized
    : "";
}

function fail(code) {
  const error = new Error(code);
  error.code = code;
  throw error;
}

function assert(condition, code) {
  if (!condition) fail(code);
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(path.resolve(file), "utf8"));
}

function validateCandidate(candidate, candidateFile) {
  assert(
    candidate?.candidatePayloadSha256 === EXPECTED.candidatePayloadSha256,
    "G_CANDIDATE_PAYLOAD_SHA_MISMATCH",
  );
  assert(
    sha256File(candidateFile) === EXPECTED.candidateFileSha256,
    "G_CANDIDATE_FILE_SHA_MISMATCH",
  );
  assert(
    candidate?.evaluation?.operationalDecision === "EVALUATION_PASS",
    "G_OPERATIONAL_EVALUATION_NOT_PASS",
  );
  assert(
    candidate?.evaluation?.assessmentDecision ===
      "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS",
    "G_ASSESSMENT_NOT_PASS",
  );
  assert(
    Number(candidate?.evaluation?.failedCheckCount) === 0,
    "G_FAILED_CHECK_COUNT_NONZERO",
  );
  assert(
    Number(candidate?.evaluation?.absoluteCostFailureCount) === 0,
    "G_ABSOLUTE_COST_FAILURE_COUNT_NONZERO",
  );
  assert(
    candidate?.evaluation?.cacheCostAvoidancePassed === true,
    "G_CACHE_COST_AVOIDANCE_NOT_PASS",
  );
  assert(
    candidate?.methodology?.actualOperationalTelemetry === false,
    "G_UNEXPECTED_OPERATIONAL_TELEMETRY_CLAIM",
  );
  assert(
    candidate?.integrity?.evaluator?.worktreeEqualsHead === true,
    "G_EVALUATOR_WORKTREE_DRIFT",
  );
  assert(
    normalizeSha256(candidate?.integrity?.evaluator?.worktreeSha256) ===
      EXPECTED.evaluatorSha256,
    "G_EVALUATOR_SHA_MISMATCH",
  );
  assert(
    normalizeSha256(candidate?.integrity?.finalBaselineSha256) ===
      EXPECTED.finalBaselineSha256,
    "G_FINAL_BASELINE_SHA_MISMATCH",
  );
}

function validateRotation(rotation, rotationFile) {
  assert(
    sha256File(rotationFile) === EXPECTED.rotationPlanFileSha256,
    "G_ROTATION_PLAN_FILE_SHA_MISMATCH",
  );
  assert(
    rotation?.decision === "READY_FOR_LOCAL_APPROVAL_REBINDING",
    "G_ROTATION_DECISION_INVALID",
  );
  assert(rotation?.valid === true, "G_ROTATION_NOT_VALID");
  assert(
    normalizeSha256(rotation?.proposedAllowlistSha256) ===
      EXPECTED.allowlistSha256,
    "G_ROTATED_ALLOWLIST_SHA_MISMATCH",
  );
  assert(
    rotation?.approvalRebinding?.f14CandidatePreserved === true,
    "G_F14_CANDIDATE_NOT_PRESERVED",
  );
  assert(
    rotation?.guardrails?.actualOperationalTelemetry === false,
    "G_ROTATION_TELEMETRY_BOUNDARY_INVALID",
  );
}

function validateReceipt(receipt, receiptFile) {
  assert(
    sha256File(receiptFile) === EXPECTED.approvalReceiptFileSha256,
    "G_RECEIPT_FILE_SHA_MISMATCH",
  );
  assert(
    normalizeSha256(receipt?.approvalReceiptPayloadSha256) ===
      EXPECTED.approvalReceiptPayloadSha256,
    "G_RECEIPT_PAYLOAD_SHA_MISMATCH",
  );
  assert(
    receipt?.decision === "INTERNAL_ALLOWLIST_CANARY_MANUAL_APPROVAL_GRANTED",
    "G_RECEIPT_DECISION_INVALID",
  );
  assert(
    normalizeSha256(receipt?.immutableBindings?.candidatePayloadSha256) ===
      EXPECTED.candidatePayloadSha256,
    "G_RECEIPT_CANDIDATE_BINDING_MISMATCH",
  );
  assert(
    normalizeSha256(receipt?.immutableBindings?.allowlistSha256) ===
      EXPECTED.allowlistSha256,
    "G_RECEIPT_ALLOWLIST_BINDING_MISMATCH",
  );
  assert(
    receipt?.manualApproval?.approvedByOperator === true,
    "G_MANUAL_APPROVAL_MISSING",
  );
  assert(
    receipt?.manualApproval?.approvalIsRuntimeActivation === false,
    "G_MANUAL_APPROVAL_RUNTIME_BOUNDARY_INVALID",
  );
  assert(
    receipt?.authorizationBoundary?.runtimeCanaryAuthorized === false,
    "G_RECEIPT_RUNTIME_ALREADY_AUTHORIZED",
  );
  assert(
    receipt?.authorizationBoundary?.percentageRolloutAuthorized === false,
    "G_RECEIPT_ROLLOUT_ALREADY_AUTHORIZED",
  );
  assert(
    receipt?.authorizationBoundary?.productionPromotionAuthorized === false,
    "G_RECEIPT_PROMOTION_ALREADY_AUTHORIZED",
  );
}

function buildFinalEvaluationEvidenceBundle({
  candidateFile,
  rotationFile,
  receiptFile,
  approvalBindingGateFile,
  composedServiceFile,
} = {}) {
  for (const [name, file] of Object.entries({
    candidateFile,
    rotationFile,
    receiptFile,
    approvalBindingGateFile,
    composedServiceFile,
  })) {
    assert(
      typeof file === "string" && file.trim(),
      `G_${name.toUpperCase()}_REQUIRED`,
    );
    assert(
      fs.existsSync(path.resolve(file)),
      `G_${name.toUpperCase()}_MISSING`,
    );
  }

  const candidate = readJson(candidateFile);
  const rotation = readJson(rotationFile);
  const receipt = readJson(receiptFile);

  validateCandidate(candidate, candidateFile);
  validateRotation(rotation, rotationFile);
  validateReceipt(receipt, receiptFile);

  assert(
    sha256File(approvalBindingGateFile) === EXPECTED.approvalBindingGateSha256,
    "G_APPROVAL_BINDING_GATE_SHA_MISMATCH",
  );
  assert(
    sha256File(composedServiceFile) === EXPECTED.composedServiceSha256,
    "G_COMPOSED_SERVICE_SHA_MISMATCH",
  );

  const payload = {
    version: BUNDLE_VERSION,
    scope: SCOPE,
    sourcePatch: "15.3.2-G",
    decision: DECISION,
    failClosed: true,
    immutableBindings: {
      finalBaselineSha256: EXPECTED.finalBaselineSha256,
      evaluatorSha256: EXPECTED.evaluatorSha256,
      candidatePayloadSha256: EXPECTED.candidatePayloadSha256,
      candidateFileSha256: EXPECTED.candidateFileSha256,
      rotationPlanFileSha256: EXPECTED.rotationPlanFileSha256,
      allowlistSha256: EXPECTED.allowlistSha256,
      approvalReceiptPayloadSha256: EXPECTED.approvalReceiptPayloadSha256,
      approvalReceiptFileSha256: EXPECTED.approvalReceiptFileSha256,
      approvalBindingGateSha256: EXPECTED.approvalBindingGateSha256,
      composedServiceSha256: EXPECTED.composedServiceSha256,
    },
    evaluationSnapshot: {
      operationalDecision: candidate.evaluation.operationalDecision,
      assessmentDecision: candidate.evaluation.assessmentDecision,
      failedCheckCount: Number(candidate.evaluation.failedCheckCount),
      absoluteCostFailureCount: Number(
        candidate.evaluation.absoluteCostFailureCount,
      ),
      cacheCostAvoidancePassed:
        candidate.evaluation.cacheCostAvoidancePassed === true,
      actualOperationalTelemetry: false,
    },
    readiness: {
      eligible: true,
      decision: DECISION,
      bootstrapOnly: true,
      internalAllowlistOnly: true,
      rolloutPercent: 0,
      manualApprovalBound: true,
      currentCodeBaselineVerified: true,
      actualTrafficEvidenceRequiredFor15_3_4: true,
      guardrails: {
        failClosed: true,
        productionRouteAutoWired: false,
        productionCandidateMergeAutoAuthorized: false,
        productionReadyAssignmentAllowed: false,
        percentageRolloutAuthorized: false,
        productionPromotionAuthorized: false,
      },
    },
    legacy15_3EvidenceContract: {
      version: "query_candidate_planner_internal_canary_evidence_bundle_v1",
      realShadowTrafficRequired: true,
      actualTrafficRequired: true,
      syntheticForbidden: true,
      satisfiedByThisBundle: false,
      substitutionForbidden: true,
      reason: "G_IS_PRE_CANARY_BOOTSTRAP_READINESS_NOT_REAL_SHADOW_TRAFFIC",
    },
    authorizationBoundary: {
      internalCanaryBootstrapReadinessEstablished: true,
      runtimeAutoActivationAuthorized: false,
      actualInternalUserExposureAuthorized: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
      productionMergeAutoAuthorized: false,
    },
    nextStep: {
      patch: "15.3.3",
      action: "INTERNAL_USER_ACTUAL_OPERATIONAL_EXPOSURE",
      requiresExplicitRuntimeBootstrapIntegration: true,
      requiresActualOperationalTelemetryCollection: true,
      requiresKillSwitchAndPrimaryFallback: true,
      requiresAllowlistOnly: true,
      requiresRolloutPercentZero: true,
    },
    guardrails: {
      rawIdentityIncluded: false,
      rawRowsIncluded: false,
      providerCallsExecutedByBundleBuilder: 0,
      actualOperationalTelemetry: false,
      railwayModified: false,
      environmentModified: false,
      routeModified: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    },
  };

  const bundlePayloadSha256 = sha256Json(payload);
  return Object.freeze({ ...payload, bundlePayloadSha256 });
}

function verifyFinalEvaluationEvidenceBundle(bundle = {}) {
  assert(bundle.version === BUNDLE_VERSION, "G_BUNDLE_VERSION_INVALID");
  assert(bundle.scope === SCOPE, "G_BUNDLE_SCOPE_INVALID");
  assert(bundle.decision === DECISION, "G_BUNDLE_DECISION_INVALID");
  assert(bundle.failClosed === true, "G_BUNDLE_FAIL_CLOSED_REQUIRED");

  for (const [key, expected] of Object.entries(EXPECTED)) {
    const observed = normalizeSha256(bundle?.immutableBindings?.[key]);
    assert(observed === expected, `G_BUNDLE_BINDING_MISMATCH_${key}`);
  }

  assert(
    bundle?.evaluationSnapshot?.operationalDecision === "EVALUATION_PASS",
    "G_BUNDLE_OPERATIONAL_DECISION_INVALID",
  );
  assert(
    bundle?.evaluationSnapshot?.failedCheckCount === 0,
    "G_BUNDLE_FAILED_CHECK_COUNT_NONZERO",
  );
  assert(
    bundle?.evaluationSnapshot?.actualOperationalTelemetry === false,
    "G_BUNDLE_TELEMETRY_CLAIM_INVALID",
  );

  assert(
    bundle?.readiness?.eligible === true,
    "G_BUNDLE_READINESS_NOT_ELIGIBLE",
  );
  assert(
    bundle?.readiness?.bootstrapOnly === true,
    "G_BUNDLE_BOOTSTRAP_ONLY_REQUIRED",
  );
  assert(
    bundle?.readiness?.internalAllowlistOnly === true,
    "G_BUNDLE_ALLOWLIST_ONLY_REQUIRED",
  );
  assert(
    Number(bundle?.readiness?.rolloutPercent) === 0,
    "G_BUNDLE_ROLLOUT_PERCENT_MUST_BE_ZERO",
  );
  assert(
    bundle?.readiness?.actualTrafficEvidenceRequiredFor15_3_4 === true,
    "G_BUNDLE_ACTUAL_TRAFFIC_NEXT_GATE_REQUIRED",
  );

  assert(
    bundle?.legacy15_3EvidenceContract?.satisfiedByThisBundle === false,
    "G_BUNDLE_MUST_NOT_CLAIM_LEGACY_EVIDENCE_SATISFACTION",
  );
  assert(
    bundle?.legacy15_3EvidenceContract?.substitutionForbidden === true,
    "G_BUNDLE_LEGACY_EVIDENCE_SUBSTITUTION_MUST_BE_FORBIDDEN",
  );

  assert(
    bundle?.authorizationBoundary?.runtimeAutoActivationAuthorized === false,
    "G_BUNDLE_RUNTIME_AUTO_ACTIVATION_FORBIDDEN",
  );
  assert(
    bundle?.authorizationBoundary?.actualInternalUserExposureAuthorized ===
      false,
    "G_BUNDLE_ACTUAL_EXPOSURE_NOT_AUTHORIZED_IN_G",
  );
  assert(
    bundle?.authorizationBoundary?.percentageRolloutAuthorized === false,
    "G_BUNDLE_PERCENTAGE_ROLLOUT_FORBIDDEN",
  );
  assert(
    bundle?.authorizationBoundary?.productionPromotionAuthorized === false,
    "G_BUNDLE_PRODUCTION_PROMOTION_FORBIDDEN",
  );

  assert(
    bundle?.guardrails?.providerCallsExecutedByBundleBuilder === 0,
    "G_BUNDLE_PROVIDER_CALLS_NONZERO",
  );
  assert(
    bundle?.guardrails?.actualOperationalTelemetry === false,
    "G_BUNDLE_ACTUAL_TELEMETRY_FORBIDDEN",
  );
  assert(
    bundle?.guardrails?.railwayModified === false,
    "G_BUNDLE_RAILWAY_MUTATION_FORBIDDEN",
  );
  assert(
    bundle?.guardrails?.environmentModified === false,
    "G_BUNDLE_ENV_MUTATION_FORBIDDEN",
  );
  assert(
    bundle?.guardrails?.routeModified === false,
    "G_BUNDLE_ROUTE_MUTATION_FORBIDDEN",
  );

  const copy = JSON.parse(JSON.stringify(bundle));
  const observed = normalizeSha256(copy.bundlePayloadSha256);
  delete copy.bundlePayloadSha256;
  assert(observed === sha256Json(copy), "G_BUNDLE_PAYLOAD_SHA_INVALID");

  const serialized = JSON.stringify(bundle);
  for (const forbidden of [
    '"immutableAccountId"',
    '"accountId"',
    '"tenantId"',
    '"email"',
    '"rawRows"',
    '"providerRawResponse"',
  ]) {
    assert(
      !serialized.includes(forbidden),
      "G_BUNDLE_PRIVACY_BOUNDARY_INVALID",
    );
  }
  return true;
}

module.exports = Object.freeze({
  BUNDLE_VERSION,
  SCOPE,
  DECISION,
  EXPECTED,
  canonicalJson,
  sha256Json,
  sha256File,
  normalizeSha256,
  buildFinalEvaluationEvidenceBundle,
  verifyFinalEvaluationEvidenceBundle,
});
