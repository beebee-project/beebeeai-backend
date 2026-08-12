const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const RECEIPT_VERSION =
  "query_candidate_planner_internal_canary_manual_approval_receipt_v1";

const CANDIDATE_VERSION =
  "query_candidate_planner_internal_canary_evidence_candidate_v1";

const SCOPE = "INTERNAL_ALLOWLIST_CANARY_ONLY";

const APPROVAL_DECISION = "INTERNAL_ALLOWLIST_CANARY_MANUAL_APPROVAL_GRANTED";

const EXPECTED_CANDIDATE_PAYLOAD_SHA256 =
  "928F6A6E0AA8683D63A5A2CB62199FA460EB84494B119EB7E171000843D484EA";

const ALLOWLIST_ENV_NAME = "QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256";

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

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(canonicalJson(value))
    .digest("hex")
    .toUpperCase();
}

function sha256Buffer(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex").toUpperCase();
}

function sha256File(file) {
  return sha256Buffer(fs.readFileSync(path.resolve(file)));
}

function normalizeSha256(value, code = "SHA256_INVALID") {
  const normalized = String(value || "")
    .trim()
    .toUpperCase();
  assert(
    /^[A-F0-9]{64}$/.test(normalized),
    code,
    `Expected 64-character SHA-256 but found: ${value}`,
  );
  assert(!/^0{64}$/.test(normalized), code, "All-zero SHA-256 is forbidden.");
  return normalized;
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(path.resolve(file), "utf8"));
}

function validateCandidate(
  candidate = {},
  expectedPayloadSha256 = EXPECTED_CANDIDATE_PAYLOAD_SHA256,
) {
  const expected = normalizeSha256(
    expectedPayloadSha256,
    "EXPECTED_CANDIDATE_SHA_INVALID",
  );

  assert(candidate.version === CANDIDATE_VERSION, "CANDIDATE_VERSION_INVALID");
  assert(candidate.scope === SCOPE, "CANDIDATE_SCOPE_INVALID");

  assert(
    candidate.evaluation?.operationalDecision === "EVALUATION_PASS",
    "CANDIDATE_EVALUATION_NOT_PASS",
  );
  assert(
    candidate.evaluation?.assessmentDecision ===
      "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS",
    "CANDIDATE_ASSESSMENT_NOT_PASS",
  );
  assert(
    Number(candidate.evaluation?.failedCheckCount) === 0 &&
      Number(candidate.evaluation?.absoluteCostFailureCount) === 0 &&
      candidate.evaluation?.cacheCostAvoidancePassed === true,
    "CANDIDATE_EVALUATION_RESULT_INVALID",
  );

  assert(
    candidate.eligibility?.decision ===
      "ELIGIBLE_FOR_INTERNAL_ALLOWLIST_CANARY_REVIEW" &&
      candidate.eligibility?.internalCanaryReviewEligible === true &&
      candidate.eligibility?.manualOperatorApprovalRequired === true,
    "CANDIDATE_REVIEW_ELIGIBILITY_INVALID",
  );

  assert(
    candidate.eligibility?.internalCanaryAuthorized === false &&
      candidate.eligibility?.percentageRolloutAuthorized === false &&
      candidate.eligibility?.productionPromotionAuthorized === false &&
      candidate.eligibility?.productionMergeAuthorized === false,
    "CANDIDATE_PREAPPROVAL_BOUNDARY_INVALID",
  );

  assert(
    candidate.methodology?.actualOperationalTelemetry === false &&
      candidate.methodology?.internalCanaryEvidence === false &&
      candidate.methodology?.productionPromotionEvidence === false,
    "CANDIDATE_EVIDENCE_BOUNDARY_INVALID",
  );

  assert(
    candidate.integrity?.evaluator?.worktreeEqualsHead === true,
    "CANDIDATE_EVALUATOR_DRIFT",
  );

  const observedPayloadSha = normalizeSha256(
    candidate.candidatePayloadSha256,
    "CANDIDATE_PAYLOAD_SHA_INVALID",
  );

  assert(observedPayloadSha === expected, "CANDIDATE_PAYLOAD_SHA_DRIFT");

  const copy = JSON.parse(JSON.stringify(candidate));
  delete copy.candidatePayloadSha256;

  assert(
    sha256Json(copy) === observedPayloadSha,
    "CANDIDATE_PAYLOAD_INTEGRITY_INVALID",
  );

  return true;
}

function buildManualApprovalReceipt({
  candidateBundleFile,
  allowlistSha256,
  approve,
} = {}) {
  assert(
    typeof candidateBundleFile === "string" &&
      candidateBundleFile.trim().length > 0,
    "CANDIDATE_BUNDLE_PATH_REQUIRED",
  );

  const candidatePath = path.resolve(candidateBundleFile);
  assert(fs.existsSync(candidatePath), "CANDIDATE_BUNDLE_MISSING");

  assert(
    String(approve || "").toLowerCase() === "true",
    "EXPLICIT_MANUAL_APPROVAL_REQUIRED",
    "--approve true is required to issue the manual approval receipt.",
  );

  const normalizedAllowlistSha256 = normalizeSha256(
    allowlistSha256,
    "ALLOWLIST_SHA256_INVALID",
  );

  const candidate = readJson(candidatePath);
  validateCandidate(candidate);

  const payload = {
    version: RECEIPT_VERSION,
    scope: SCOPE,
    sourcePatch: "15.3.2-F.1.5",
    decision: APPROVAL_DECISION,

    manualApproval: {
      approvedByOperator: true,
      evidenceBundleReviewed: true,
      allowlistHashReviewed: true,
      explicitApprovalFlagRequired: true,
      approvalIsRuntimeActivation: false,
    },

    immutableBindings: {
      candidatePayloadSha256: candidate.candidatePayloadSha256,
      candidateBundleFileSha256: sha256File(candidatePath),
      allowlistSha256: normalizedAllowlistSha256,
      allowlistEnvironmentVariableName: ALLOWLIST_ENV_NAME,
    },

    evidenceSnapshot: {
      operationalDecision: candidate.evaluation.operationalDecision,
      assessmentDecision: candidate.evaluation.assessmentDecision,
      failedCheckCount: Number(candidate.evaluation.failedCheckCount),
      absoluteCostFailureCount: Number(
        candidate.evaluation.absoluteCostFailureCount,
      ),
      cacheCostAvoidancePassed:
        candidate.evaluation.cacheCostAvoidancePassed === true,
      cacheCostAvoidanceActual: Number(
        candidate.evaluation.cacheCostAvoidanceActual,
      ),
      cacheCostAvoidanceThreshold: Number(
        candidate.evaluation.cacheCostAvoidanceThreshold,
      ),
      evaluatorSha256: candidate.integrity.evaluator.worktreeSha256,
      evaluatorWorktreeEqualsHead:
        candidate.integrity.evaluator.worktreeEqualsHead === true,
      finalBaselineSha256: candidate.integrity.finalBaselineSha256,
      actualOperationalTelemetry:
        candidate.methodology.actualOperationalTelemetry === true,
    },

    thresholdContract: {
      averageCostMicrousdMax: Number(
        candidate.thresholds.averageCostMicrousdMax,
      ),
      providerCallAverageCostMicrousdMax: Number(
        candidate.thresholds.providerCallAverageCostMicrousdMax,
      ),
      monthlyProjectedCostMicrousdMax: Number(
        candidate.thresholds.monthlyProjectedCostMicrousdMax,
      ),
      cacheCostAvoidanceRateMin: Number(
        candidate.thresholds.cacheCostAvoidanceRateMin,
      ),
      providerCallRateMax: Number(candidate.thresholds.providerCallRateMax),
      warmAverageCostMicrousdMax: Number(
        candidate.thresholds.warmAverageCostMicrousdMax,
      ),
    },

    authorizationBoundary: {
      internalCanaryApprovalGranted: true,
      runtimeGateBindingApplied: false,
      runtimeCanaryAuthorized: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
      productionMergeAuthorized: false,
    },

    nextGateRequirements: {
      candidatePayloadSha256MustMatch: true,
      approvalReceiptPayloadSha256MustMatch: true,
      allowlistSha256MustMatch: true,
      killSwitchMustPermit: true,
      featureFlagMustPermit: true,
      runtimeGateIntegrationRequired: true,
    },

    guardrails: {
      noGateMutation: true,
      noEnvironmentMutation: true,
      noRouteMutation: true,
      noFeatureFlagMutation: true,
      noKillSwitchMutation: true,
      noAllowlistMutation: true,
      providerCallsExecutedByReceiptBuilder: 0,
      privateOutputDoNotCommit: true,
    },

    sanitization: {
      immutableAccountIdsIncluded: false,
      allowlistSubjectsIncluded: false,
      environmentValuesIncluded: false,
      rawEvaluationRowsIncluded: false,
      rawTokenUsageIncluded: false,
    },
  };

  const approvalReceiptPayloadSha256 = sha256Json(payload);

  const receipt = Object.freeze({
    ...payload,
    approvalReceiptPayloadSha256,
  });

  verifyManualApprovalReceipt(receipt);
  return receipt;
}

function verifyManualApprovalReceipt(receipt = {}) {
  assert(receipt.version === RECEIPT_VERSION, "RECEIPT_VERSION_INVALID");
  assert(receipt.scope === SCOPE, "RECEIPT_SCOPE_INVALID");
  assert(receipt.decision === APPROVAL_DECISION, "RECEIPT_DECISION_INVALID");

  const binding = receipt.immutableBindings || {};

  assert(
    normalizeSha256(
      binding.candidatePayloadSha256,
      "RECEIPT_CANDIDATE_SHA_INVALID",
    ) === EXPECTED_CANDIDATE_PAYLOAD_SHA256,
    "RECEIPT_CANDIDATE_SHA_DRIFT",
  );

  normalizeSha256(
    binding.candidateBundleFileSha256,
    "RECEIPT_CANDIDATE_FILE_SHA_INVALID",
  );
  normalizeSha256(binding.allowlistSha256, "RECEIPT_ALLOWLIST_SHA_INVALID");

  assert(
    binding.allowlistEnvironmentVariableName === ALLOWLIST_ENV_NAME,
    "RECEIPT_ALLOWLIST_ENV_NAME_INVALID",
  );

  assert(
    receipt.manualApproval?.approvedByOperator === true &&
      receipt.manualApproval?.evidenceBundleReviewed === true &&
      receipt.manualApproval?.allowlistHashReviewed === true &&
      receipt.manualApproval?.approvalIsRuntimeActivation === false,
    "RECEIPT_MANUAL_APPROVAL_INVALID",
  );

  assert(
    receipt.evidenceSnapshot?.operationalDecision === "EVALUATION_PASS" &&
      receipt.evidenceSnapshot?.assessmentDecision ===
        "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS" &&
      Number(receipt.evidenceSnapshot?.failedCheckCount) === 0 &&
      Number(receipt.evidenceSnapshot?.absoluteCostFailureCount) === 0 &&
      receipt.evidenceSnapshot?.cacheCostAvoidancePassed === true,
    "RECEIPT_EVIDENCE_SNAPSHOT_INVALID",
  );

  assert(
    receipt.evidenceSnapshot?.evaluatorWorktreeEqualsHead === true &&
      receipt.evidenceSnapshot?.actualOperationalTelemetry === false,
    "RECEIPT_EVIDENCE_BOUNDARY_INVALID",
  );

  assert(
    receipt.authorizationBoundary?.internalCanaryApprovalGranted === true &&
      receipt.authorizationBoundary?.runtimeGateBindingApplied === false &&
      receipt.authorizationBoundary?.runtimeCanaryAuthorized === false &&
      receipt.authorizationBoundary?.percentageRolloutAuthorized === false &&
      receipt.authorizationBoundary?.productionPromotionAuthorized === false &&
      receipt.authorizationBoundary?.productionMergeAuthorized === false,
    "RECEIPT_AUTHORIZATION_BOUNDARY_INVALID",
  );

  assert(
    receipt.guardrails?.noGateMutation === true &&
      receipt.guardrails?.noEnvironmentMutation === true &&
      receipt.guardrails?.noRouteMutation === true &&
      receipt.guardrails?.noFeatureFlagMutation === true &&
      receipt.guardrails?.noKillSwitchMutation === true &&
      receipt.guardrails?.noAllowlistMutation === true &&
      Number(receipt.guardrails?.providerCallsExecutedByReceiptBuilder) === 0,
    "RECEIPT_GUARDRAIL_INVALID",
  );

  const copy = JSON.parse(JSON.stringify(receipt));
  const observed = copy.approvalReceiptPayloadSha256;
  delete copy.approvalReceiptPayloadSha256;

  assert(observed === sha256Json(copy), "RECEIPT_PAYLOAD_SHA_INVALID");

  const serialized = JSON.stringify(receipt);
  for (const forbidden of [
    '"immutableAccountId"',
    '"allowlistSubjects"',
    '"responseId"',
    '"inputTokens"',
    '"outputTokens"',
  ]) {
    assert(
      !serialized.includes(forbidden),
      "RECEIPT_SANITIZATION_FAILURE",
      forbidden,
    );
  }

  return true;
}

module.exports = Object.freeze({
  RECEIPT_VERSION,
  CANDIDATE_VERSION,
  SCOPE,
  APPROVAL_DECISION,
  EXPECTED_CANDIDATE_PAYLOAD_SHA256,
  ALLOWLIST_ENV_NAME,
  canonicalJson,
  sha256Json,
  sha256File,
  normalizeSha256,
  validateCandidate,
  buildManualApprovalReceipt,
  verifyManualApprovalReceipt,
});
