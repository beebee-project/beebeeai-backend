const SHA256_RE = /^[A-F0-9]{64}$/;

const ROTATION_VERSION =
  "query_candidate_planner_internal_canary_subject_rotation_v1";

const EXPECTED_CANDIDATE_PAYLOAD_SHA256 =
  "928F6A6E0AA8683D63A5A2CB62199FA460EB84494B119EB7E171000843D484EA";

function normalizeSha256(value) {
  const normalized = String(value || "")
    .trim()
    .toUpperCase();
  return SHA256_RE.test(normalized) && !/^0{64}$/.test(normalized)
    ? normalized
    : "";
}

function defaultDeriveSubject(request) {
  const {
    deriveQueryCandidatePlannerInternalCanarySubject,
  } = require("./queryCandidatePlannerInternalCanarySubject");
  return deriveQueryCandidatePlannerInternalCanarySubject(request);
}

function safeSubject(subject = {}) {
  return Object.freeze({
    complete: subject?.complete === true,
    subjectSha256: normalizeSha256(subject?.subjectSha256),
    source: String(subject?.source || "")
      .trim()
      .slice(0, 80),
    rawIdentityIncluded: false,
  });
}

function buildQueryCandidatePlannerInternalCanarySubjectRotation({
  currentAllowlistSha256 = "",
  request = {},
  deriveSubject = defaultDeriveSubject,
} = {}) {
  const current = normalizeSha256(currentAllowlistSha256);
  if (!current) {
    return Object.freeze({
      version: ROTATION_VERSION,
      valid: false,
      decision: "BLOCK",
      reason: "F_1_7_CURRENT_ALLOWLIST_SHA_INVALID",
      failClosed: true,
      rawIdentityIncluded: false,
      environmentModified: false,
      railwayModified: false,
      routeModified: false,
      providerCallsExecutedByRotation: 0,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    });
  }

  let derived;
  try {
    derived = deriveSubject(request);
  } catch {
    return Object.freeze({
      version: ROTATION_VERSION,
      valid: false,
      decision: "BLOCK",
      reason: "F_1_7_SUBJECT_DERIVATION_FAILED",
      failClosed: true,
      currentAllowlistSha256: current,
      rawIdentityIncluded: false,
      environmentModified: false,
      railwayModified: false,
      routeModified: false,
      providerCallsExecutedByRotation: 0,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    });
  }

  const subject = safeSubject(derived);
  if (!subject.complete || !subject.subjectSha256) {
    return Object.freeze({
      version: ROTATION_VERSION,
      valid: false,
      decision: "BLOCK",
      reason: "F_1_7_IMMUTABLE_CANARY_SUBJECT_REQUIRED",
      failClosed: true,
      currentAllowlistSha256: current,
      subject,
      rawIdentityIncluded: false,
      environmentModified: false,
      railwayModified: false,
      routeModified: false,
      providerCallsExecutedByRotation: 0,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    });
  }

  if (subject.subjectSha256 === current) {
    return Object.freeze({
      version: ROTATION_VERSION,
      valid: false,
      decision: "BLOCK",
      reason: "F_1_7_ROTATION_REQUIRES_NEW_SUBJECT",
      failClosed: true,
      currentAllowlistSha256: current,
      proposedAllowlistSha256: subject.subjectSha256,
      subject,
      rawIdentityIncluded: false,
      environmentModified: false,
      railwayModified: false,
      routeModified: false,
      providerCallsExecutedByRotation: 0,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    });
  }

  return Object.freeze({
    version: ROTATION_VERSION,
    valid: true,
    decision: "READY_FOR_LOCAL_APPROVAL_REBINDING",
    reason: "F_1_7_CANARY_SUBJECT_ROTATION_PREPARED",
    failClosed: true,
    currentAllowlistSha256: current,
    proposedAllowlistSha256: subject.subjectSha256,
    subject,
    immutableBindings: Object.freeze({
      candidatePayloadSha256: EXPECTED_CANDIDATE_PAYLOAD_SHA256,
      priorAllowlistSha256: current,
      proposedAllowlistSha256: subject.subjectSha256,
    }),
    approvalRebinding: Object.freeze({
      f14CandidatePreserved: true,
      f15ReceiptReissueRequired: true,
      f16ApprovalBindingReverificationRequired: true,
      f161CompositionReverificationRequired: true,
      runtimePreflightE2ERequired: true,
    }),
    guardrails: Object.freeze({
      localProcessOnlyUntilExplicitPromotionStep: true,
      rawIdentityIncluded: false,
      environmentModified: false,
      railwayModified: false,
      routeModified: false,
      providerCallsExecutedByRotation: 0,
      actualOperationalTelemetry: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    }),
  });
}

module.exports = Object.freeze({
  ROTATION_VERSION,
  EXPECTED_CANDIDATE_PAYLOAD_SHA256,
  normalizeSha256,
  buildQueryCandidatePlannerInternalCanarySubjectRotation,
});
