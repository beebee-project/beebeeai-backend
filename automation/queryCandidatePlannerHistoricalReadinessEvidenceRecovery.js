const crypto = require("crypto");

const CAPSULE_VERSION =
  "query_candidate_planner_historical_live_cache_parity_readiness_capsule_v1";

const SOURCE_VERSION =
  "query_candidate_planner_live_cache_parity_readiness_evidence_v1";

const SOURCE_FACTS = Object.freeze({
  sourceVersion: SOURCE_VERSION,
  model: "gpt-5.6-terra",
  origin: Object.freeze({
    status: "SHADOW_COMPLETED",
    invocationStatus: "CALLED",
    providerCallCount: 1,
    cacheHit: false,
  }),
  replay: Object.freeze({
    status: "SHADOW_COMPLETED",
    invocationStatus: "CACHE_HIT",
    providerCallCount: 0,
    cacheHit: true,
    plannerResolutionSource: "L3_SEMANTIC",
    reentrySource: "L4_REENTRY",
  }),
  parityAudit: Object.freeze({
    valid: true,
    observedProviderCallCount: 1,
    encryptedPersistentFileCount: 3,
    plaintextPersistentFileCount: 0,
    auditSha256:
      "9f231f6354d70a92b0461b930bf876fc0c723bc5c62eac8f26092ac28a54b5b2",
    replayAuditSha256:
      "77380cab79603663ec5cbed085c02e96af3274b0c84ae78343272078b9e77d66",
    originFingerprintSha256:
      "00c30b508086432a71ae8827b177f1c9ee4f09d18f5f8941121a963d2eb24ff0",
    replayFingerprintSha256:
      "00c30b508086432a71ae8827b177f1c9ee4f09d18f5f8941121a963d2eb24ff0",
  }),
  readinessGate: Object.freeze({
    eligible: true,
    decision: "ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW",
    gateSha256:
      "12fe722248ff2403a334ffbe735f97eec7cc52de7099a118c4144fd16d3e7823",
    productionPromotionAllowed: false,
    productionRouteAutoWired: false,
    productionCandidateMergeAllowed: false,
    productionReadyAssignmentAllowed: false,
    manualPromotionReviewRequired: true,
    failClosed: true,
  }),
});

function stableValue(value) {
  if (Array.isArray(value)) return value.map(stableValue);
  if (!value || typeof value !== "object") return value;
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

function buildHistoricalReadinessCapsule() {
  const payload = {
    version: SOURCE_VERSION,
    origin: {
      status: SOURCE_FACTS.origin.status,
      invocationStatus: SOURCE_FACTS.origin.invocationStatus,
      providerCallCount: SOURCE_FACTS.origin.providerCallCount,
      cacheHit: SOURCE_FACTS.origin.cacheHit,
      model: SOURCE_FACTS.model,
    },
    replay: {
      status: SOURCE_FACTS.replay.status,
      invocationStatus: SOURCE_FACTS.replay.invocationStatus,
      providerCallCount: SOURCE_FACTS.replay.providerCallCount,
      cacheHit: SOURCE_FACTS.replay.cacheHit,
      model: SOURCE_FACTS.model,
      cache: {
        plannerResolution: {
          source: SOURCE_FACTS.replay.plannerResolutionSource,
        },
        reentry: {
          source: SOURCE_FACTS.replay.reentrySource,
        },
      },
    },
    parityAudit: {
      valid: SOURCE_FACTS.parityAudit.valid,
      observedProviderCallCount:
        SOURCE_FACTS.parityAudit.observedProviderCallCount,
      persistentFiles: {
        encryptedFileCount:
          SOURCE_FACTS.parityAudit.encryptedPersistentFileCount,
        plaintextFileCount:
          SOURCE_FACTS.parityAudit.plaintextPersistentFileCount,
        encryptedOnly: true,
      },
      auditSha256: SOURCE_FACTS.parityAudit.auditSha256,
      replayAuditSha256: SOURCE_FACTS.parityAudit.replayAuditSha256,
      originFingerprintSha256: SOURCE_FACTS.parityAudit.originFingerprintSha256,
      replayFingerprintSha256: SOURCE_FACTS.parityAudit.replayFingerprintSha256,
    },
    readinessGate: {
      eligible: SOURCE_FACTS.readinessGate.eligible,
      decision: SOURCE_FACTS.readinessGate.decision,
      guardrails: {
        shadowOnlyEvidence: true,
        productionPromotionAllowed:
          SOURCE_FACTS.readinessGate.productionPromotionAllowed,
        productionRouteAutoWired:
          SOURCE_FACTS.readinessGate.productionRouteAutoWired,
        productionCandidateMergeAllowed:
          SOURCE_FACTS.readinessGate.productionCandidateMergeAllowed,
        productionReadyAssignmentAllowed:
          SOURCE_FACTS.readinessGate.productionReadyAssignmentAllowed,
        manualPromotionReviewRequired:
          SOURCE_FACTS.readinessGate.manualPromotionReviewRequired,
        failClosed: SOURCE_FACTS.readinessGate.failClosed,
      },
      gateSha256: SOURCE_FACTS.readinessGate.gateSha256,
    },
    recovery: {
      version: CAPSULE_VERSION,
      source: "ARCHIVED_PATCH_13_3_VERIFIED_READINESS_EVIDENCE",
      purpose: "PATCH_15_3_2_F_1_E_X_COMPATIBILITY_EVALUATION_INPUT",
      historicalEvidenceOnly: true,
      liveProviderCallExecutedByRecovery: false,
      providerCallsExecutedByRecovery: 0,
      actualHistoricalLiveProviderEvidence: true,
      currentOperationalTelemetry: false,
      responseIdIncluded: false,
      tokenUsageValuesIncluded: false,
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      originalSourceVersion: SOURCE_VERSION,
      originalParityAuditSha256: SOURCE_FACTS.parityAudit.auditSha256,
      originalReadinessGateSha256: SOURCE_FACTS.readinessGate.gateSha256,
      originalReplayAuditSha256: SOURCE_FACTS.parityAudit.replayAuditSha256,
      productionPromotionAuthorized: false,
      productionRouteChanged: false,
      privateOutputDoNotCommit: true,
    },
  };

  return Object.freeze({
    ...payload,
    recoveryCapsuleSha256: sha256Json(payload),
  });
}

function validateHistoricalReadinessCapsule(capsule = {}) {
  if (capsule.version !== SOURCE_VERSION) {
    throw new Error("Historical readiness source version mismatch.");
  }
  if (
    capsule.origin?.status !== "SHADOW_COMPLETED" ||
    capsule.origin?.invocationStatus !== "CALLED" ||
    Number(capsule.origin?.providerCallCount) !== 1
  ) {
    throw new Error("Historical origin evidence mismatch.");
  }
  if (
    capsule.replay?.status !== "SHADOW_COMPLETED" ||
    capsule.replay?.invocationStatus !== "CACHE_HIT" ||
    Number(capsule.replay?.providerCallCount) !== 0
  ) {
    throw new Error("Historical replay evidence mismatch.");
  }
  if (
    capsule.replay?.cache?.plannerResolution?.source !== "L3_SEMANTIC" ||
    capsule.replay?.cache?.reentry?.source !== "L4_REENTRY"
  ) {
    throw new Error("Historical cache source evidence mismatch.");
  }
  if (
    capsule.parityAudit?.valid !== true ||
    Number(capsule.parityAudit?.observedProviderCallCount) !== 1 ||
    Number(capsule.parityAudit?.persistentFiles?.encryptedFileCount) !== 3 ||
    Number(capsule.parityAudit?.persistentFiles?.plaintextFileCount) !== 0
  ) {
    throw new Error("Historical parity evidence mismatch.");
  }
  if (
    capsule.parityAudit?.auditSha256 !== SOURCE_FACTS.parityAudit.auditSha256 ||
    capsule.parityAudit?.replayAuditSha256 !==
      SOURCE_FACTS.parityAudit.replayAuditSha256
  ) {
    throw new Error("Historical parity audit SHA mismatch.");
  }
  if (
    capsule.readinessGate?.eligible !== true ||
    capsule.readinessGate?.gateSha256 !== SOURCE_FACTS.readinessGate.gateSha256
  ) {
    throw new Error("Historical readiness gate evidence mismatch.");
  }
  if (
    capsule.readinessGate?.guardrails?.productionPromotionAllowed !== false ||
    capsule.readinessGate?.guardrails?.productionRouteAutoWired !== false
  ) {
    throw new Error("Historical production isolation mismatch.");
  }
  if (
    capsule.recovery?.historicalEvidenceOnly !== true ||
    capsule.recovery?.providerCallsExecutedByRecovery !== 0 ||
    capsule.recovery?.productionPromotionAuthorized !== false
  ) {
    throw new Error("Historical recovery guardrail mismatch.");
  }

  const copy = JSON.parse(JSON.stringify(capsule));
  delete copy.recoveryCapsuleSha256;
  if (sha256Json(copy) !== capsule.recoveryCapsuleSha256) {
    throw new Error("Historical recovery capsule SHA mismatch.");
  }
  return true;
}

module.exports = Object.freeze({
  CAPSULE_VERSION,
  SOURCE_VERSION,
  SOURCE_FACTS,
  sha256Json,
  buildHistoricalReadinessCapsule,
  validateHistoricalReadinessCapsule,
});
