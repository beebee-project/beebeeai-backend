const path = require("path");
const { normalizeText, sha256 } = require("./queryCandidateObservation");
const {
  compareReplaySafeShadowResolutions,
} = require("./queryCandidatePlannerCacheOperationalControls");
const {
  CACHE_READ_SOURCE,
} = require("./queryCandidatePlannerHierarchicalEncryptedCache");

const QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_AUDIT_VERSION =
  "query_candidate_planner_live_cache_parity_audit_v1";
const QUERY_CANDIDATE_PLANNER_PRODUCTION_READINESS_GATE_VERSION =
  "query_candidate_planner_production_readiness_gate_v1";
const QUERY_CANDIDATE_PLANNER_PRODUCTION_READINESS_POLICY_VERSION =
  "candidate_planner_production_readiness_policy_v1";

const READINESS_DECISION = Object.freeze({
  ELIGIBLE: "ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW",
  BLOCKED: "NOT_ELIGIBLE",
});

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function normalizedCount(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) && number >= 0 ? Math.floor(number) : 0;
}

function shadowCounts(resolution = {}) {
  return {
    accepted: normalizedCount(resolution.counts?.accepted),
    resolved: normalizedCount(resolution.counts?.resolved),
    ready: normalizedCount(resolution.counts?.ready),
    ranked: normalizedCount(resolution.counts?.ranked),
  };
}

function countsAligned(counts = {}) {
  return (
    counts.accepted > 0 &&
    counts.accepted === counts.resolved &&
    counts.accepted === counts.ready &&
    counts.accepted === counts.ranked
  );
}

function productionIsolation(resolution = {}) {
  return {
    productionCandidateMerge:
      resolution.integrity?.productionCandidateMerge === true,
    productionReadyAssignment:
      resolution.integrity?.productionReadyAssignment === true,
    productionRouteChanged:
      resolution.integrity?.productionRouteChanged === true,
  };
}

function privacyBoundaryValid(resolution = {}) {
  return (
    resolution.privacy?.rawRowsSent === false &&
    resolution.privacy?.sampleValuesSent === false &&
    resolution.privacy?.originalFileSent === false &&
    resolution.privacy?.fileNameSent === false &&
    resolution.cache?.encryptedPersistentOnly === true &&
    resolution.cache?.plaintextPersistenceAllowed === false
  );
}

function inspectPersistentCacheFiles(persistentFiles = []) {
  const files = asArray(persistentFiles)
    .map((value) => normalizeText(value || ""))
    .filter(Boolean);
  const encryptedFileCount = files.filter(
    (file) => path.extname(file).toLowerCase() === ".enc",
  ).length;
  const plaintextFileCount = files.length - encryptedFileCount;
  return {
    totalFileCount: files.length,
    encryptedFileCount,
    plaintextFileCount,
    encryptedOnly: files.length > 0 && plaintextFileCount === 0,
  };
}

function buildLiveProviderCacheHitParityAudit({
  origin,
  replay,
  observedProviderCallCount,
  persistentFiles = [],
} = {}) {
  const errors = [];
  const originInvocation = normalizeText(
    origin?.plannerResolution?.invocation?.status || "",
  );
  const replayInvocation = normalizeText(
    replay?.plannerResolution?.invocation?.status || "",
  );
  const originProviderCallCount = normalizedCount(
    origin?.plannerResolution?.invocation?.providerCallCount,
  );
  const replayProviderCallCount = normalizedCount(
    replay?.plannerResolution?.invocation?.providerCallCount,
  );
  const observedCalls = normalizedCount(observedProviderCallCount);
  const originCounts = shadowCounts(origin);
  const replayCounts = shadowCounts(replay);
  const originFailureCode = normalizeText(
    origin?.plannerResolution?.invocation?.failureCode || "",
  );
  const originResponseIdPresent = Boolean(
    normalizeText(origin?.plannerResolution?.invocation?.responseId || ""),
  );
  const originTotalTokens = normalizedCount(
    origin?.plannerResolution?.usage?.totalTokens,
  );
  const replayAudit = compareReplaySafeShadowResolutions({ origin, replay });
  const fileAudit = inspectPersistentCacheFiles(persistentFiles);
  const originIsolation = productionIsolation(origin);
  const replayIsolation = productionIsolation(replay);
  const productionIsolationVerified = [originIsolation, replayIsolation].every(
    (isolation) =>
      isolation.productionCandidateMerge === false &&
      isolation.productionReadyAssignment === false &&
      isolation.productionRouteChanged === false,
  );
  const persistentPlannerSource = normalizeText(
    replay?.cache?.plannerResolution?.source || "",
  );
  const persistentReentrySource = normalizeText(
    replay?.cache?.reentry?.source || "",
  );
  const liveProviderVerified =
    normalizeText(origin?.status || "") === "SHADOW_COMPLETED" &&
    originInvocation === "CALLED" &&
    originProviderCallCount === 1 &&
    observedCalls === 1 &&
    originResponseIdPresent &&
    originTotalTokens > 0 &&
    !originFailureCode;
  const persistentCacheHitVerified =
    normalizeText(replay?.status || "") === "SHADOW_COMPLETED" &&
    replayInvocation === "CACHE_HIT" &&
    replayProviderCallCount === 0 &&
    replay?.cache?.plannerProvider?.cacheHit === true &&
    replay?.cache?.plannerResolution?.hit === true &&
    persistentPlannerSource === CACHE_READ_SOURCE.L3_SEMANTIC &&
    replay?.cache?.reentry?.hit === true &&
    persistentReentrySource === CACHE_READ_SOURCE.L4_REENTRY;
  const countsParityVerified =
    countsAligned(originCounts) &&
    countsAligned(replayCounts) &&
    originCounts.accepted === replayCounts.accepted &&
    originCounts.resolved === replayCounts.resolved &&
    originCounts.ready === replayCounts.ready &&
    originCounts.ranked === replayCounts.ranked;
  const replaySafeVerified = replayAudit.valid === true;
  const encryptedPersistenceVerified =
    fileAudit.encryptedOnly === true && fileAudit.encryptedFileCount >= 3;
  const privacyBoundaryVerified =
    privacyBoundaryValid(origin) && privacyBoundaryValid(replay);

  if (normalizeText(origin?.status || "") !== "SHADOW_COMPLETED") {
    errors.push({ code: "ORIGIN_NOT_SHADOW_COMPLETED" });
  }
  if (originInvocation !== "CALLED") {
    errors.push({ code: "ORIGIN_NOT_LIVE_PROVIDER_CALLED" });
  }
  if (originProviderCallCount !== 1) {
    errors.push({ code: "ORIGIN_PROVIDER_CALL_COUNT_INVALID" });
  }
  if (observedCalls !== 1) {
    errors.push({ code: "OBSERVED_PROVIDER_CALL_COUNT_INVALID" });
  }
  if (!originResponseIdPresent) {
    errors.push({ code: "ORIGIN_RESPONSE_ID_MISSING" });
  }
  if (originTotalTokens <= 0) {
    errors.push({ code: "ORIGIN_TOKEN_USAGE_MISSING" });
  }
  if (originFailureCode) {
    errors.push({ code: "ORIGIN_FAILURE_CODE_PRESENT" });
  }
  if (!countsAligned(originCounts)) {
    errors.push({ code: "ORIGIN_COUNTS_NOT_ALIGNED" });
  }
  if (normalizeText(replay?.status || "") !== "SHADOW_COMPLETED") {
    errors.push({ code: "REPLAY_NOT_SHADOW_COMPLETED" });
  }
  if (replayInvocation !== "CACHE_HIT") {
    errors.push({ code: "REPLAY_NOT_CACHE_HIT" });
  }
  if (replayProviderCallCount !== 0) {
    errors.push({ code: "REPLAY_PROVIDER_CALL_OCCURRED" });
  }
  if (replay?.cache?.plannerProvider?.cacheHit !== true) {
    errors.push({ code: "REPLAY_PLANNER_PROVIDER_CACHE_MISS" });
  }
  if (
    replay?.cache?.plannerResolution?.hit !== true ||
    persistentPlannerSource !== CACHE_READ_SOURCE.L3_SEMANTIC
  ) {
    errors.push({ code: "REPLAY_PLANNER_RESOLUTION_NOT_PERSISTENT_HIT" });
  }
  if (
    replay?.cache?.reentry?.hit !== true ||
    persistentReentrySource !== CACHE_READ_SOURCE.L4_REENTRY
  ) {
    errors.push({ code: "REPLAY_REENTRY_NOT_PERSISTENT_HIT" });
  }
  if (!countsParityVerified) {
    errors.push({ code: "LIVE_CACHE_COUNTS_PARITY_MISMATCH" });
  }
  if (!replaySafeVerified) {
    errors.push({
      code: "REPLAY_SAFE_AUDIT_FAILED",
      replayAuditSha256: normalizeText(replayAudit.auditSha256 || ""),
    });
  }
  if (fileAudit.encryptedFileCount < 3) {
    errors.push({ code: "ENCRYPTED_CACHE_ARTIFACT_COUNT_INSUFFICIENT" });
  }
  if (fileAudit.plaintextFileCount > 0) {
    errors.push({ code: "PLAINTEXT_CACHE_ARTIFACT_FOUND" });
  }
  if (!privacyBoundaryVerified) {
    errors.push({ code: "PRIVACY_BOUNDARY_VIOLATION" });
  }
  if (!productionIsolationVerified) {
    errors.push({ code: "PRODUCTION_ISOLATION_VIOLATION" });
  }

  const document = {
    version: QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_AUDIT_VERSION,
    valid: errors.length === 0,
    errorCount: errors.length,
    errors,
    originInvocation,
    replayInvocation,
    originProviderCallCount,
    replayProviderCallCount,
    observedProviderCallCount: observedCalls,
    originResponseIdPresent,
    originTokenUsagePositive: originTotalTokens > 0,
    originFailureCodePresent: Boolean(originFailureCode),
    originCounts,
    replayCounts,
    cacheSources: {
      plannerResolution: persistentPlannerSource,
      reentry: persistentReentrySource,
    },
    persistentFiles: fileAudit,
    checks: {
      liveProviderVerified,
      persistentCacheHitVerified,
      countsParityVerified,
      replaySafeVerified,
      encryptedPersistenceVerified,
      privacyBoundaryVerified,
      productionIsolationVerified,
    },
    replayAuditSha256: normalizeText(replayAudit.auditSha256 || ""),
    originFingerprintSha256: normalizeText(
      replayAudit.originFingerprintSha256 || "",
    ),
    replayFingerprintSha256: normalizeText(
      replayAudit.replayFingerprintSha256 || "",
    ),
    productionIsolation: replayIsolation,
    responseIdIncluded: false,
    tokenUsageValuesIncluded: false,
    plaintextIdentifiersIncluded: false,
  };
  document.auditSha256 = sha256(document);
  return Object.freeze(document);
}

function evaluateCandidatePlannerProductionReadiness({ parityAudit } = {}) {
  const blockingReasons = [];
  const audit =
    parityAudit && typeof parityAudit === "object" ? parityAudit : {};
  const checks = {
    parityAuditVersionValid:
      audit.version === QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_AUDIT_VERSION,
    parityAuditValid: audit.valid === true,
    liveProviderVerified: audit.checks?.liveProviderVerified === true,
    persistentCacheHitVerified:
      audit.checks?.persistentCacheHitVerified === true,
    countsParityVerified: audit.checks?.countsParityVerified === true,
    replaySafeVerified: audit.checks?.replaySafeVerified === true,
    encryptedPersistenceVerified:
      audit.checks?.encryptedPersistenceVerified === true,
    privacyBoundaryVerified: audit.checks?.privacyBoundaryVerified === true,
    productionIsolationVerified:
      audit.checks?.productionIsolationVerified === true,
  };
  for (const [name, passed] of Object.entries(checks)) {
    if (!passed) blockingReasons.push(name);
  }
  const eligible = blockingReasons.length === 0;
  const document = {
    version: QUERY_CANDIDATE_PLANNER_PRODUCTION_READINESS_GATE_VERSION,
    policyVersion: QUERY_CANDIDATE_PLANNER_PRODUCTION_READINESS_POLICY_VERSION,
    eligible,
    decision: eligible
      ? READINESS_DECISION.ELIGIBLE
      : READINESS_DECISION.BLOCKED,
    blockingReasonCount: blockingReasons.length,
    blockingReasons,
    checks,
    evidence: {
      parityAuditSha256: normalizeText(audit.auditSha256 || ""),
      replayAuditSha256: normalizeText(audit.replayAuditSha256 || ""),
      originFingerprintSha256: normalizeText(
        audit.originFingerprintSha256 || "",
      ),
      replayFingerprintSha256: normalizeText(
        audit.replayFingerprintSha256 || "",
      ),
      encryptedPersistentFileCount: normalizedCount(
        audit.persistentFiles?.encryptedFileCount,
      ),
    },
    guardrails: {
      shadowOnlyEvidence: true,
      productionPromotionAllowed: false,
      productionRouteAutoWired: false,
      productionCandidateMergeAllowed: false,
      productionReadyAssignmentAllowed: false,
      manualPromotionReviewRequired: true,
      failClosed: true,
    },
    privacy: {
      responseIdIncluded: false,
      tokenUsageValuesIncluded: false,
      originalFileNameIncluded: false,
      tenantIdIncluded: false,
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
    },
  };
  document.gateSha256 = sha256(document);
  return Object.freeze(document);
}

module.exports = {
  QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_AUDIT_VERSION,
  QUERY_CANDIDATE_PLANNER_PRODUCTION_READINESS_GATE_VERSION,
  QUERY_CANDIDATE_PLANNER_PRODUCTION_READINESS_POLICY_VERSION,
  READINESS_DECISION,
  inspectPersistentCacheFiles,
  buildLiveProviderCacheHitParityAudit,
  evaluateCandidatePlannerProductionReadiness,
};
