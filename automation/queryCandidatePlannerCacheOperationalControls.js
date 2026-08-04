"use strict";

const {
  normalizeText,
  sha256,
  stableStringify,
} = require("./queryCandidateObservation");
const {
  CACHE_INVALIDATION_REASONS,
} = require("./queryCandidatePlannerHierarchicalEncryptedCache");

const QUERY_CANDIDATE_PLANNER_CACHE_REPLAY_AUDIT_VERSION =
  "query_candidate_planner_cache_replay_audit_v1";
const QUERY_CANDIDATE_PLANNER_UPLOAD_INVALIDATION_VERSION =
  "query_candidate_planner_upload_invalidation_v1";

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function normalizeSha(value, fieldName, { required = false } = {}) {
  const normalized = normalizeText(value || "").toLowerCase();
  if (!normalized && !required) return "";
  if (!/^[a-f0-9]{64}$/.test(normalized)) {
    throw new Error(`${fieldName}는 SHA-256 64자리 hex여야 합니다.`);
  }
  return normalized;
}

function acceptedPlannerItemSha256s(resolution = {}) {
  return asArray(resolution.plannerResolution?.proposals)
    .filter((proposal) => proposal.disposition === "ACCEPTED_FOR_REVALIDATION")
    .map((proposal) => normalizeSha(proposal.plannerItemSha256, "plannerItemSha256", {
      required: true,
    }))
    .sort();
}

function buildReplaySafeShadowFingerprint(resolution = {}) {
  const fingerprint = {
    status: normalizeText(resolution.status || ""),
    planner: {
      decision: normalizeText(resolution.plannerResolution?.decision || ""),
      acceptedPlannerItemSha256s: acceptedPlannerItemSha256s(resolution),
      modelOutputSha256: normalizeSha(
        resolution.plannerResolution?.modelOutputSha256,
        "modelOutputSha256",
      ),
      plannerInputSha256: normalizeSha(
        resolution.plannerResolution?.source?.inputSha256,
        "plannerInputSha256",
      ),
    },
    reentry: {
      bundleSha256: normalizeSha(resolution.reentry?.bundleSha256, "bundleSha256"),
      candidateResolutionSha256: normalizeSha(
        resolution.reentry?.candidateResolutionSha256,
        "candidateResolutionSha256",
      ),
      candidateFamilyResolutionSha256: normalizeSha(
        resolution.reentry?.candidateFamilyResolutionSha256,
        "candidateFamilyResolutionSha256",
      ),
      candidateFeasibilityResolutionSha256: normalizeSha(
        resolution.reentry?.candidateFeasibilityResolutionSha256,
        "candidateFeasibilityResolutionSha256",
      ),
      candidateRankingResolutionSha256: normalizeSha(
        resolution.reentry?.candidateRankingResolutionSha256,
        "candidateRankingResolutionSha256",
      ),
      items: asArray(resolution.reentry?.items)
        .map((item) => ({
          candidateId: normalizeText(item.candidateId || ""),
          resolverResult: normalizeText(item.resolverResult || ""),
          feasibilityStatus: normalizeText(item.feasibilityStatus || ""),
          rankingDisposition: normalizeText(item.rankingDisposition || ""),
          shadowRank: Number(item.shadowRank || 0),
          productionCandidateMerged: item.productionCandidateMerged === true,
          productionReadyAssigned: item.productionReadyAssigned === true,
        }))
        .sort((left, right) => {
          return left.candidateId.localeCompare(right.candidateId);
        }),
    },
    counts: {
      accepted: Number(resolution.counts?.accepted || 0),
      resolved: Number(resolution.counts?.resolved || 0),
      ready: Number(resolution.counts?.ready || 0),
      ranked: Number(resolution.counts?.ranked || 0),
    },
    productionIsolation: {
      productionCandidateMerge:
        resolution.integrity?.productionCandidateMerge === true,
      productionReadyAssignment:
        resolution.integrity?.productionReadyAssignment === true,
      productionRouteChanged:
        resolution.integrity?.productionRouteChanged === true,
    },
  };
  return Object.freeze({
    ...fingerprint,
    fingerprintSha256: sha256(fingerprint),
  });
}

function compareReplaySafeShadowResolutions({ origin, replay } = {}) {
  const originFingerprint = buildReplaySafeShadowFingerprint(origin);
  const replayFingerprint = buildReplaySafeShadowFingerprint(replay);
  const errors = [];
  if (normalizeText(origin?.status || "") !== "SHADOW_COMPLETED") {
    errors.push({ code: "ORIGIN_NOT_COMPLETED" });
  }
  if (normalizeText(replay?.status || "") !== "SHADOW_COMPLETED") {
    errors.push({ code: "REPLAY_NOT_COMPLETED" });
  }
  if (normalizeText(replay?.plannerResolution?.invocation?.status || "") !== "CACHE_HIT") {
    errors.push({ code: "REPLAY_NOT_CACHE_HIT" });
  }
  if (Number(replay?.plannerResolution?.invocation?.providerCallCount || 0) !== 0) {
    errors.push({ code: "REPLAY_PROVIDER_CALL_OCCURRED" });
  }
  if (originFingerprint.fingerprintSha256 !== replayFingerprint.fingerprintSha256) {
    errors.push({
      code: "REPLAY_FINGERPRINT_MISMATCH",
      originFingerprintSha256: originFingerprint.fingerprintSha256,
      replayFingerprintSha256: replayFingerprint.fingerprintSha256,
    });
  }
  if (
    replayFingerprint.productionIsolation.productionCandidateMerge ||
    replayFingerprint.productionIsolation.productionReadyAssignment ||
    replayFingerprint.productionIsolation.productionRouteChanged
  ) {
    errors.push({ code: "PRODUCTION_ISOLATION_VIOLATION" });
  }
  const document = {
    version: QUERY_CANDIDATE_PLANNER_CACHE_REPLAY_AUDIT_VERSION,
    valid: errors.length === 0,
    errorCount: errors.length,
    errors,
    originInvocation: normalizeText(
      origin?.plannerResolution?.invocation?.status || "",
    ),
    replayInvocation: normalizeText(
      replay?.plannerResolution?.invocation?.status || "",
    ),
    replayProviderCallCount: Number(
      replay?.plannerResolution?.invocation?.providerCallCount || 0,
    ),
    originFingerprintSha256: originFingerprint.fingerprintSha256,
    replayFingerprintSha256: replayFingerprint.fingerprintSha256,
    productionIsolation: replayFingerprint.productionIsolation,
  };
  document.auditSha256 = sha256(document);
  return Object.freeze(document);
}

function buildUploadInvalidationTags({
  uploadFingerprintSha256,
  queryJsonSha256,
} = {}) {
  const tags = {
    uploadFingerprintSha256: normalizeSha(
      uploadFingerprintSha256,
      "uploadFingerprintSha256",
    ),
    queryJsonSha256: normalizeSha(queryJsonSha256, "queryJsonSha256"),
  };
  if (!tags.uploadFingerprintSha256 && !tags.queryJsonSha256) {
    throw new Error("업로드 무효화에는 uploadFingerprintSha256 또는 queryJsonSha256가 필요합니다.");
  }
  return Object.freeze(tags);
}

async function invalidateCandidatePlannerUploadCache({
  hierarchicalCache,
  tenantId,
  cacheSecret,
  uploadFingerprintSha256,
  queryJsonSha256,
} = {}) {
  if (!hierarchicalCache || typeof hierarchicalCache.invalidateByTags !== "function") {
    throw new TypeError("hierarchicalCache.invalidateByTags 함수가 필요합니다.");
  }
  const tags = buildUploadInvalidationTags({
    uploadFingerprintSha256,
    queryJsonSha256,
  });
  const result = await hierarchicalCache.invalidateByTags({
    tenantId,
    cacheSecret,
    tags,
    reason: CACHE_INVALIDATION_REASONS.UPLOAD_DELETED,
  });
  const document = {
    version: QUERY_CANDIDATE_PLANNER_UPLOAD_INVALIDATION_VERSION,
    reason: CACHE_INVALIDATION_REASONS.UPLOAD_DELETED,
    tenantDigest: normalizeSha(result.tenantDigest, "tenantDigest", {
      required: true,
    }),
    tagDigestSha256: sha256(tags),
    scanned: Number(result.scanned || 0),
    removed: Number(result.removed || 0),
    corruptRemoved: Number(result.corruptRemoved || 0),
    plaintextIdentifiersIncluded: false,
  };
  document.invalidationSha256 = sha256(document);
  return Object.freeze(document);
}

function buildContractRotationIdentityProbe({
  buildIdentity,
  base,
  variants = {},
} = {}) {
  if (typeof buildIdentity !== "function") {
    throw new TypeError("buildIdentity 함수가 필요합니다.");
  }
  const baseline = buildIdentity(base);
  const results = {};
  for (const [name, override] of Object.entries(variants)) {
    const identity = buildIdentity({ ...base, ...override });
    results[name] = {
      keyDigest: identity.keyDigest,
      changed: identity.keyDigest !== baseline.keyDigest,
    };
  }
  const document = {
    version: "query_candidate_planner_cache_contract_rotation_probe_v1",
    baselineKeyDigest: baseline.keyDigest,
    variants: results,
    allRotationsInvalidate: Object.values(results).every(
      (result) => result.changed === true,
    ),
  };
  document.probeSha256 = sha256(stableStringify(document));
  return Object.freeze(document);
}

module.exports = {
  QUERY_CANDIDATE_PLANNER_CACHE_REPLAY_AUDIT_VERSION,
  QUERY_CANDIDATE_PLANNER_UPLOAD_INVALIDATION_VERSION,
  buildReplaySafeShadowFingerprint,
  compareReplaySafeShadowResolutions,
  buildUploadInvalidationTags,
  invalidateCandidatePlannerUploadCache,
  buildContractRotationIdentityProbe,
};
