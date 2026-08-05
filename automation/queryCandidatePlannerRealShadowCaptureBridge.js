"use strict";

const crypto = require("crypto");
const {
  runQueryCandidatePlannerApiShadow,
} = require("./queryCandidatePlannerApiShadowRunner");

const BRIDGE_VERSION =
  "query_candidate_planner_real_shadow_capture_bridge_v1";
const MAX_CAPTURE_AGE_MS = 2 * 60 * 1000;
const MAX_CAPTURES = 1000;
const captures = new Map();

function text(value, maxLength = 160) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : 0;
}

function boolean(value) {
  return value === true;
}

function first(...values) {
  for (const value of values) {
    if (value !== undefined && value !== null && value !== "") return value;
  }
  return undefined;
}

function safeCandidates(resolution = {}) {
  const lists = [
    resolution.plannerResolution?.items,
    resolution.plannerResolution?.candidates,
    resolution.candidateResolution?.items,
    resolution.rankingResolution?.items,
    resolution.items,
    resolution.candidates,
    resolution.topCandidates,
  ].filter(Array.isArray);
  const seen = new Set();
  const output = [];
  for (const list of lists) {
    for (const [index, candidate] of list.entries()) {
      if (!candidate || typeof candidate !== "object") continue;
      const candidateId = text(
        candidate.candidateId || candidate.id || candidate.recipeId || candidate.recipeType,
        160,
      );
      if (!candidateId || seen.has(candidateId)) continue;
      seen.add(candidateId);
      output.push(Object.freeze({
        candidateId,
        rank: Number.isInteger(candidate.rank) && candidate.rank > 0
          ? candidate.rank
          : index + 1,
        status: text(candidate.status || "ACCEPTED", 40).toUpperCase(),
        productionEligible: candidate.productionEligible !== false,
      }));
      if (output.length >= 100) break;
    }
    if (output.length >= 100) break;
  }
  return Object.freeze(output);
}

function sanitizeShadowResolutionForEvidence(resolution = {}) {
  const invocation =
    resolution.plannerResolution?.invocation ||
    resolution.invocation ||
    resolution.providerInvocation ||
    {};
  const usage = invocation.usage || resolution.usage || resolution.tokenUsage || {};
  const cache =
    resolution.plannerResolution?.cache ||
    resolution.cache ||
    resolution.cacheResolution ||
    {};
  const fallback =
    resolution.fallback || resolution.plannerResolution?.fallback || {};
  return Object.freeze({
    status: text(resolution.status || "UNKNOWN", 80),
    candidates: safeCandidates(resolution),
    businessDomainProfile: Object.freeze({
      primaryDomain: text(
        resolution.businessDomainProfile?.primaryDomain ||
          resolution.semanticProfile?.primaryDomain ||
          resolution.primaryDomain ||
          resolution.domain ||
          "UNKNOWN",
        100,
      ),
      datasetIntent: text(
        resolution.businessDomainProfile?.datasetIntent ||
          resolution.semanticProfile?.datasetIntent ||
          resolution.datasetIntent ||
          resolution.intent ||
          "UNKNOWN",
        100,
      ),
      reviewRequired:
        resolution.businessDomainProfile?.reviewRequired === true ||
        resolution.semanticProfile?.reviewRequired === true ||
        resolution.reviewRequired === true,
    }),
    fallback: Object.freeze({
      applied: resolution.fallbackApplied === true || fallback.applied === true,
      reason: text(resolution.fallbackReason || fallback.reason, 120),
    }),
    unsupportedRejected: resolution.unsupportedRejected === true,
    plannerResolution: Object.freeze({
      reviewRequired: resolution.plannerResolution?.reviewRequired === true,
      invocation: Object.freeze({
        status: text(invocation.status, 80),
        providerCallCount: number(first(
          invocation.providerCallCount,
          resolution.providerCallCount,
        )),
        modelId: text(first(
          invocation.modelId,
          invocation.model,
          resolution.modelId,
        ), 120),
        inputTokens: Math.trunc(number(first(
          usage.inputTokens,
          usage.promptTokens,
          invocation.inputTokens,
        ))),
        outputTokens: Math.trunc(number(first(
          usage.outputTokens,
          usage.completionTokens,
          invocation.outputTokens,
        ))),
        observedCostMicrousd: Math.trunc(number(first(
          invocation.observedCostMicrousd,
          resolution.observedCostMicrousd,
        ))),
      }),
    }),
    cache: Object.freeze({
      readAttempted: boolean(first(cache.readAttempted, cache.readAllowed)),
      hit: boolean(first(cache.hit, cache.cacheHit)),
      level: text(first(cache.level, cache.cacheLevel, cache.hitLevel), 20).toUpperCase(),
      writeAttempted: boolean(first(cache.writeAttempted, cache.writeAllowed)),
      writeSucceeded: boolean(first(cache.writeSucceeded, cache.written)),
    }),
    policy: Object.freeze({
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      plannerEscalationAllowed: false,
    }),
    privacy: Object.freeze({
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      fileNameIncluded: false,
      userIdentityIncluded: false,
      rawProviderResponseIncluded: false,
    }),
  });
}

function prune(now = Date.now()) {
  for (const [key, entry] of captures.entries()) {
    if (now - entry.capturedAt > MAX_CAPTURE_AGE_MS) captures.delete(key);
  }
  while (captures.size > MAX_CAPTURES) {
    captures.delete(captures.keys().next().value);
  }
}

function captureKey(requestFingerprintSha256) {
  return String(requestFingerprintSha256 || "").toLowerCase();
}

async function runQueryCandidatePlannerApiShadowWithRealEvidenceCapture(args = {}) {
  const resolution = await runQueryCandidatePlannerApiShadow(args);
  const key = captureKey(args.safeContext?.requestFingerprintSha256);
  if (key) {
    prune();
    captures.set(key, Object.freeze({
      version: BRIDGE_VERSION,
      capturedAt: Date.now(),
      resolution: sanitizeShadowResolutionForEvidence(resolution),
      captureSha256: crypto
        .createHash("sha256")
        .update(JSON.stringify(sanitizeShadowResolutionForEvidence(resolution)))
        .digest("hex"),
    }));
  }
  return resolution;
}

function takeQueryCandidatePlannerRealShadowCapture(requestFingerprintSha256) {
  prune();
  const key = captureKey(requestFingerprintSha256);
  const entry = captures.get(key) || null;
  if (key) captures.delete(key);
  return entry;
}

function resetQueryCandidatePlannerRealShadowCapturesForTests() {
  captures.clear();
}

module.exports = Object.freeze({
  BRIDGE_VERSION,
  MAX_CAPTURE_AGE_MS,
  sanitizeShadowResolutionForEvidence,
  runQueryCandidatePlannerApiShadowWithRealEvidenceCapture,
  takeQueryCandidatePlannerRealShadowCapture,
  resetQueryCandidatePlannerRealShadowCapturesForTests,
});
