"use strict";

const crypto = require("crypto");
const {
  getQueryCandidatePlannerInternalPreviewConfig,
} = require("./queryCandidatePlannerInternalPreviewConfig");

const STORE_VERSION =
  "query_candidate_planner_internal_preview_store_v1";
const ENTRY_VERSION =
  "query_candidate_planner_internal_preview_entry_v1";

function text(value, maxLength = 160) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function number(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function boolean(value) {
  return value === true;
}

function sha256(value) {
  return crypto
    .createHash("sha256")
    .update(String(value == null ? "" : value))
    .digest("hex");
}

function safeSha256(value) {
  const normalized = text(value, 128).toLowerCase();
  return /^[a-f0-9]{64}$/.test(normalized) ? normalized : "";
}

function sanitizeDecision(decision = {}) {
  return Object.freeze({
    operation: text(decision.operation || decision.operationName, 80),
    allowed: decision.allowed === true,
    reason: text(decision.reason, 120),
  });
}

function sanitizeCounts(counts = {}) {
  return Object.freeze({
    primary: number(counts.primary),
    shadow: number(counts.shadow),
    shared: number(counts.shared),
    primaryOnly: number(counts.primaryOnly),
    shadowOnly: number(counts.shadowOnly),
  });
}

function sanitizeMetrics(metrics = {}) {
  return Object.freeze({
    exactOrder: boolean(metrics.exactOrder),
    top1Same: boolean(metrics.top1Same),
    top3Overlap: number(metrics.top3Overlap),
    jaccard: number(metrics.jaccard),
    rankAgreement: number(metrics.rankAgreement),
  });
}

function sanitizeComparison(comparison = null) {
  if (!comparison || typeof comparison !== "object") return null;
  return Object.freeze({
    version: text(comparison.version, 100),
    policyVersion: text(comparison.policyVersion, 100),
    verdict: text(comparison.verdict || "NOT_AVAILABLE", 60),
    counts: sanitizeCounts(comparison.counts),
    metrics: sanitizeMetrics(comparison.metrics),
    fingerprints: Object.freeze({
      primaryOrderSha256: safeSha256(
        comparison.fingerprints?.primaryOrderSha256,
      ),
      shadowOrderSha256: safeSha256(
        comparison.fingerprints?.shadowOrderSha256,
      ),
      sharedSetSha256: safeSha256(
        comparison.fingerprints?.sharedSetSha256,
      ),
      rawIdentifiersIncluded: false,
    }),
  });
}

function sanitizeCacheIdentity(identity = {}) {
  return Object.freeze({
    version: text(identity.version, 100),
    complete: identity.complete === true,
    reason: text(identity.reason, 120),
    source: text(identity.source, 80),
    uploadFingerprintSha256: safeSha256(
      identity.uploadFingerprintSha256,
    ),
    queryJsonSha256: safeSha256(identity.queryJsonSha256),
    tenantIdIncluded: false,
    originalFileNameIncluded: false,
    queryTablesKeyIncluded: false,
  });
}

function sanitizeGuardrails(guardrails = {}) {
  return Object.freeze({
    shadowOnly: guardrails.shadowOnly !== false,
    primaryResponseAuthority:
      guardrails.primaryResponseAuthority !== false,
    responsePayloadMutation: false,
    responseHeaderMutation: false,
    responseStatusMutation: false,
    productionCandidateMerge: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    candidateExecutionAvailable: false,
    candidateSelectionAvailable: false,
  });
}

function sanitizeObservation(observation = {}, {
  observedAt = new Date().toISOString(),
  sequence = 0,
} = {}) {
  const requestFingerprintSha256 = safeSha256(
    observation.requestFingerprintSha256,
  );
  const primaryResponseSha256 = safeSha256(
    observation.primaryResponseSha256,
  );
  const id = sha256([
    ENTRY_VERSION,
    observedAt,
    sequence,
    requestFingerprintSha256,
    text(observation.status),
  ].join(":"));

  return Object.freeze({
    version: ENTRY_VERSION,
    id,
    observedAt,
    status: text(observation.status || "UNKNOWN", 60),
    reason: text(observation.reason, 120),
    requestFingerprintSha256,
    primaryResponseSha256,
    primaryResponseUnchanged:
      observation.primaryResponseUnchanged !== false,
    latencyMs: Math.max(0, number(observation.latencyMs)),
    decisions: Object.freeze({
      shadow: sanitizeDecision(observation.featureDecision),
      provider: sanitizeDecision(observation.providerDecision),
      cacheRead: sanitizeDecision(observation.cacheReadDecision),
      cacheWrite: sanitizeDecision(observation.cacheWriteDecision),
    }),
    shadow: Object.freeze({
      status: text(observation.shadow?.status, 80),
      invocationStatus: text(
        observation.shadow?.invocationStatus,
        80,
      ),
      providerCallCount: Math.max(
        0,
        number(observation.shadow?.providerCallCount),
      ),
      accepted: Math.max(0, number(observation.shadow?.accepted)),
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
    }),
    cacheLifecycle: Object.freeze({
      identity: sanitizeCacheIdentity(
        observation.cacheLifecycle?.identity,
      ),
      cacheReadAllowed:
        observation.cacheLifecycle?.cacheReadAllowed === true,
      cacheWriteAllowed:
        observation.cacheLifecycle?.cacheWriteAllowed === true,
      tenantIdIncluded: false,
      cacheSecretIncluded: false,
    }),
    comparison: sanitizeComparison(observation.comparison),
    guardrails: sanitizeGuardrails(observation.guardrails),
    privacy: Object.freeze({
      rawPrimaryResponseIncluded: false,
      rawCandidatePayloadIncluded: false,
      rawIdentifiersIncluded: false,
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      fileNameIncluded: false,
      originalFileNameIncluded: false,
      queryTablesKeyIncluded: false,
      tenantIdIncluded: false,
      emailIncluded: false,
      rawErrorMessageIncluded: false,
    }),
  });
}

function createQueryCandidatePlannerInternalPreviewStore({
  maxEntries = 100,
  ttlMs = 24 * 60 * 60 * 1000,
  now = Date.now,
} = {}) {
  const capacity = Math.max(1, Math.min(500, Number(maxEntries) || 100));
  const ttl = Math.max(1000, Number(ttlMs) || 24 * 60 * 60 * 1000);
  let sequence = 0;
  let entries = [];

  function nowMs() {
    const value = Number(now());
    return Number.isFinite(value) ? value : Date.now();
  }

  function prune() {
    const threshold = nowMs() - ttl;
    entries = entries.filter((entry) => {
      const observed = Date.parse(entry.observedAt);
      return Number.isFinite(observed) && observed >= threshold;
    });
    if (entries.length > capacity) {
      entries = entries.slice(entries.length - capacity);
    }
  }

  function record(observation = {}) {
    prune();
    sequence += 1;
    const observedAt = new Date(nowMs()).toISOString();
    const entry = sanitizeObservation(observation, {
      observedAt,
      sequence,
    });
    entries.push(entry);
    if (entries.length > capacity) entries.shift();
    return entry;
  }

  function list({ limit = 50, status = "", verdict = "" } = {}) {
    prune();
    const normalizedStatus = text(status, 60).toUpperCase();
    const normalizedVerdict = text(verdict, 60).toUpperCase();
    const max = Math.max(1, Math.min(100, Number(limit) || 50));
    return Object.freeze(
      entries
        .filter((entry) =>
          normalizedStatus
            ? entry.status.toUpperCase() === normalizedStatus
            : true,
        )
        .filter((entry) =>
          normalizedVerdict
            ? text(entry.comparison?.verdict).toUpperCase() ===
              normalizedVerdict
            : true,
        )
        .slice()
        .reverse()
        .slice(0, max),
    );
  }

  function summary() {
    prune();
    const statusCounts = {};
    const verdictCounts = {};
    let latencyTotal = 0;
    let providerCallTotal = 0;
    for (const entry of entries) {
      statusCounts[entry.status] = (statusCounts[entry.status] || 0) + 1;
      const verdict = entry.comparison?.verdict || "NOT_AVAILABLE";
      verdictCounts[verdict] = (verdictCounts[verdict] || 0) + 1;
      latencyTotal += entry.latencyMs;
      providerCallTotal += entry.shadow.providerCallCount;
    }
    return Object.freeze({
      version: STORE_VERSION,
      total: entries.length,
      capacity,
      ttlMs: ttl,
      lastObservedAt: entries.at(-1)?.observedAt || "",
      averageLatencyMs: entries.length
        ? Math.round((latencyTotal / entries.length) * 100) / 100
        : 0,
      providerCallTotal,
      statusCounts: Object.freeze({ ...statusCounts }),
      verdictCounts: Object.freeze({ ...verdictCounts }),
      persistence: "MEMORY_ONLY",
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
    });
  }

  function clearForTests() {
    entries = [];
    sequence = 0;
  }

  return Object.freeze({
    version: STORE_VERSION,
    record,
    list,
    summary,
    clearForTests,
  });
}

let runtimeStore = null;
let runtimeStoreSignature = "";

function getQueryCandidatePlannerInternalPreviewStore() {
  const config = getQueryCandidatePlannerInternalPreviewConfig();
  const signature = `${config.maxEntries}:${config.ttlMs}`;
  if (!runtimeStore || runtimeStoreSignature !== signature) {
    runtimeStore = createQueryCandidatePlannerInternalPreviewStore({
      maxEntries: config.maxEntries,
      ttlMs: config.ttlMs,
    });
    runtimeStoreSignature = signature;
  }
  return runtimeStore;
}

function recordQueryCandidatePlannerInternalPreviewObservation(
  observation = {},
) {
  const config = getQueryCandidatePlannerInternalPreviewConfig();
  if (!config.enabled) return null;
  try {
    return getQueryCandidatePlannerInternalPreviewStore().record(
      observation,
    );
  } catch (_error) {
    return null;
  }
}

function resetQueryCandidatePlannerInternalPreviewStoreForTests({
  store = null,
} = {}) {
  runtimeStore = store;
  if (store) {
    const config = getQueryCandidatePlannerInternalPreviewConfig();
    runtimeStoreSignature = `${config.maxEntries}:${config.ttlMs}`;
  } else {
    runtimeStoreSignature = "";
  }
  return runtimeStore;
}

module.exports = Object.freeze({
  STORE_VERSION,
  ENTRY_VERSION,
  safeSha256,
  sanitizeObservation,
  createQueryCandidatePlannerInternalPreviewStore,
  getQueryCandidatePlannerInternalPreviewStore,
  recordQueryCandidatePlannerInternalPreviewObservation,
  resetQueryCandidatePlannerInternalPreviewStoreForTests,
});
