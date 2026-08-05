"use strict";

const crypto = require("crypto");
const {
  deriveQueryCandidatePlannerInternalCanarySubject,
} = require("./queryCandidatePlannerInternalCanarySubject");
const {
  buildShadowAccuracyObservation,
  findForbiddenPaths,
} = require("./queryCandidatePlannerShadowAccuracyEvaluator");
const {
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
  resolveRegistryCase,
} = require("./queryCandidatePlannerRealShadowEvidenceConfig");
const {
  takeQueryCandidatePlannerRealShadowCapture,
} = require("./queryCandidatePlannerRealShadowCaptureBridge");
const {
  createMongoRealShadowEvidenceStore,
} = require("./queryCandidatePlannerRealShadowEvidenceStore");

const COLLECTOR_VERSION =
  "query_candidate_planner_real_shadow_evidence_collector_v1";
const RECORD_VERSION =
  "query_candidate_planner_real_shadow_evidence_record_v1";
const SHA256_RE = /^[a-f0-9]{64}$/i;

function text(value, maxLength = 160) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : 0;
}

function sha256(value) {
  return crypto
    .createHash("sha256")
    .update(typeof value === "string" ? value : JSON.stringify(value))
    .digest("hex");
}

function safeSha256(value) {
  const normalized = text(value, 64).toLowerCase();
  return SHA256_RE.test(normalized) ? normalized : "";
}

function subjectAllowed(config, subject) {
  return subject.complete === true &&
    config.allowlist.includes(String(subject.subjectSha256).toLowerCase());
}

function nowMs(now) {
  const value = typeof now === "function" ? Number(now()) : Number(now);
  return Number.isFinite(value) ? value : Date.now();
}

function expiresAt(observedAt, ttlDays) {
  return new Date(Date.parse(observedAt) + ttlDays * 86400000).toISOString();
}

function executionStatus(observation = {}) {
  const status = text(observation.status, 60).toUpperCase();
  if (status.includes("TIMEOUT")) return "TIMEOUT";
  if (status.includes("FAILED") || status.includes("ERROR")) return "ERROR";
  if (status === "BLOCKED") return "BLOCKED";
  return "SUCCESS";
}

function normalizeCache(resolution = {}, observation = {}) {
  const cache = resolution.cache || {};
  const hit = cache.hit === true;
  let level = text(cache.level, 20).toUpperCase();
  if (hit && !["L1", "L2", "L3", "L4"].includes(level)) level = "L1";
  if (!hit && !["MISS", "NONE"].includes(level)) {
    level = observation.cacheReadDecision?.allowed === true ? "MISS" : "NONE";
  }
  return Object.freeze({
    readAttempted:
      cache.readAttempted === true ||
      observation.cacheReadDecision?.allowed === true,
    hit,
    level,
    writeAttempted:
      cache.writeAttempted === true ||
      observation.cacheWriteDecision?.allowed === true,
    writeSucceeded: cache.writeSucceeded === true,
  });
}

function normalizeProvider(resolution = {}) {
  const invocation = resolution.plannerResolution?.invocation || {};
  const providerCallCount = Math.trunc(number(invocation.providerCallCount));
  return Object.freeze({
    called: providerCallCount > 0,
    modelId: text(invocation.modelId, 120),
    inputTokens: Math.trunc(number(invocation.inputTokens)),
    outputTokens: Math.trunc(number(invocation.outputTokens)),
    observedCostMicrousd: Math.trunc(number(invocation.observedCostMicrousd)),
    providerCallCount,
  });
}

function buildExecutionRecord({
  observation,
  request,
  config,
  capture,
  registryCase,
  subject,
  now = Date.now,
} = {}) {
  const observedAt = new Date(nowMs(now)).toISOString();
  const requestFingerprintSha256 = safeSha256(
    observation.requestFingerprintSha256,
  );
  const uploadFingerprintSha256 = safeSha256(
    observation.cacheLifecycle?.identity?.uploadFingerprintSha256,
  );
  const resolution = capture?.resolution || {};
  const shadowAccuracyObservation = buildShadowAccuracyObservation({
    caseId: registryCase.caseId,
    apiShadowObservation: observation,
    shadowResolution: resolution,
  });
  const provider = normalizeProvider(resolution);
  const cache = normalizeCache(resolution, observation);
  const payload = Object.freeze({
    version: RECORD_VERSION,
    kind: "EXECUTION",
    source: "REAL_SHADOW_TRAFFIC",
    actualTraffic: true,
    synthetic: false,
    route: "POST /api/automation/analysis-candidates",
    caseId: registryCase.caseId,
    scenarioId: registryCase.scenarioId,
    shadowAccuracyObservation,
    operational: Object.freeze({
      status: executionStatus(observation),
      latencyMs: Math.max(0, number(observation.latencyMs)),
      expectedColdCostMicrousd: registryCase.expectedColdCostMicrousd,
      modelIdFallback: registryCase.modelId,
      cache,
      provider,
      lifecycleHints: Object.freeze({
        uploadFingerprintSha256,
        afterDownload: false,
        afterReupload: false,
        staleCacheReused: false,
      }),
    }),
    guardrails: Object.freeze({
      primaryResponseUnchanged: observation.primaryResponseUnchanged !== false,
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      plannerEscalationAllowed: false,
      semanticProfilerOnly: true,
    }),
    privacy: Object.freeze({
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      fileNameIncluded: false,
      originalFileNameIncluded: false,
      userIdentityIncluded: false,
      tenantIdIncluded: false,
      rawProviderResponseIncluded: false,
    }),
  });
  const forbidden = findForbiddenPaths(payload);
  if (forbidden.length > 0) {
    const error = new Error(`forbidden evidence payload field: ${forbidden[0]}`);
    error.code = "REAL_SHADOW_EVIDENCE_PRIVACY_VIOLATION";
    throw error;
  }
  const recordId = sha256({
    version: RECORD_VERSION,
    kind: "EXECUTION",
    observedAt,
    requestFingerprintSha256,
    observationId: shadowAccuracyObservation.observationId,
  });
  return Object.freeze({
    recordId,
    kind: "EXECUTION",
    observedAt,
    expiresAt: expiresAt(observedAt, config.ttlDays),
    subjectTagSha256: subject.subjectTagSha256,
    requestFingerprintSha256,
    uploadFingerprintSha256,
    caseId: registryCase.caseId,
    scenarioId: registryCase.scenarioId,
    payload,
  });
}

function lifecycleEventName(action) {
  if (action === "DELETE") return "DELETE";
  if (action === "UPLOAD_REPLACEMENT") return "REUPLOAD";
  return "DOWNLOAD";
}

function buildLifecycleRecord({
  observation,
  config,
  registryCase,
  subject,
  now = Date.now,
} = {}) {
  const observedAt = new Date(nowMs(now)).toISOString();
  const uploadFingerprintSha256 = safeSha256(
    observation.identity?.uploadFingerprintSha256,
  );
  const event = lifecycleEventName(text(observation.action, 80));
  const payload = Object.freeze({
    version: RECORD_VERSION,
    kind: "LIFECYCLE",
    source: "REAL_SHADOW_TRAFFIC",
    actualTraffic: true,
    synthetic: false,
    caseId: registryCase.caseId,
    scenarioId: registryCase.scenarioId,
    lifecycle: Object.freeze({
      event,
      action: text(observation.action, 80),
      cacheDisposition: text(observation.cacheDisposition, 80) || "UNKNOWN",
      invalidationAttempted: observation.invalidation != null,
      invalidationSucceeded:
        observation.invalidation?.invalidated === true ||
        observation.cacheDisposition === "INVALIDATED",
      staleCacheReused: false,
      uploadFingerprintSha256,
    }),
    privacy: Object.freeze({
      rawRowsIncluded: false,
      fileNameIncluded: false,
      originalFileNameIncluded: false,
      userIdentityIncluded: false,
      tenantIdIncluded: false,
      storageObjectKeyIncluded: false,
      cacheSecretIncluded: false,
    }),
  });
  const recordId = sha256({
    version: RECORD_VERSION,
    kind: "LIFECYCLE",
    observedAt,
    event,
    uploadFingerprintSha256,
    caseId: registryCase.caseId,
  });
  return Object.freeze({
    recordId,
    kind: "LIFECYCLE",
    observedAt,
    expiresAt: expiresAt(observedAt, config.ttlDays),
    subjectTagSha256: subject.subjectTagSha256,
    requestFingerprintSha256: "",
    uploadFingerprintSha256,
    caseId: registryCase.caseId,
    scenarioId: registryCase.scenarioId,
    payload,
  });
}

function runtimeStore(config) {
  return createMongoRealShadowEvidenceStore({ secret: config.secret });
}

async function recordQueryCandidatePlannerRealShadowObservation(
  observation = {},
  context = {},
  {
    env = process.env,
    store = null,
    capture = null,
    now = Date.now,
  } = {},
) {
  const config = parseQueryCandidatePlannerRealShadowEvidenceConfig(env);
  if (!config.enabled) return Object.freeze({ stored: false, reason: "COLLECTOR_DISABLED" });
  if (!config.configurationValid) return Object.freeze({ stored: false, reason: config.reason });
  if (observation.version !== "query_candidate_planner_api_shadow_observation_v1") {
    return Object.freeze({ stored: false, reason: "API_SHADOW_OBSERVATION_REQUIRED" });
  }
  const request = context.req || {};
  const subject = deriveQueryCandidatePlannerInternalCanarySubject(request);
  if (!subjectAllowed(config, subject)) {
    return Object.freeze({ stored: false, reason: "COLLECTOR_SUBJECT_NOT_ALLOWLISTED" });
  }
  const requestFingerprintSha256 = safeSha256(observation.requestFingerprintSha256);
  const uploadFingerprintSha256 = safeSha256(
    observation.cacheLifecycle?.identity?.uploadFingerprintSha256,
  );
  const registryCase = resolveRegistryCase(config, {
    requestFingerprintSha256,
    uploadFingerprintSha256,
  });
  if (!registryCase) {
    return Object.freeze({ stored: false, reason: "REAL_SHADOW_CASE_NOT_REGISTERED" });
  }
  const captured = capture ||
    takeQueryCandidatePlannerRealShadowCapture(requestFingerprintSha256);
  if (!captured && ["COMPLETED", "COMPLETED_SAFE"].includes(observation.status)) {
    return Object.freeze({ stored: false, reason: "REAL_SHADOW_RESOLUTION_CAPTURE_REQUIRED" });
  }
  try {
    const record = buildExecutionRecord({
      observation,
      request,
      config,
      capture: captured || { resolution: {} },
      registryCase,
      subject,
      now,
    });
    return await (store || runtimeStore(config)).record(record);
  } catch (error) {
    return Object.freeze({
      stored: false,
      reason: String(error?.code || "REAL_SHADOW_EVIDENCE_COLLECTION_FAILED"),
    });
  }
}

async function recordQueryCandidatePlannerRealShadowLifecycleObservation(
  observation = {},
  context = {},
  { env = process.env, store = null, now = Date.now } = {},
) {
  const config = parseQueryCandidatePlannerRealShadowEvidenceConfig(env);
  if (!config.enabled) return Object.freeze({ stored: false, reason: "COLLECTOR_DISABLED" });
  if (!config.configurationValid) return Object.freeze({ stored: false, reason: config.reason });
  const request = context.req || {};
  const subject = deriveQueryCandidatePlannerInternalCanarySubject(request);
  if (!subjectAllowed(config, subject)) {
    return Object.freeze({ stored: false, reason: "COLLECTOR_SUBJECT_NOT_ALLOWLISTED" });
  }
  const uploadFingerprintSha256 = safeSha256(
    observation.identity?.uploadFingerprintSha256,
  );
  const registryCase = resolveRegistryCase(config, { uploadFingerprintSha256 });
  if (!registryCase) {
    return Object.freeze({ stored: false, reason: "REAL_SHADOW_CASE_NOT_REGISTERED" });
  }
  try {
    const record = buildLifecycleRecord({
      observation,
      config,
      registryCase,
      subject,
      now,
    });
    return await (store || runtimeStore(config)).record(record);
  } catch (error) {
    return Object.freeze({
      stored: false,
      reason: String(error?.code || "REAL_SHADOW_LIFECYCLE_COLLECTION_FAILED"),
    });
  }
}

module.exports = Object.freeze({
  COLLECTOR_VERSION,
  RECORD_VERSION,
  buildExecutionRecord,
  buildLifecycleRecord,
  recordQueryCandidatePlannerRealShadowObservation,
  recordQueryCandidatePlannerRealShadowLifecycleObservation,
});
