"use strict";

const crypto = require("crypto");
const {
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
} = require("./queryCandidatePlannerRealShadowEvidenceConfig");
const {
  evaluateRealShadowLimitedActivationRuntime,
} = require("./queryCandidatePlannerRealShadowLimitedActivation");

const COLLECTION_VERSION =
  "query_candidate_planner_real_shadow_observation_collection_v1";
const REQUIRED_CASES = 10;
const REQUIRED_EXECUTIONS_PER_CASE = 5;
const REQUIRED_EXECUTIONS_TOTAL = 50;
const REQUIRED_DOWNLOADS_PER_CASE = 1;
const REQUIRED_DELETES_PER_CASE = 1;
const REQUIRED_LIFECYCLE_TOTAL = 20;
const REQUIRED_UPLOAD_IDENTITIES_PER_CASE = 2;
const BUILDER_MIN_EXECUTIONS_TOTAL = 30;
const BUILDER_MIN_EXECUTIONS_PER_CASE = 3;

function text(value, maxLength = 240) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function sha256(value) {
  return crypto
    .createHash("sha256")
    .update(typeof value === "string" ? value : JSON.stringify(value))
    .digest("hex");
}

function iso(value) {
  const ms = Date.parse(String(value || ""));
  return Number.isFinite(ms) ? new Date(ms).toISOString() : "";
}

function recordTime(record) {
  return iso(record?.observedAt || record?.payload?.observedAt);
}

function isActualRealShadowRecord(record) {
  const payload = record?.payload || {};
  return Boolean(
    record &&
    typeof record === "object" &&
    (record.source === "REAL_SHADOW_TRAFFIC" || payload.source === "REAL_SHADOW_TRAFFIC") &&
    (record.actualTraffic === true || payload.actualTraffic === true) &&
    record.synthetic !== true &&
    payload.synthetic !== true
  );
}

function withinWindow(record, from, to) {
  const when = Date.parse(recordTime(record));
  if (!Number.isFinite(when)) return false;
  const fromMs = Date.parse(String(from || ""));
  const toMs = Date.parse(String(to || ""));
  if (Number.isFinite(fromMs) && when < fromMs) return false;
  if (Number.isFinite(toMs) && when > toMs) return false;
  return true;
}

function registryCases(config) {
  const cases = Array.isArray(config?.registry?.registry?.cases)
    ? config.registry.registry.cases
    : [];
  const byId = new Map();
  for (const item of cases) {
    const caseId = text(item?.caseId, 160);
    if (caseId && !byId.has(caseId)) byId.set(caseId, item);
  }
  return [...byId.values()].sort((a, b) =>
    String(a.caseId).localeCompare(String(b.caseId)),
  );
}

function walk(value, visitor, path = "") {
  if (!value || typeof value !== "object") return false;
  if (Array.isArray(value)) {
    return value.some((item, index) => walk(item, visitor, `${path}[${index}]`));
  }
  for (const [key, item] of Object.entries(value)) {
    const nextPath = path ? `${path}.${key}` : key;
    if (visitor(key, item, nextPath)) return true;
    if (walk(item, visitor, nextPath)) return true;
  }
  return false;
}

function privacyViolation(record) {
  return walk(record?.payload || {}, (key, value) =>
    /included$/i.test(key) && value === true,
  );
}

function guardrailViolation(record) {
  const unsafeTrue = new Set([
    "productionCandidateMerge",
    "productionReadyAssignment",
    "productionRouteChanged",
    "plannerEscalationAllowed",
    "responsePayloadMutation",
    "responseHeaderMutation",
    "responseStatusMutation",
  ]);
  return walk(record?.payload || {}, (key, value) => {
    if (unsafeTrue.has(key) && value === true) return true;
    if (key === "semanticProfilerOnly" && value === false) return true;
    if (key === "primaryResponseUnchanged" && value === false) return true;
    return false;
  });
}

function lifecycleEvent(record) {
  return text(record?.payload?.lifecycle?.event, 80).toUpperCase();
}

function executionIdentity(record) {
  return text(
    record?.uploadFingerprintSha256 ||
      record?.payload?.operational?.lifecycleHints?.uploadFingerprintSha256,
    64,
  ).toLowerCase();
}

function summarizeCase(caseId, records) {
  const executions = records.filter((record) =>
    record.kind === "EXECUTION" && text(record.caseId, 160) === caseId,
  );
  const lifecycle = records.filter((record) =>
    record.kind === "LIFECYCLE" && text(record.caseId, 160) === caseId,
  );
  const downloads = lifecycle.filter((record) => lifecycleEvent(record) === "DOWNLOAD");
  const deletes = lifecycle.filter((record) => lifecycleEvent(record) === "DELETE");
  const identities = new Set(
    executions.map(executionIdentity).filter(Boolean),
  );
  const builderMinimumReady = executions.length >= BUILDER_MIN_EXECUTIONS_PER_CASE;
  const protocolReady =
    executions.length >= REQUIRED_EXECUTIONS_PER_CASE &&
    downloads.length >= REQUIRED_DOWNLOADS_PER_CASE &&
    deletes.length >= REQUIRED_DELETES_PER_CASE &&
    identities.size >= REQUIRED_UPLOAD_IDENTITIES_PER_CASE;
  return Object.freeze({
    caseId,
    executionCount: executions.length,
    lifecycleCount: lifecycle.length,
    downloadCount: downloads.length,
    deleteCount: deletes.length,
    distinctUploadIdentityCount: identities.size,
    builderMinimumReady,
    protocolReady,
  });
}

function evaluateRealShadowObservationCollection({
  records = [],
  env = process.env,
  from = "",
  to = "",
} = {}) {
  const limitedRuntime = evaluateRealShadowLimitedActivationRuntime({ env });
  const config = parseQueryCandidatePlannerRealShadowEvidenceConfig(env);
  const cases = registryCases(config);
  const filtered = records
    .filter(isActualRealShadowRecord)
    .filter((record) => withinWindow(record, from, to))
    .slice()
    .sort((a, b) =>
      recordTime(a).localeCompare(recordTime(b)) ||
      text(a.recordId, 128).localeCompare(text(b.recordId, 128)),
    );

  const caseSummaries = cases.map((item) => summarizeCase(item.caseId, filtered));
  const executionCount = filtered.filter((record) => record.kind === "EXECUTION").length;
  const lifecycleCount = filtered.filter((record) => record.kind === "LIFECYCLE").length;
  const privacyViolationCount = filtered.filter(privacyViolation).length;
  const guardrailViolationCount = filtered.filter(guardrailViolation).length;
  const builderMinimumReady =
    executionCount >= BUILDER_MIN_EXECUTIONS_TOTAL &&
    cases.length === REQUIRED_CASES &&
    caseSummaries.every((item) => item.builderMinimumReady);
  const protocolReady =
    executionCount >= REQUIRED_EXECUTIONS_TOTAL &&
    lifecycleCount >= REQUIRED_LIFECYCLE_TOTAL &&
    cases.length === REQUIRED_CASES &&
    caseSummaries.every((item) => item.protocolReady);

  const errors = [];
  if (!limitedRuntime.ready) {
    errors.push(...(limitedRuntime.errors || [limitedRuntime.reason]).map(
      (value) => `PATCH_15_3_2_D_RUNTIME:${value}`,
    ));
  }
  if (!config.registry.valid || cases.length !== REQUIRED_CASES) {
    errors.push("REAL_SHADOW_OBSERVATION_COLLECTION_REQUIRES_10_CASE_REGISTRY");
  }
  if (!builderMinimumReady) {
    errors.push("REAL_SHADOW_OBSERVATION_BUILDER_MINIMUM_NOT_MET");
  }
  if (!protocolReady) {
    errors.push("REAL_SHADOW_OBSERVATION_COLLECTION_PROTOCOL_INCOMPLETE");
  }
  if (privacyViolationCount > 0) {
    errors.push("REAL_SHADOW_OBSERVATION_PRIVACY_VIOLATION");
  }
  if (guardrailViolationCount > 0) {
    errors.push("REAL_SHADOW_OBSERVATION_GUARDRAIL_VIOLATION");
  }

  const uniqueErrors = Object.freeze([...new Set(errors)]);
  const ready = uniqueErrors.length === 0;
  const recordIds = filtered.map((record) => text(record.recordId, 128)).filter(Boolean).sort();

  return Object.freeze({
    version: COLLECTION_VERSION,
    phase: "15.3-B",
    patch: "15.3.2-E",
    ready,
    reason: ready
      ? "REAL_SHADOW_OBSERVATION_COLLECTION_COMPLETE"
      : uniqueErrors[0] || "REAL_SHADOW_OBSERVATION_COLLECTION_BLOCKED",
    errors: uniqueErrors,
    from: iso(from),
    to: iso(to),
    registryCaseCount: cases.length,
    executionCount,
    lifecycleCount,
    totalRecordCount: filtered.length,
    builderMinimumReady,
    protocolReady,
    privacyViolationCount,
    guardrailViolationCount,
    caseSummaries: Object.freeze(caseSummaries),
    collectionRecordSetSha256: sha256(recordIds),
    readyForPatch15_3_2_F: ready,
    internalCanaryEnabledByThisOperation: false,
    productionPromotionAuthorized: false,
    rawRecordsIncluded: false,
    fingerprintsIncluded: false,
  });
}

function buildObservationCollectionSummary(result, {
  finalizedAt = new Date().toISOString(),
} = {}) {
  if (!result || result.ready !== true) {
    const error = new Error(result?.reason || "REAL_SHADOW_OBSERVATION_COLLECTION_NOT_READY");
    error.code = "REAL_SHADOW_OBSERVATION_COLLECTION_NOT_READY";
    throw error;
  }
  return Object.freeze({
    version: "query_candidate_planner_real_shadow_observation_collection_summary_v1",
    phase: "15.3-B",
    patch: "15.3.2-E",
    finalizedAt: iso(finalizedAt),
    from: result.from,
    to: result.to,
    registryCaseCount: result.registryCaseCount,
    executionCount: result.executionCount,
    lifecycleCount: result.lifecycleCount,
    totalRecordCount: result.totalRecordCount,
    caseSummaries: result.caseSummaries,
    builderMinimumReady: result.builderMinimumReady,
    collectionProtocolComplete: result.protocolReady,
    privacyViolationCount: result.privacyViolationCount,
    guardrailViolationCount: result.guardrailViolationCount,
    collectionRecordSetSha256: result.collectionRecordSetSha256,
    readyForPatch15_3_2_F: true,
    collectorEnabledByThisOperation: false,
    internalCanaryEnabledByThisOperation: false,
    productionPromotionAuthorized: false,
    rawRecordsIncluded: false,
    fingerprintsIncluded: false,
    privateOutputDoNotCommit: true,
  });
}

module.exports = Object.freeze({
  COLLECTION_VERSION,
  REQUIRED_CASES,
  REQUIRED_EXECUTIONS_PER_CASE,
  REQUIRED_EXECUTIONS_TOTAL,
  REQUIRED_DOWNLOADS_PER_CASE,
  REQUIRED_DELETES_PER_CASE,
  REQUIRED_LIFECYCLE_TOTAL,
  REQUIRED_UPLOAD_IDENTITIES_PER_CASE,
  BUILDER_MIN_EXECUTIONS_TOTAL,
  BUILDER_MIN_EXECUTIONS_PER_CASE,
  isActualRealShadowRecord,
  summarizeCase,
  evaluateRealShadowObservationCollection,
  buildObservationCollectionSummary,
  sha256,
});
