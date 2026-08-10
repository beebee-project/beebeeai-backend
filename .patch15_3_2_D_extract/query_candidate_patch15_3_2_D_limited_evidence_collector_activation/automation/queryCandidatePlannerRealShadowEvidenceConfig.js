"use strict";

const crypto = require("crypto");

const CONFIG_VERSION =
  "query_candidate_planner_real_shadow_evidence_config_v1";
const REGISTRY_VERSION =
  "query_candidate_planner_real_shadow_case_registry_v1";
const SHA256_RE = /^[a-f0-9]{64}$/i;

function text(value, maxLength = 240) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function booleanEnv(value, fallback = false) {
  const normalized = text(value).toLowerCase();
  if (!normalized) return fallback;
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off"].includes(normalized)) return false;
  return fallback;
}

function integerEnv(value, fallback, { min, max }) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed)) return fallback;
  return Math.max(min, Math.min(max, parsed));
}

function sha256(value) {
  return crypto.createHash("sha256").update(String(value)).digest("hex");
}

function parseAllowlist(raw) {
  return Object.freeze(
    [...new Set(
      text(raw, 20000)
        .split(",")
        .map((entry) => entry.trim().toLowerCase())
        .filter((entry) => SHA256_RE.test(entry)),
    )],
  );
}

function parseRegistry(raw) {
  if (!text(raw, 500000)) {
    return Object.freeze({
      valid: false,
      reason: "REAL_SHADOW_CASE_REGISTRY_NOT_CONFIGURED",
      registry: null,
      byRequestFingerprint: new Map(),
      byUploadFingerprint: new Map(),
    });
  }
  let value;
  try {
    value = JSON.parse(String(raw));
  } catch (_error) {
    return Object.freeze({
      valid: false,
      reason: "REAL_SHADOW_CASE_REGISTRY_JSON_INVALID",
      registry: null,
      byRequestFingerprint: new Map(),
      byUploadFingerprint: new Map(),
    });
  }
  const errors = [];
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    errors.push("REGISTRY_OBJECT_REQUIRED");
  }
  if (value?.version !== REGISTRY_VERSION) errors.push("REGISTRY_VERSION_INVALID");
  if (!Array.isArray(value?.cases) || value.cases.length === 0) {
    errors.push("REGISTRY_CASES_REQUIRED");
  }
  const byRequestFingerprint = new Map();
  const byUploadFingerprint = new Map();
  const caseIds = new Set();
  for (const [index, item] of (value?.cases || []).entries()) {
    const caseId = text(item?.caseId, 160);
    const scenarioId = text(item?.scenarioId || caseId, 160);
    const requestFingerprintSha256 = text(
      item?.requestFingerprintSha256,
      64,
    ).toLowerCase();
    const uploadFingerprintSha256 = text(
      item?.uploadFingerprintSha256,
      64,
    ).toLowerCase();
    if (!caseId) errors.push(`cases[${index}].caseId required`);
    if (!scenarioId) errors.push(`cases[${index}].scenarioId required`);
    if (caseIds.has(caseId)) errors.push(`duplicate caseId: ${caseId}`);
    caseIds.add(caseId);
    if (!SHA256_RE.test(requestFingerprintSha256) &&
        !SHA256_RE.test(uploadFingerprintSha256)) {
      errors.push(`cases[${index}] requires request or upload fingerprint`);
    }
    const normalized = Object.freeze({
      caseId,
      scenarioId,
      requestFingerprintSha256:
        SHA256_RE.test(requestFingerprintSha256)
          ? requestFingerprintSha256
          : "",
      uploadFingerprintSha256:
        SHA256_RE.test(uploadFingerprintSha256)
          ? uploadFingerprintSha256
          : "",
      expectedColdCostMicrousd:
        Number.isInteger(item?.expectedColdCostMicrousd) &&
        item.expectedColdCostMicrousd >= 0
          ? item.expectedColdCostMicrousd
          : 0,
      modelId: text(item?.modelId || "semantic_profiler_default", 120),
    });
    if (normalized.requestFingerprintSha256) {
      if (byRequestFingerprint.has(normalized.requestFingerprintSha256)) {
        errors.push(`duplicate request fingerprint: ${normalized.requestFingerprintSha256}`);
      }
      byRequestFingerprint.set(normalized.requestFingerprintSha256, normalized);
    }
    if (normalized.uploadFingerprintSha256) {
      if (byUploadFingerprint.has(normalized.uploadFingerprintSha256)) {
        errors.push(`duplicate upload fingerprint: ${normalized.uploadFingerprintSha256}`);
      }
      byUploadFingerprint.set(normalized.uploadFingerprintSha256, normalized);
    }
  }
  return Object.freeze({
    valid: errors.length === 0,
    reason: errors.length === 0 ? "REAL_SHADOW_CASE_REGISTRY_VALID" : errors[0],
    errors: Object.freeze(errors),
    registry: errors.length === 0 ? Object.freeze({
      version: REGISTRY_VERSION,
      registryId: text(value.registryId || sha256(JSON.stringify(value)), 160),
      cases: Object.freeze([...byRequestFingerprint.values(), ...byUploadFingerprint.values()]),
    }) : null,
    byRequestFingerprint,
    byUploadFingerprint,
  });
}

function parseQueryCandidatePlannerRealShadowEvidenceConfig(env = process.env) {
  const requestedEnabled = booleanEnv(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED,
    false,
  );
  const killSwitch = booleanEnv(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH,
    false,
  );
  const enabled = requestedEnabled && !killSwitch;
  const secret = text(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET,
    1000,
  );
  const allowlist = parseAllowlist(
    env.QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256,
  );
  const registry = parseRegistry(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON,
  );
  const ttlDays = integerEnv(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_TTL_DAYS,
    7,
    { min: 1, max: 30 },
  );
  const maxRecords = integerEnv(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_MAX_RECORDS,
    5000,
    { min: 30, max: 50000 },
  );
  const errors = [];
  if (enabled && secret.length < 32) errors.push("REAL_SHADOW_EVIDENCE_SECRET_REQUIRED");
  if (enabled && allowlist.length === 0) errors.push("REAL_SHADOW_COLLECTOR_ALLOWLIST_REQUIRED");
  if (enabled && !registry.valid) errors.push(registry.reason);
  return Object.freeze({
    version: CONFIG_VERSION,
    requestedEnabled,
    killSwitch,
    enabled,
    configurationValid: errors.length === 0,
    reason: errors.length === 0
      ? enabled
        ? "REAL_SHADOW_EVIDENCE_COLLECTION_ENABLED"
        : requestedEnabled && killSwitch
          ? "REAL_SHADOW_EVIDENCE_COLLECTION_KILL_SWITCHED"
          : "REAL_SHADOW_EVIDENCE_COLLECTION_DISABLED"
      : errors[0],
    errors: Object.freeze(errors),
    ttlDays,
    maxRecords,
    secret,
    allowlist,
    registry,
    failClosed: true,
  });
}

function resolveRegistryCase(config, { requestFingerprintSha256 = "", uploadFingerprintSha256 = "" } = {}) {
  const requestKey = text(requestFingerprintSha256, 64).toLowerCase();
  const uploadKey = text(uploadFingerprintSha256, 64).toLowerCase();
  return config?.registry?.byRequestFingerprint?.get(requestKey) ||
    config?.registry?.byUploadFingerprint?.get(uploadKey) ||
    null;
}

module.exports = Object.freeze({
  CONFIG_VERSION,
  REGISTRY_VERSION,
  SHA256_RE,
  parseAllowlist,
  parseRegistry,
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
  resolveRegistryCase,
});
