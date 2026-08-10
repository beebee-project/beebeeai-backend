const crypto = require("crypto");

const CONFIG_VERSION = "query_candidate_planner_internal_preview_config_v1";
const ENABLED_ENV = "QUERY_CANDIDATE_INTERNAL_PREVIEW_ENABLED";
const TOKEN_ENV = "QUERY_CANDIDATE_INTERNAL_PREVIEW_TOKEN";
const MAX_ENTRIES_ENV = "QUERY_CANDIDATE_INTERNAL_PREVIEW_MAX_ENTRIES";
const TTL_MS_ENV = "QUERY_CANDIDATE_INTERNAL_PREVIEW_TTL_MS";
const MIN_TOKEN_BYTES = 24;
const DEFAULT_MAX_ENTRIES = 100;
const MAX_MAX_ENTRIES = 500;
const DEFAULT_TTL_MS = 24 * 60 * 60 * 1000;
const MIN_TTL_MS = 60 * 1000;
const MAX_TTL_MS = 7 * 24 * 60 * 60 * 1000;

function text(value) {
  return String(value == null ? "" : value).trim();
}

function parseBoolean(value) {
  const raw = text(value).toLowerCase();
  if (!raw) return Object.freeze({ valid: true, value: false });
  if (["1", "true", "yes", "on"].includes(raw)) {
    return Object.freeze({ valid: true, value: true });
  }
  if (["0", "false", "no", "off"].includes(raw)) {
    return Object.freeze({ valid: true, value: false });
  }
  return Object.freeze({ valid: false, value: false });
}

function boundedInteger(value, fallback, min, max) {
  const parsed = Number.parseInt(text(value), 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(max, Math.max(min, parsed));
}

function constantTimeEqual(leftValue, rightValue) {
  const left = Buffer.from(String(leftValue || ""), "utf8");
  const right = Buffer.from(String(rightValue || ""), "utf8");
  if (!left.length || left.length !== right.length) return false;
  return crypto.timingSafeEqual(left, right);
}

function createQueryCandidatePlannerInternalPreviewConfig({
  env = process.env,
} = {}) {
  const enabledSetting = parseBoolean(env[ENABLED_ENV]);
  const requested = enabledSetting.valid && enabledSetting.value;
  const token = text(env[TOKEN_ENV]);
  const tokenByteLength = Buffer.byteLength(token, "utf8");
  const tokenConfigured = tokenByteLength >= MIN_TOKEN_BYTES;

  let reason = "FEATURE_DISABLED";
  if (!enabledSetting.valid) reason = "INVALID_ENABLED_VALUE";
  else if (requested && !tokenConfigured) reason = "TOKEN_REQUIRED";
  else if (requested && tokenConfigured) reason = "INTERNAL_PREVIEW_READY";

  const enabled = requested && tokenConfigured;
  const maxEntries = boundedInteger(
    env[MAX_ENTRIES_ENV],
    DEFAULT_MAX_ENTRIES,
    1,
    MAX_MAX_ENTRIES,
  );
  const ttlMs = boundedInteger(
    env[TTL_MS_ENV],
    DEFAULT_TTL_MS,
    MIN_TTL_MS,
    MAX_TTL_MS,
  );

  function verifyToken(candidate) {
    if (!enabled) return false;
    return constantTimeEqual(token, text(candidate));
  }

  function publicSnapshot() {
    return Object.freeze({
      version: CONFIG_VERSION,
      enabled,
      requested,
      configured: tokenConfigured,
      reason,
      maxEntries,
      ttlMs,
      access: Object.freeze({
        authenticationRequired: true,
        internalTokenRequired: true,
        headerName: "x-beebee-internal-preview-token",
        queryTokenAccepted: false,
        tokenIncluded: false,
        tokenHashIncluded: false,
      }),
      guardrails: Object.freeze({
        readOnly: true,
        observationOnly: true,
        productionCandidateMerge: false,
        productionReadyAssignment: false,
        productionRouteChanged: false,
        candidateExecutionAvailable: false,
        candidateSelectionAvailable: false,
      }),
    });
  }

  return Object.freeze({
    version: CONFIG_VERSION,
    enabled,
    requested,
    configured: tokenConfigured,
    reason,
    maxEntries,
    ttlMs,
    verifyToken,
    publicSnapshot,
  });
}

let runtimeConfig = null;

function getQueryCandidatePlannerInternalPreviewConfig() {
  if (!runtimeConfig) {
    runtimeConfig = createQueryCandidatePlannerInternalPreviewConfig({
      env: process.env,
    });
  }
  return runtimeConfig;
}

function resetQueryCandidatePlannerInternalPreviewConfigForTests({
  config = null,
} = {}) {
  runtimeConfig = config;
  return runtimeConfig;
}

module.exports = Object.freeze({
  CONFIG_VERSION,
  ENABLED_ENV,
  TOKEN_ENV,
  MAX_ENTRIES_ENV,
  TTL_MS_ENV,
  MIN_TOKEN_BYTES,
  DEFAULT_MAX_ENTRIES,
  DEFAULT_TTL_MS,
  parseBoolean,
  constantTimeEqual,
  createQueryCandidatePlannerInternalPreviewConfig,
  getQueryCandidatePlannerInternalPreviewConfig,
  resetQueryCandidatePlannerInternalPreviewConfigForTests,
});
