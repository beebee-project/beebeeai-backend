"use strict";

const CONFIG_VERSION =
  "query_candidate_planner_internal_allowlist_canary_config_v1";

const ENV_KEYS = Object.freeze({
  enabled: "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED",
  killSwitch: "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH",
  evidenceJson: "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_EVIDENCE_JSON",
  timeoutMs: "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_TIMEOUT_MS",
  llmMode: "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_LLM_MODE",
});

const LLM_MODES = Object.freeze({
  SEMANTIC_PROFILER_ONLY: "SEMANTIC_PROFILER_ONLY",
});

const DEFAULTS = Object.freeze({
  enabled: false,
  killSwitch: true,
  evidenceJson: "",
  timeoutMs: 15000,
  llmMode: LLM_MODES.SEMANTIC_PROFILER_ONLY,
});

const TRUE_VALUES = new Set(["1", "true"]);
const FALSE_VALUES = new Set(["0", "false"]);

function strictBoolean(env, key, fallback) {
  if (!Object.prototype.hasOwnProperty.call(env || {}, key)) {
    return { value: fallback, valid: true, source: "DEFAULT" };
  }
  const normalized = String(env[key]).trim().toLowerCase();
  if (TRUE_VALUES.has(normalized)) {
    return { value: true, valid: true, source: "ENV" };
  }
  if (FALSE_VALUES.has(normalized)) {
    return { value: false, valid: true, source: "ENV" };
  }
  return {
    value: fallback,
    valid: false,
    source: "INVALID_ENV_FAIL_CLOSED",
  };
}

function timeoutValue(env, key, fallback) {
  if (!Object.prototype.hasOwnProperty.call(env || {}, key)) {
    return { value: fallback, valid: true, source: "DEFAULT" };
  }
  const raw = String(env[key]).trim();
  const value = Number(raw);
  const valid = /^\d+$/.test(raw) && Number.isInteger(value) &&
    value >= 1000 && value <= 60000;
  return {
    value: valid ? value : fallback,
    valid,
    source: valid ? "ENV" : "INVALID_ENV_FAIL_CLOSED",
  };
}

function llmModeValue(env, key, fallback) {
  if (!Object.prototype.hasOwnProperty.call(env || {}, key)) {
    return { value: fallback, valid: true, source: "DEFAULT" };
  }
  const value = String(env[key]).trim().toUpperCase();
  const valid = Object.values(LLM_MODES).includes(value);
  return {
    value: valid ? value : fallback,
    valid,
    source: valid ? "ENV" : "INVALID_ENV_FAIL_CLOSED",
  };
}

function parseQueryCandidatePlannerInternalCanaryConfig(
  env = process.env,
) {
  const enabled = strictBoolean(
    env,
    ENV_KEYS.enabled,
    DEFAULTS.enabled,
  );
  const killSwitch = strictBoolean(
    env,
    ENV_KEYS.killSwitch,
    DEFAULTS.killSwitch,
  );
  const timeoutMs = timeoutValue(
    env,
    ENV_KEYS.timeoutMs,
    DEFAULTS.timeoutMs,
  );
  const llmMode = llmModeValue(
    env,
    ENV_KEYS.llmMode,
    DEFAULTS.llmMode,
  );
  const evidenceJson = Object.prototype.hasOwnProperty.call(
    env || {},
    ENV_KEYS.evidenceJson,
  )
    ? String(env[ENV_KEYS.evidenceJson] || "").trim()
    : DEFAULTS.evidenceJson;

  const invalidEnvironmentKeys = [];
  if (!enabled.valid) invalidEnvironmentKeys.push(ENV_KEYS.enabled);
  if (!killSwitch.valid) invalidEnvironmentKeys.push(ENV_KEYS.killSwitch);
  if (!timeoutMs.valid) invalidEnvironmentKeys.push(ENV_KEYS.timeoutMs);
  if (!llmMode.valid) invalidEnvironmentKeys.push(ENV_KEYS.llmMode);

  return Object.freeze({
    version: CONFIG_VERSION,
    enabled: enabled.value,
    killSwitch: killSwitch.value,
    timeoutMs: timeoutMs.value,
    llmMode: llmMode.value,
    plannerEscalationAllowed: false,
    evidenceJson,
    evidenceConfigured: Boolean(evidenceJson),
    configurationValid: invalidEnvironmentKeys.length === 0,
    invalidEnvironmentKeys: Object.freeze(invalidEnvironmentKeys),
    failClosed: true,
    defaults: Object.freeze({
      enabled: false,
      killSwitch: true,
      audience: "ALLOWLIST_ONLY",
      rolloutPercent: 0,
      llmMode: LLM_MODES.SEMANTIC_PROFILER_ONLY,
    }),
  });
}

module.exports = Object.freeze({
  CONFIG_VERSION,
  ENV_KEYS,
  LLM_MODES,
  DEFAULTS,
  parseQueryCandidatePlannerInternalCanaryConfig,
});
