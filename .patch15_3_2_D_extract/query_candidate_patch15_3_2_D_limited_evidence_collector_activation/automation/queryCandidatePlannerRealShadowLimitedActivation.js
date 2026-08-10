"use strict";

const {
  evaluateRealShadowSecureRuntime,
} = require("./queryCandidatePlannerRealShadowSecureDeployment");
const {
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
} = require("./queryCandidatePlannerRealShadowEvidenceConfig");

const LIMITED_ACTIVATION_VERSION =
  "query_candidate_planner_real_shadow_limited_activation_v1";

function text(value, maxLength = 20000) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function booleanValue(value, fallback = false) {
  const normalized = text(value, 20).toLowerCase();
  if (!normalized) return fallback;
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off"].includes(normalized)) return false;
  return fallback;
}

function integerValue(value, fallback) {
  const n = Number(value);
  return Number.isInteger(n) ? n : fallback;
}

function baselineSecureRuntime(env) {
  return evaluateRealShadowSecureRuntime({
    env: {
      ...env,
      QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED: "0",
    },
  });
}

function commonLimitedErrors(env) {
  const errors = [];
  const baseline = baselineSecureRuntime(env);
  if (!baseline.ready) {
    for (const error of baseline.errors || [baseline.reason]) {
      errors.push(`PATCH_15_3_2_C_BASELINE:${error}`);
    }
  }

  const config = parseQueryCandidatePlannerRealShadowEvidenceConfig(env);
  if (!config.configurationValid) errors.push(config.reason);
  if (config.allowlist.length !== 1) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_ALLOWLIST_MUST_HAVE_EXACTLY_ONE_ENTRY");
  }

  const registryCases = Array.isArray(config.registry?.registry?.cases)
    ? config.registry.registry.cases
    : [];
  const uniqueCaseIds = new Set(registryCases.map((item) => item.caseId));
  if (!config.registry.valid || uniqueCaseIds.size !== 10) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_REQUIRES_10_CASE_REGISTRY");
  }

  const ttlDays = integerValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_TTL_DAYS,
    config.ttlDays,
  );
  const maxRecords = integerValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_MAX_RECORDS,
    config.maxRecords,
  );
  if (ttlDays < 1 || ttlDays > 7) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_TTL_MUST_BE_1_TO_7_DAYS");
  }
  if (maxRecords < 30 || maxRecords > 5000) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_MAX_RECORDS_MUST_BE_30_TO_5000");
  }

  return Object.freeze({
    errors,
    baseline,
    config,
    ttlDays,
    maxRecords,
    registryCaseCount: uniqueCaseIds.size,
  });
}

function evaluateRealShadowLimitedActivationPreflight({ env = process.env } = {}) {
  const state = commonLimitedErrors(env);
  const errors = [...state.errors];
  const requestedEnabled = booleanValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED,
    false,
  );
  const killSwitch = booleanValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH,
    true,
  );

  if (requestedEnabled) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_PREFLIGHT_REQUIRES_COLLECTOR_DISABLED");
  }
  if (!killSwitch) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_PREFLIGHT_REQUIRES_KILL_SWITCH_ACTIVE");
  }

  const uniqueErrors = Object.freeze([...new Set(errors)]);
  const ready = uniqueErrors.length === 0;
  return Object.freeze({
    version: LIMITED_ACTIVATION_VERSION,
    phase: "15.3-B",
    patch: "15.3.2-D",
    stage: "PRE_ACTIVATION",
    ready,
    reason: ready
      ? "REAL_SHADOW_LIMITED_ACTIVATION_PREFLIGHT_PASS"
      : uniqueErrors[0] || "REAL_SHADOW_LIMITED_ACTIVATION_PREFLIGHT_BLOCKED",
    errors: uniqueErrors,
    registryCaseCount: state.registryCaseCount,
    allowlistEntryCount: state.config.allowlist.length,
    ttlDays: state.ttlDays,
    maxRecords: state.maxRecords,
    collectorRequestedEnabled: requestedEnabled,
    collectorEffectiveEnabled: state.config.enabled,
    collectorKillSwitchActive: killSwitch,
    productionSafety: state.baseline.productionSafety,
    readyForActivation: ready,
    internalCanaryEnabledByThisOperation: false,
    productionPromotionAuthorized: false,
    rawSecretIncluded: false,
  });
}

function evaluateRealShadowLimitedActivationRuntime({ env = process.env } = {}) {
  const state = commonLimitedErrors(env);
  const errors = [...state.errors];
  const requestedEnabled = booleanValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED,
    false,
  );
  const killSwitch = booleanValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH,
    true,
  );

  if (!requestedEnabled) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_COLLECTOR_MUST_BE_ENABLED");
  }
  if (killSwitch) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_KILL_SWITCH_MUST_BE_RELEASED");
  }
  if (!state.config.enabled) {
    errors.push("REAL_SHADOW_LIMITED_ACTIVATION_COLLECTOR_NOT_EFFECTIVELY_ENABLED");
  }

  const uniqueErrors = Object.freeze([...new Set(errors)]);
  const ready = uniqueErrors.length === 0;
  return Object.freeze({
    version: LIMITED_ACTIVATION_VERSION,
    phase: "15.3-B",
    patch: "15.3.2-D",
    stage: "LIMITED_ACTIVE",
    ready,
    reason: ready
      ? "REAL_SHADOW_LIMITED_COLLECTOR_ACTIVE_READY_FOR_OBSERVATION_COLLECTION"
      : uniqueErrors[0] || "REAL_SHADOW_LIMITED_ACTIVATION_BLOCKED",
    errors: uniqueErrors,
    registryCaseCount: state.registryCaseCount,
    allowlistEntryCount: state.config.allowlist.length,
    ttlDays: state.ttlDays,
    maxRecords: state.maxRecords,
    collectorRequestedEnabled: requestedEnabled,
    collectorEffectiveEnabled: state.config.enabled,
    collectorKillSwitchActive: killSwitch,
    productionSafety: state.baseline.productionSafety,
    readyForPatch15_3_2_E: ready,
    internalCanaryEnabledByThisOperation: false,
    productionPromotionAuthorized: false,
    rawSecretIncluded: false,
  });
}

module.exports = Object.freeze({
  LIMITED_ACTIVATION_VERSION,
  evaluateRealShadowLimitedActivationPreflight,
  evaluateRealShadowLimitedActivationRuntime,
});
