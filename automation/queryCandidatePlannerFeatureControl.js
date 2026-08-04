"use strict";

const CONTROL_VERSION = "query_candidate_planner_feature_control_v1";
const POLICY_VERSION = "candidate_planner_feature_flags_kill_switch_policy_v1";
const READINESS_DECISION =
  "ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW";

const OPERATIONS = Object.freeze({
  SHADOW_EXECUTION: "SHADOW_EXECUTION",
  PROVIDER_CALL: "PROVIDER_CALL",
  CACHE_READ: "CACHE_READ",
  CACHE_WRITE: "CACHE_WRITE",
  PRODUCTION_CANDIDATE_MERGE: "PRODUCTION_CANDIDATE_MERGE",
  PRODUCTION_READY_ASSIGNMENT: "PRODUCTION_READY_ASSIGNMENT",
  PRODUCTION_ROUTE: "PRODUCTION_ROUTE",
});

const SCOPES = Object.freeze({
  GLOBAL: "GLOBAL",
  PROVIDER: "PROVIDER",
  CACHE: "CACHE",
  PRODUCTION: "PRODUCTION",
});

const ENV_KEYS = Object.freeze({
  featureEnabled: "QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED",
  shadowEnabled: "QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED",
  providerEnabled: "QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED",
  cacheReadEnabled: "QUERY_CANDIDATE_PLANNER_CACHE_READ_ENABLED",
  cacheWriteEnabled: "QUERY_CANDIDATE_PLANNER_CACHE_WRITE_ENABLED",
  productionEnabled: "QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED",
  productionCandidateMergeEnabled:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED",
  productionReadyAssignmentEnabled:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED",
  productionRouteEnabled:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED",
  globalKillSwitch: "QUERY_CANDIDATE_PLANNER_KILL_SWITCH",
  providerKillSwitch: "QUERY_CANDIDATE_PLANNER_PROVIDER_KILL_SWITCH",
  cacheKillSwitch: "QUERY_CANDIDATE_PLANNER_CACHE_KILL_SWITCH",
  productionKillSwitch:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH",
});

const DEFAULTS = Object.freeze({
  featureEnabled: false,
  shadowEnabled: true,
  providerEnabled: false,
  cacheReadEnabled: true,
  cacheWriteEnabled: true,
  productionEnabled: false,
  productionCandidateMergeEnabled: false,
  productionReadyAssignmentEnabled: false,
  productionRouteEnabled: false,
  globalKillSwitch: false,
  providerKillSwitch: false,
  cacheKillSwitch: false,
  productionKillSwitch: true,
});

const BOOLEAN_TRUE = new Set(["1", "true"]);
const BOOLEAN_FALSE = new Set(["0", "false"]);

function parseStrictBoolean(env, envKey, fallback) {
  const present = Object.prototype.hasOwnProperty.call(env || {}, envKey);
  if (!present || env[envKey] === undefined || env[envKey] === null) {
    return { value: fallback, valid: true, source: "DEFAULT" };
  }

  const normalized = String(env[envKey]).trim().toLowerCase();
  if (BOOLEAN_TRUE.has(normalized)) {
    return { value: true, valid: true, source: "ENV" };
  }
  if (BOOLEAN_FALSE.has(normalized)) {
    return { value: false, valid: true, source: "ENV" };
  }

  return {
    value: fallback,
    valid: false,
    source: "INVALID_ENV_FAIL_CLOSED",
  };
}

function parseEnvironment(env = process.env) {
  const parsed = {};
  const invalidEnvironmentKeys = [];

  for (const [name, envKey] of Object.entries(ENV_KEYS)) {
    parsed[name] = parseStrictBoolean(env, envKey, DEFAULTS[name]);
    if (!parsed[name].valid) invalidEnvironmentKeys.push(envKey);
  }

  const flags = Object.freeze({
    featureEnabled: parsed.featureEnabled.value,
    shadowEnabled: parsed.shadowEnabled.value,
    providerEnabled: parsed.providerEnabled.value,
    cacheReadEnabled: parsed.cacheReadEnabled.value,
    cacheWriteEnabled: parsed.cacheWriteEnabled.value,
    productionEnabled: parsed.productionEnabled.value,
    productionCandidateMergeEnabled:
      parsed.productionCandidateMergeEnabled.value,
    productionReadyAssignmentEnabled:
      parsed.productionReadyAssignmentEnabled.value,
    productionRouteEnabled: parsed.productionRouteEnabled.value,
  });

  const killSwitches = Object.freeze({
    global: parsed.globalKillSwitch.value,
    provider: parsed.providerKillSwitch.value,
    cache: parsed.cacheKillSwitch.value,
    production: parsed.productionKillSwitch.value,
  });

  const sources = {};
  for (const [name, result] of Object.entries(parsed)) {
    sources[name] = result.source;
  }

  return Object.freeze({
    flags,
    killSwitches,
    configurationValid: invalidEnvironmentKeys.length === 0,
    invalidEnvironmentKeys: Object.freeze([...invalidEnvironmentKeys]),
    sources: Object.freeze(sources),
  });
}

function isProductionOperation(operation) {
  return (
    operation === OPERATIONS.PRODUCTION_CANDIDATE_MERGE ||
    operation === OPERATIONS.PRODUCTION_READY_ASSIGNMENT ||
    operation === OPERATIONS.PRODUCTION_ROUTE
  );
}

function evaluateReadinessGate(readinessGate) {
  const guardrails = readinessGate?.guardrails || {};
  const checks = {
    objectPresent: Boolean(readinessGate && typeof readinessGate === "object"),
    eligible: readinessGate?.eligible === true,
    decision: readinessGate?.decision === READINESS_DECISION,
    manualReviewRequired: guardrails.manualPromotionReviewRequired === true,
    failClosed: guardrails.failClosed === true,
    routeNotAutoWired: guardrails.productionRouteAutoWired === false,
    mergeNotAutoAllowed: guardrails.productionCandidateMergeAllowed === false,
    readyNotAutoAllowed:
      guardrails.productionReadyAssignmentAllowed === false,
  };

  const valid = Object.values(checks).every(Boolean);
  return Object.freeze({
    valid,
    checks: Object.freeze(checks),
    reason: valid ? "PATCH13_3_READINESS_EVIDENCE_VALID" : "READINESS_EVIDENCE_INVALID",
  });
}

function normalizeScope(scope) {
  const value = String(scope || "").trim().toUpperCase();
  if (!Object.values(SCOPES).includes(value)) {
    throw new Error(`Unsupported kill-switch scope: ${scope}`);
  }
  return value;
}

function safeAuditText(value, fallback) {
  const normalized = String(value || "").trim();
  if (!normalized) return fallback;
  return normalized.slice(0, 120).replace(/[\r\n\t]/g, " ");
}

function createQueryCandidatePlannerFeatureControl({
  env = process.env,
  now = () => new Date(),
  maxAuditEvents = 100,
} = {}) {
  const boot = parseEnvironment(env);
  const runtime = {
    revision: 0,
    killSwitches: {
      global: false,
      provider: false,
      cache: false,
      production: false,
    },
    auditEvents: [],
  };

  function timestamp() {
    const value = now();
    return value instanceof Date ? value.toISOString() : new Date(value).toISOString();
  }

  function record(event) {
    runtime.auditEvents.push(
      Object.freeze({
        version: "query_candidate_planner_feature_control_audit_event_v1",
        at: timestamp(),
        revision: runtime.revision,
        ...event,
      }),
    );
    if (runtime.auditEvents.length > maxAuditEvents) {
      runtime.auditEvents.splice(0, runtime.auditEvents.length - maxAuditEvents);
    }
  }

  function effectiveKillSwitches() {
    return Object.freeze({
      global: boot.killSwitches.global || runtime.killSwitches.global,
      provider: boot.killSwitches.provider || runtime.killSwitches.provider,
      cache: boot.killSwitches.cache || runtime.killSwitches.cache,
      production:
        boot.killSwitches.production || runtime.killSwitches.production,
    });
  }

  function snapshot() {
    return Object.freeze({
      version: CONTROL_VERSION,
      policyVersion: POLICY_VERSION,
      configurationValid: boot.configurationValid,
      invalidEnvironmentKeys: Object.freeze([...boot.invalidEnvironmentKeys]),
      flags: boot.flags,
      killSwitches: effectiveKillSwitches(),
      killSwitchSources: Object.freeze({
        global: Object.freeze({
          environment: boot.killSwitches.global,
          runtime: runtime.killSwitches.global,
        }),
        provider: Object.freeze({
          environment: boot.killSwitches.provider,
          runtime: runtime.killSwitches.provider,
        }),
        cache: Object.freeze({
          environment: boot.killSwitches.cache,
          runtime: runtime.killSwitches.cache,
        }),
        production: Object.freeze({
          environment: boot.killSwitches.production,
          runtime: runtime.killSwitches.production,
        }),
      }),
      runtimeRevision: runtime.revision,
      failClosed: true,
      productionRouteChanged: false,
      productionCandidateMerge: false,
      productionReadyAssignment: false,
    });
  }

  function deny(operation, reason, extra = {}) {
    return Object.freeze({
      version: "query_candidate_planner_feature_control_decision_v1",
      operation,
      allowed: false,
      decision: "DENY",
      reason,
      failClosed: true,
      runtimeRevision: runtime.revision,
      ...extra,
    });
  }

  function allow(operation, extra = {}) {
    return Object.freeze({
      version: "query_candidate_planner_feature_control_decision_v1",
      operation,
      allowed: true,
      decision: "ALLOW",
      reason: "FEATURE_CONTROL_ALLOW",
      failClosed: true,
      runtimeRevision: runtime.revision,
      ...extra,
    });
  }

  function evaluate(operation, { readinessGate = null } = {}) {
    if (!Object.values(OPERATIONS).includes(operation)) {
      return deny(operation, "UNKNOWN_OPERATION");
    }

    if (!boot.configurationValid) {
      return deny(operation, "INVALID_ENVIRONMENT_CONFIGURATION", {
        invalidEnvironmentKeys: Object.freeze([
          ...boot.invalidEnvironmentKeys,
        ]),
      });
    }

    const kills = effectiveKillSwitches();
    if (kills.global) return deny(operation, "GLOBAL_KILL_SWITCH_ACTIVE");
    if (!boot.flags.featureEnabled) {
      return deny(operation, "FEATURE_DISABLED");
    }

    if (operation === OPERATIONS.SHADOW_EXECUTION) {
      if (!boot.flags.shadowEnabled) return deny(operation, "SHADOW_DISABLED");
      return allow(operation);
    }

    if (operation === OPERATIONS.PROVIDER_CALL) {
      if (kills.provider) return deny(operation, "PROVIDER_KILL_SWITCH_ACTIVE");
      if (!boot.flags.providerEnabled) return deny(operation, "PROVIDER_DISABLED");
      return allow(operation);
    }

    if (operation === OPERATIONS.CACHE_READ || operation === OPERATIONS.CACHE_WRITE) {
      if (kills.cache) return deny(operation, "CACHE_KILL_SWITCH_ACTIVE");
      const enabled =
        operation === OPERATIONS.CACHE_READ
          ? boot.flags.cacheReadEnabled
          : boot.flags.cacheWriteEnabled;
      if (!enabled) {
        return deny(
          operation,
          operation === OPERATIONS.CACHE_READ
            ? "CACHE_READ_DISABLED"
            : "CACHE_WRITE_DISABLED",
        );
      }
      return allow(operation);
    }

    if (isProductionOperation(operation)) {
      if (kills.production) {
        return deny(operation, "PRODUCTION_KILL_SWITCH_ACTIVE");
      }
      if (!boot.flags.productionEnabled) {
        return deny(operation, "PRODUCTION_FEATURE_DISABLED");
      }

      const operationFlag = {
        [OPERATIONS.PRODUCTION_CANDIDATE_MERGE]:
          boot.flags.productionCandidateMergeEnabled,
        [OPERATIONS.PRODUCTION_READY_ASSIGNMENT]:
          boot.flags.productionReadyAssignmentEnabled,
        [OPERATIONS.PRODUCTION_ROUTE]: boot.flags.productionRouteEnabled,
      }[operation];

      if (!operationFlag) return deny(operation, "PRODUCTION_OPERATION_DISABLED");

      const readiness = evaluateReadinessGate(readinessGate);
      if (!readiness.valid) {
        return deny(operation, readiness.reason, { readiness });
      }
      return allow(operation, { readiness });
    }

    return deny(operation, "UNREACHABLE_FAIL_CLOSED");
  }

  function activateKillSwitch({
    scope = SCOPES.GLOBAL,
    reason = "MANUAL_EMERGENCY_STOP",
    actor = "SYSTEM",
  } = {}) {
    const normalizedScope = normalizeScope(scope);
    const key = normalizedScope.toLowerCase();
    runtime.killSwitches[key] = true;
    runtime.revision += 1;
    record({
      action: "ACTIVATE",
      scope: normalizedScope,
      reason: safeAuditText(reason, "MANUAL_EMERGENCY_STOP"),
      actor: safeAuditText(actor, "SYSTEM"),
    });
    return snapshot();
  }

  function releaseRuntimeKillSwitch({
    scope = SCOPES.GLOBAL,
    reason = "INCIDENT_RESOLVED",
    actor = "SYSTEM",
  } = {}) {
    const normalizedScope = normalizeScope(scope);
    const key = normalizedScope.toLowerCase();
    runtime.killSwitches[key] = false;
    runtime.revision += 1;
    record({
      action: "RELEASE_RUNTIME_ONLY",
      scope: normalizedScope,
      reason: safeAuditText(reason, "INCIDENT_RESOLVED"),
      actor: safeAuditText(actor, "SYSTEM"),
      environmentKillSwitchStillAuthoritative: boot.killSwitches[key] === true,
    });
    return snapshot();
  }

  function getAuditEvents() {
    return Object.freeze([...runtime.auditEvents]);
  }

  return Object.freeze({
    version: CONTROL_VERSION,
    policyVersion: POLICY_VERSION,
    evaluate,
    snapshot,
    activateKillSwitch,
    releaseRuntimeKillSwitch,
    getAuditEvents,
  });
}

module.exports = Object.freeze({
  CONTROL_VERSION,
  POLICY_VERSION,
  READINESS_DECISION,
  OPERATIONS,
  SCOPES,
  ENV_KEYS,
  DEFAULTS,
  parseEnvironment,
  evaluateReadinessGate,
  createQueryCandidatePlannerFeatureControl,
});
