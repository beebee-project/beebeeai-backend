"use strict";

const crypto = require("crypto");
const {
  OPERATIONS,
} = require("./queryCandidatePlannerFeatureControl");
const {
  ADAPTER_VERSION,
  PROMOTION_GATE_DECISION_VERSION,
} = require("./queryCandidatePlannerControlledProductionMergeAdapter");

const GATE_VERSION =
  "query_candidate_planner_controlled_production_promotion_gate_v1";
const POLICY_VERSION =
  "query_candidate_planner_controlled_production_promotion_policy_v1";
const CONFIG_VERSION =
  "query_candidate_planner_controlled_production_promotion_config_v1";
const SNAPSHOT_VERSION =
  "query_candidate_planner_controlled_production_promotion_snapshot_v1";
const ROLLOUT_ALGORITHM = "SHA256_MOD_10000_V1";

const AUDIENCE_MODES = Object.freeze({
  BLOCKED: "BLOCKED",
  ALLOWLIST: "ALLOWLIST",
  ROLLOUT: "ROLLOUT",
});

const ENV_KEYS = Object.freeze({
  enabled: "QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED",
  audienceMode: "QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE",
  allowlistSha256:
    "QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256",
  rolloutPercent:
    "QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT",
  rolloutSalt: "QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_SALT",
});

const DEFAULTS = Object.freeze({
  enabled: false,
  audienceMode: AUDIENCE_MODES.BLOCKED,
  allowlistSha256: Object.freeze([]),
  rolloutPercent: 0,
  rolloutSalt: "",
});

const TRUE_VALUES = new Set(["1", "true"]);
const FALSE_VALUES = new Set(["0", "false"]);
const SHA256_RE = /^[a-f0-9]{64}$/i;
const MIN_ROLLOUT_SALT_LENGTH = 16;
const MAX_ROLLOUT_SALT_LENGTH = 200;

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function sha256(value) {
  return crypto.createHash("sha256").update(String(value)).digest("hex");
}

function normalizeSha256(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return SHA256_RE.test(normalized) ? normalized : "";
}

function parseStrictBoolean(env, key, fallback) {
  const present = Object.prototype.hasOwnProperty.call(env || {}, key);
  if (!present || env[key] === undefined || env[key] === null) {
    return Object.freeze({ value: fallback, valid: true, source: "DEFAULT" });
  }
  const normalized = String(env[key]).trim().toLowerCase();
  if (TRUE_VALUES.has(normalized)) {
    return Object.freeze({ value: true, valid: true, source: "ENV" });
  }
  if (FALSE_VALUES.has(normalized)) {
    return Object.freeze({ value: false, valid: true, source: "ENV" });
  }
  return Object.freeze({
    value: fallback,
    valid: false,
    source: "INVALID_ENV_FAIL_CLOSED",
  });
}

function parseAudienceMode(env, key, fallback) {
  const present = Object.prototype.hasOwnProperty.call(env || {}, key);
  if (!present || env[key] === undefined || env[key] === null) {
    return Object.freeze({ value: fallback, valid: true, source: "DEFAULT" });
  }
  const normalized = String(env[key]).trim().toUpperCase();
  if (Object.values(AUDIENCE_MODES).includes(normalized)) {
    return Object.freeze({ value: normalized, valid: true, source: "ENV" });
  }
  return Object.freeze({
    value: fallback,
    valid: false,
    source: "INVALID_ENV_FAIL_CLOSED",
  });
}

function parseAllowlist(env, key) {
  const present = Object.prototype.hasOwnProperty.call(env || {}, key);
  if (!present || env[key] === undefined || env[key] === null) {
    return Object.freeze({
      value: Object.freeze([]),
      valid: true,
      source: "DEFAULT",
      invalidEntries: Object.freeze([]),
    });
  }

  const raw = String(env[key]);
  const entries = raw
    .split(/[\s,;]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
  const normalized = [];
  const invalidEntries = [];
  const seen = new Set();

  for (const entry of entries) {
    const hash = normalizeSha256(entry);
    if (!hash) {
      invalidEntries.push(sha256(`invalid:${entry}`).slice(0, 16));
      continue;
    }
    if (seen.has(hash)) continue;
    seen.add(hash);
    normalized.push(hash);
  }

  normalized.sort();
  return Object.freeze({
    value: Object.freeze(normalized),
    valid: invalidEntries.length === 0,
    source:
      invalidEntries.length === 0 ? "ENV" : "INVALID_ENV_FAIL_CLOSED",
    invalidEntries: Object.freeze(invalidEntries),
  });
}

function parseRolloutPercent(env, key, fallback) {
  const present = Object.prototype.hasOwnProperty.call(env || {}, key);
  if (!present || env[key] === undefined || env[key] === null) {
    return Object.freeze({ value: fallback, valid: true, source: "DEFAULT" });
  }
  const raw = String(env[key]).trim();
  if (!/^\d{1,3}$/.test(raw)) {
    return Object.freeze({
      value: fallback,
      valid: false,
      source: "INVALID_ENV_FAIL_CLOSED",
    });
  }
  const value = Number(raw);
  if (!Number.isInteger(value) || value < 0 || value > 100) {
    return Object.freeze({
      value: fallback,
      valid: false,
      source: "INVALID_ENV_FAIL_CLOSED",
    });
  }
  return Object.freeze({ value, valid: true, source: "ENV" });
}

function parseRolloutSalt(env, key) {
  const present = Object.prototype.hasOwnProperty.call(env || {}, key);
  if (!present || env[key] === undefined || env[key] === null) {
    return Object.freeze({ value: "", valid: true, source: "DEFAULT" });
  }
  const value = String(env[key]).trim();
  const valid =
    value.length >= MIN_ROLLOUT_SALT_LENGTH &&
    value.length <= MAX_ROLLOUT_SALT_LENGTH &&
    !/[\r\n\t]/.test(value);
  return Object.freeze({
    value: valid ? value : "",
    valid,
    source: valid ? "ENV" : "INVALID_ENV_FAIL_CLOSED",
  });
}

function parsePromotionGateEnvironment(env = process.env) {
  const enabled = parseStrictBoolean(env, ENV_KEYS.enabled, DEFAULTS.enabled);
  const audienceMode = parseAudienceMode(
    env,
    ENV_KEYS.audienceMode,
    DEFAULTS.audienceMode,
  );
  const allowlist = parseAllowlist(env, ENV_KEYS.allowlistSha256);
  const rolloutPercent = parseRolloutPercent(
    env,
    ENV_KEYS.rolloutPercent,
    DEFAULTS.rolloutPercent,
  );
  const rolloutSalt = parseRolloutSalt(env, ENV_KEYS.rolloutSalt);

  const invalidEnvironmentKeys = [];
  if (!enabled.valid) invalidEnvironmentKeys.push(ENV_KEYS.enabled);
  if (!audienceMode.valid) invalidEnvironmentKeys.push(ENV_KEYS.audienceMode);
  if (!allowlist.valid) invalidEnvironmentKeys.push(ENV_KEYS.allowlistSha256);
  if (!rolloutPercent.valid) {
    invalidEnvironmentKeys.push(ENV_KEYS.rolloutPercent);
  }
  if (!rolloutSalt.valid) invalidEnvironmentKeys.push(ENV_KEYS.rolloutSalt);

  if (
    audienceMode.value === AUDIENCE_MODES.ALLOWLIST &&
    allowlist.value.length === 0
  ) {
    invalidEnvironmentKeys.push(ENV_KEYS.allowlistSha256);
  }
  if (
    audienceMode.value === AUDIENCE_MODES.ROLLOUT &&
    rolloutPercent.value > 0 &&
    !rolloutSalt.value
  ) {
    invalidEnvironmentKeys.push(ENV_KEYS.rolloutSalt);
  }

  const uniqueInvalidKeys = Object.freeze([
    ...new Set(invalidEnvironmentKeys),
  ]);

  return Object.freeze({
    version: CONFIG_VERSION,
    enabled: enabled.value,
    audienceMode: audienceMode.value,
    allowlistSha256: allowlist.value,
    allowlistCount: allowlist.value.length,
    rolloutPercent: rolloutPercent.value,
    rolloutBasisPoints: rolloutPercent.value * 100,
    rolloutSalt: rolloutSalt.value,
    rolloutSaltSha256: rolloutSalt.value ? sha256(rolloutSalt.value) : "",
    configurationValid: uniqueInvalidKeys.length === 0,
    invalidEnvironmentKeys: uniqueInvalidKeys,
    sources: Object.freeze({
      enabled: enabled.source,
      audienceMode: audienceMode.source,
      allowlistSha256: allowlist.source,
      rolloutPercent: rolloutPercent.source,
      rolloutSalt: rolloutSalt.source,
    }),
    failClosed: true,
  });
}

function deterministicRolloutBucket({ subjectSha256, salt } = {}) {
  const subject = normalizeSha256(subjectSha256);
  const normalizedSalt = String(salt || "").trim();
  if (!subject || !normalizedSalt) return null;
  const digest = sha256(`${POLICY_VERSION}:${normalizedSalt}:${subject}`);
  return Number.parseInt(digest.slice(0, 8), 16) % 10000;
}

function safeSubjectTag(subjectSha256) {
  const normalized = normalizeSha256(subjectSha256);
  return normalized ? sha256(`promotion-subject:${normalized}`) : "";
}

function guardrails() {
  return Object.freeze({
    defaultBlocked: true,
    routeWired: false,
    controllerWired: false,
    primaryResponseAuthority: true,
    responsePayloadMutation: false,
    responseHeaderMutation: false,
    responseStatusMutation: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    rawIdentityAccepted: false,
    rawAllowlistExposed: false,
    rolloutSaltExposed: false,
    failClosed: true,
  });
}

function baseDecision({
  allowed,
  reason,
  operation,
  adapterVersion,
  config,
  featureDecision = null,
  subjectSha256 = "",
  allowlistMatched = false,
  rolloutBucket = null,
  rolloutSelected = false,
  audiencePath = "NONE",
}) {
  const safeConfig = config || parsePromotionGateEnvironment({});
  return Object.freeze({
    version: PROMOTION_GATE_DECISION_VERSION,
    gateVersion: GATE_VERSION,
    policyVersion: POLICY_VERSION,
    allowed: allowed === true,
    decision: allowed === true ? "ALLOW" : "BLOCK",
    operation,
    reason,
    failClosed: true,
    adapterVersion,
    configuration: Object.freeze({
      valid: safeConfig.configurationValid,
      enabled: safeConfig.enabled,
      audienceMode: safeConfig.audienceMode,
      allowlistCount: safeConfig.allowlistCount,
      rolloutPercent: safeConfig.rolloutPercent,
      invalidEnvironmentKeys: Object.freeze([
        ...safeConfig.invalidEnvironmentKeys,
      ]),
    }),
    featureDecision,
    audience: Object.freeze({
      path: audiencePath,
      subjectTagSha256: safeSubjectTag(subjectSha256),
      allowlistMatched: allowlistMatched === true,
      rollout: Object.freeze({
        algorithm: ROLLOUT_ALGORITHM,
        percentage: safeConfig.rolloutPercent,
        basisPoints: safeConfig.rolloutBasisPoints,
        bucket: rolloutBucket,
        selected: rolloutSelected === true,
      }),
    }),
    guardrails: guardrails(),
  });
}

function evaluateAudience({ config, subjectSha256 }) {
  const subject = normalizeSha256(subjectSha256);
  if (config.audienceMode === AUDIENCE_MODES.BLOCKED) {
    return Object.freeze({
      allowed: false,
      reason: "AUDIENCE_MODE_BLOCKED",
      audiencePath: "BLOCKED",
      subjectSha256: subject,
      allowlistMatched: false,
      rolloutBucket: null,
      rolloutSelected: false,
    });
  }
  if (!subject) {
    return Object.freeze({
      allowed: false,
      reason: "SUBJECT_SHA256_REQUIRED",
      audiencePath: config.audienceMode,
      subjectSha256: "",
      allowlistMatched: false,
      rolloutBucket: null,
      rolloutSelected: false,
    });
  }

  const allowlistMatched = config.allowlistSha256.includes(subject);
  if (config.audienceMode === AUDIENCE_MODES.ALLOWLIST) {
    return Object.freeze({
      allowed: allowlistMatched,
      reason: allowlistMatched
        ? "ALLOWLIST_MATCH"
        : "SUBJECT_NOT_ALLOWLISTED",
      audiencePath: "ALLOWLIST",
      subjectSha256: subject,
      allowlistMatched,
      rolloutBucket: null,
      rolloutSelected: false,
    });
  }

  if (allowlistMatched) {
    return Object.freeze({
      allowed: true,
      reason: "ALLOWLIST_MATCH_ROLLOUT_BYPASS",
      audiencePath: "ROLLOUT_ALLOWLIST",
      subjectSha256: subject,
      allowlistMatched: true,
      rolloutBucket: null,
      rolloutSelected: false,
    });
  }

  const rolloutBucket = deterministicRolloutBucket({
    subjectSha256: subject,
    salt: config.rolloutSalt,
  });
  const rolloutSelected =
    rolloutBucket !== null && rolloutBucket < config.rolloutBasisPoints;
  return Object.freeze({
    allowed: rolloutSelected,
    reason: rolloutSelected
      ? "DETERMINISTIC_ROLLOUT_SELECTED"
      : config.rolloutPercent === 0
        ? "ROLLOUT_PERCENT_ZERO"
        : "DETERMINISTIC_ROLLOUT_NOT_SELECTED",
    audiencePath: "ROLLOUT",
    subjectSha256: subject,
    allowlistMatched: false,
    rolloutBucket,
    rolloutSelected,
  });
}

function evaluateControlledProductionPromotionGate({
  env = process.env,
  featureControl = null,
  readinessGate = null,
  subjectSha256 = "",
  operation = OPERATIONS.PRODUCTION_CANDIDATE_MERGE,
  adapterVersion = ADAPTER_VERSION,
} = {}) {
  const config = parsePromotionGateEnvironment(env);

  if (operation !== OPERATIONS.PRODUCTION_CANDIDATE_MERGE) {
    return baseDecision({
      allowed: false,
      reason: "UNSUPPORTED_PROMOTION_OPERATION",
      operation,
      adapterVersion,
      config,
    });
  }
  if (adapterVersion !== ADAPTER_VERSION) {
    return baseDecision({
      allowed: false,
      reason: "ADAPTER_VERSION_MISMATCH",
      operation,
      adapterVersion,
      config,
    });
  }
  if (!config.configurationValid) {
    return baseDecision({
      allowed: false,
      reason: "INVALID_PROMOTION_GATE_CONFIGURATION",
      operation,
      adapterVersion,
      config,
    });
  }
  if (!config.enabled) {
    return baseDecision({
      allowed: false,
      reason: "PROMOTION_GATE_DISABLED",
      operation,
      adapterVersion,
      config,
    });
  }
  if (!featureControl || typeof featureControl.evaluate !== "function") {
    return baseDecision({
      allowed: false,
      reason: "FEATURE_CONTROL_REQUIRED",
      operation,
      adapterVersion,
      config,
    });
  }

  const featureDecision = featureControl.evaluate(operation, {
    readinessGate,
  });
  if (!featureDecision.allowed) {
    return baseDecision({
      allowed: false,
      reason: featureDecision.reason,
      operation,
      adapterVersion,
      config,
      featureDecision,
    });
  }

  const audience = evaluateAudience({ config, subjectSha256 });
  return baseDecision({
    allowed: audience.allowed,
    reason: audience.reason,
    operation,
    adapterVersion,
    config,
    featureDecision,
    subjectSha256: audience.subjectSha256,
    allowlistMatched: audience.allowlistMatched,
    rolloutBucket: audience.rolloutBucket,
    rolloutSelected: audience.rolloutSelected,
    audiencePath: audience.audiencePath,
  });
}

function createControlledProductionPromotionGate({
  env = process.env,
  featureControl = null,
} = {}) {
  const config = parsePromotionGateEnvironment(env);

  function snapshot() {
    return Object.freeze({
      version: SNAPSHOT_VERSION,
      gateVersion: GATE_VERSION,
      policyVersion: POLICY_VERSION,
      enabled: config.enabled,
      configurationValid: config.configurationValid,
      audienceMode: config.audienceMode,
      allowlistCount: config.allowlistCount,
      rolloutPercent: config.rolloutPercent,
      rolloutAlgorithm: ROLLOUT_ALGORITHM,
      defaultDecision: "BLOCK",
      routeWired: false,
      controllerWired: false,
      productionRouteChanged: false,
      productionReadyAssignment: false,
      failClosed: true,
    });
  }

  function evaluate(options = {}) {
    return evaluateControlledProductionPromotionGate({
      ...options,
      env,
      featureControl: options.featureControl || featureControl,
    });
  }

  return Object.freeze({
    version: GATE_VERSION,
    policyVersion: POLICY_VERSION,
    evaluate,
    snapshot,
  });
}

module.exports = Object.freeze({
  GATE_VERSION,
  POLICY_VERSION,
  CONFIG_VERSION,
  SNAPSHOT_VERSION,
  ROLLOUT_ALGORITHM,
  AUDIENCE_MODES,
  ENV_KEYS,
  DEFAULTS,
  MIN_ROLLOUT_SALT_LENGTH,
  MAX_ROLLOUT_SALT_LENGTH,
  normalizeSha256,
  parsePromotionGateEnvironment,
  deterministicRolloutBucket,
  evaluateAudience,
  evaluateControlledProductionPromotionGate,
  createControlledProductionPromotionGate,
});
