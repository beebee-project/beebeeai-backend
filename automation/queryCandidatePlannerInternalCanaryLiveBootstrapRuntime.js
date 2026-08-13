"use strict";

const {
  evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapGate,
  ENV: BOOTSTRAP_ENV,
  parseStrictBoolean,
} = require("./queryCandidatePlannerInternalCanaryLiveBootstrapGate");
const {
  deriveQueryCandidatePlannerInternalCanarySubject,
} = require("./queryCandidatePlannerInternalCanarySubject");
const {
  parseQueryCandidatePlannerInternalCanaryConfig,
  LLM_MODES,
} = require("./queryCandidatePlannerInternalAllowlistCanaryConfig");
const {
  runQueryCandidatePlannerInternalAllowlistCanary,
} = require("./queryCandidatePlannerInternalAllowlistCanaryService");

const RUNTIME_VERSION =
  "query_candidate_planner_internal_canary_live_bootstrap_runtime_v1";
const OBSERVATION_VERSION =
  "query_candidate_planner_internal_canary_live_bootstrap_observation_v1";
const RESULT_VERSION =
  "query_candidate_planner_internal_canary_live_bootstrap_result_v1";
const OBSERVE_ONLY_MERGE_ADAPTER_VERSION =
  "query_candidate_planner_internal_canary_live_bootstrap_no_merge_adapter_v1";

function hasOwn(object, key) {
  return Object.prototype.hasOwnProperty.call(object || {}, key);
}

function parseQueryCandidatePlannerInternalCanaryLiveBootstrapRuntimeMode(
  env = process.env,
) {
  if (!hasOwn(env, BOOTSTRAP_ENV.enabled)) {
    return Object.freeze({
      version: RUNTIME_VERSION,
      active: false,
      valid: true,
      enabled: false,
      source: "DEFAULT_DISABLED",
      reason: "LIVE_BOOTSTRAP_RUNTIME_NOT_REQUESTED",
      failClosed: true,
    });
  }

  const parsed = parseStrictBoolean(env[BOOTSTRAP_ENV.enabled]);
  if (!parsed.valid) {
    return Object.freeze({
      version: RUNTIME_VERSION,
      active: true,
      valid: false,
      enabled: false,
      source: "INVALID_ENV_FAIL_CLOSED",
      reason: "LIVE_BOOTSTRAP_RUNTIME_ENABLED_VALUE_INVALID",
      failClosed: true,
    });
  }

  if (!parsed.value) {
    return Object.freeze({
      version: RUNTIME_VERSION,
      active: false,
      valid: true,
      enabled: false,
      source: "EXPLICIT_DISABLED",
      reason: "LIVE_BOOTSTRAP_RUNTIME_DISABLED",
      failClosed: true,
    });
  }

  return Object.freeze({
    version: RUNTIME_VERSION,
    active: true,
    valid: true,
    enabled: true,
    source: "EXPLICIT_ENABLED",
    reason: "LIVE_BOOTSTRAP_RUNTIME_REQUESTED",
    failClosed: true,
  });
}

function safeSubject(subject = {}) {
  return Object.freeze({
    complete: subject?.complete === true,
    reason: String(subject?.reason || ""),
    subjectSha256: String(subject?.subjectSha256 || ""),
    subjectTagSha256: String(subject?.subjectTagSha256 || ""),
    rawIdentityIncluded: false,
  });
}

function blockedAuthorization(
  reason,
  {
    mode = null,
    subject = {},
    legacyPreflight = null,
    gate = null,
  } = {},
) {
  return Object.freeze({
    version: RUNTIME_VERSION,
    active: mode?.active === true,
    allowed: false,
    decision: "BLOCK",
    reason: String(reason || "LIVE_BOOTSTRAP_RUNTIME_BLOCKED"),
    failClosed: true,
    mode,
    subject: safeSubject(subject),
    legacyEvidence: Object.freeze({
      valid: legacyPreflight?.evidence?.valid === true,
      reason: String(
        legacyPreflight?.evidence?.reason || legacyPreflight?.reason || "",
      ),
      substituted: false,
    }),
    gateDecision: Object.freeze({
      allowed: gate?.allowed === true,
      decision: String(gate?.decision || "BLOCK"),
      reason: String(gate?.reason || ""),
    }),
    runtimeExecutionEligible: false,
    providerCallsExecutedByAuthorization: 0,
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
    productionMergeAuthorized: false,
    actualInternalUserExposureExecuted: false,
    actualOperationalTelemetry: false,
    guardrails: Object.freeze({
      singleApprovedSubjectOnly: true,
      generalUsersBlocked: true,
      legacyEvidenceSubstitutionForbidden: true,
      semanticProfilerOnly: true,
      maxProviderCalls: 1,
      primaryResponseAuthority: true,
      responsePayloadMutation: false,
      productionMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      percentageRolloutAuthorized: false,
      failClosed: true,
    }),
  });
}

function evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapRuntime({
  request = {},
  env = process.env,
  featureControl = null,
  legacyPreflight = null,
  mode = null,
  subjectDeriver = deriveQueryCandidatePlannerInternalCanarySubject,
  bootstrapGateEvaluator =
    evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapGate,
} = {}) {
  const resolvedMode =
    mode ||
    parseQueryCandidatePlannerInternalCanaryLiveBootstrapRuntimeMode(env);

  if (!resolvedMode.active) {
    return blockedAuthorization(resolvedMode.reason, {
      mode: resolvedMode,
      legacyPreflight,
    });
  }

  if (!resolvedMode.valid || !resolvedMode.enabled) {
    return blockedAuthorization(resolvedMode.reason, {
      mode: resolvedMode,
      legacyPreflight,
    });
  }

  let subject;
  try {
    subject = subjectDeriver(request);
  } catch {
    return blockedAuthorization("LIVE_BOOTSTRAP_SUBJECT_DERIVATION_FAILED", {
      mode: resolvedMode,
      legacyPreflight,
    });
  }

  let gate;
  try {
    gate = bootstrapGateEvaluator({
      env,
      featureControl,
      subject,
      legacyPreflight,
    });
  } catch {
    return blockedAuthorization("LIVE_BOOTSTRAP_GATE_EVALUATION_FAILED", {
      mode: resolvedMode,
      subject,
      legacyPreflight,
    });
  }

  if (
    gate?.allowed !== true ||
    gate?.decision !== "ALLOW" ||
    gate?.runtimeBootstrapExecutionEligible !== true
  ) {
    return blockedAuthorization(
      String(gate?.reason || "LIVE_BOOTSTRAP_GATE_BLOCKED"),
      {
        mode: resolvedMode,
        subject,
        legacyPreflight,
        gate,
      },
    );
  }

  const config = parseQueryCandidatePlannerInternalCanaryConfig(env);
  if (!config.configurationValid) {
    return blockedAuthorization(
      "LIVE_BOOTSTRAP_INTERNAL_CANARY_CONFIGURATION_INVALID",
      {
        mode: resolvedMode,
        subject,
        legacyPreflight,
        gate,
      },
    );
  }
  if (!config.enabled) {
    return blockedAuthorization(
      "LIVE_BOOTSTRAP_INTERNAL_CANARY_MUST_BE_ENABLED",
      {
        mode: resolvedMode,
        subject,
        legacyPreflight,
        gate,
      },
    );
  }
  if (config.killSwitch) {
    return blockedAuthorization(
      "LIVE_BOOTSTRAP_INTERNAL_CANARY_KILL_SWITCH_ACTIVE",
      {
        mode: resolvedMode,
        subject,
        legacyPreflight,
        gate,
      },
    );
  }
  if (
    config.llmMode !== LLM_MODES.SEMANTIC_PROFILER_ONLY ||
    config.plannerEscalationAllowed !== false
  ) {
    return blockedAuthorization(
      "LIVE_BOOTSTRAP_SEMANTIC_PROFILER_ONLY_REQUIRED",
      {
        mode: resolvedMode,
        subject,
        legacyPreflight,
        gate,
      },
    );
  }

  if (
    gate?.legacyEvidence?.valid !== false ||
    gate?.legacyEvidence?.substituted !== false ||
    String(gate?.legacyEvidence?.reason || "") !==
      "READINESS_EVIDENCE_INVALID"
  ) {
    return blockedAuthorization(
      "LIVE_BOOTSTRAP_LEGACY_EVIDENCE_BOUNDARY_INVALID",
      {
        mode: resolvedMode,
        subject,
        legacyPreflight,
        gate,
      },
    );
  }

  if (
    gate?.percentageRolloutAuthorized === true ||
    gate?.productionPromotionAuthorized === true
  ) {
    return blockedAuthorization(
      "LIVE_BOOTSTRAP_FORBIDDEN_PROMOTION_AUTHORIZATION",
      {
        mode: resolvedMode,
        subject,
        legacyPreflight,
        gate,
      },
    );
  }

  return Object.freeze({
    version: RUNTIME_VERSION,
    active: true,
    allowed: true,
    decision: "ALLOW",
    reason: "SINGLE_SUBJECT_LIVE_BOOTSTRAP_RUNTIME_AUTHORIZED",
    failClosed: true,
    mode: resolvedMode,
    subject: safeSubject(subject),
    config: Object.freeze({
      valid: true,
      enabled: true,
      killSwitch: false,
      timeoutMs: config.timeoutMs,
      llmMode: config.llmMode,
      plannerEscalationAllowed: false,
    }),
    legacyEvidence: Object.freeze({
      valid: false,
      reason: "READINESS_EVIDENCE_INVALID",
      substituted: false,
    }),
    bootstrapReadiness: gate.bootstrapReadiness || null,
    approvalBinding: gate.approvalBinding || null,
    promotionDecision: gate.promotionDecision || null,
    gateDecision: Object.freeze({
      allowed: true,
      decision: "ALLOW",
      reason: String(gate.reason || ""),
    }),
    runtimeExecutionEligible: true,
    providerCallsExecutedByAuthorization: 0,
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
    productionMergeAuthorized: false,
    actualInternalUserExposureExecuted: false,
    actualOperationalTelemetry: false,
    guardrails: Object.freeze({
      singleApprovedSubjectOnly: true,
      generalUsersBlocked: true,
      legacyEvidenceSubstitutionForbidden: true,
      semanticProfilerOnly: true,
      maxProviderCalls: 1,
      primaryResponseAuthority: true,
      responsePayloadMutation: false,
      productionMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      percentageRolloutAuthorized: false,
      failClosed: true,
    }),
  });
}

function bootstrapObserveOnlyMergeAdapter() {
  return Object.freeze({
    version: OBSERVE_ONLY_MERGE_ADAPTER_VERSION,
    status: "BOOTSTRAP_OBSERVE_ONLY",
    reason: "LIVE_BOOTSTRAP_OBSERVE_ONLY_NO_MERGE",
    applied: false,
    primaryPayloadUnchanged: true,
    mergedPayload: null,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    failClosed: true,
  });
}

function buildBootstrapExecutionPreflight(authorization = {}) {
  if (
    authorization?.allowed !== true ||
    authorization?.runtimeExecutionEligible !== true
  ) {
    return null;
  }

  return Object.freeze({
    version: "query_candidate_planner_internal_allowlist_canary_result_v1",
    status: "LIVE_BOOTSTRAP_PREFLIGHT_ALLOWED",
    allowed: true,
    reason: "SINGLE_SUBJECT_LIVE_BOOTSTRAP_PREFLIGHT_ALLOWED",
    config: authorization.config,
    subject: authorization.subject,
    evidence: Object.freeze({
      valid: false,
      reason: "READINESS_EVIDENCE_INVALID",
      evidenceSha256: "",
      summary: Object.freeze({
        source: "PATCH_15_3_3_LIVE_BOOTSTRAP",
        legacyEvidenceSubstituted: false,
        bootstrapReadinessUsedForAuthorizationOnly: true,
        actualOperationalTelemetry: false,
        productionPromotionAuthorized: false,
      }),
      rawEvidenceIncluded: false,
    }),
    readinessGate: null,
    promotionDecision: authorization.promotionDecision || null,
    approvalBinding: authorization.approvalBinding || null,
    bootstrapAuthorization: Object.freeze({
      version: RUNTIME_VERSION,
      allowed: true,
      singleApprovedSubjectOnly: true,
      legacyEvidenceSubstituted: false,
      productionMergeAuthorized: false,
      percentageRolloutAuthorized: false,
    }),
    guardrails: Object.freeze({
      allowlistOnly: true,
      generalUsersBlocked: true,
      deterministicRolloutEnabled: false,
      rolloutPercent: 0,
      primaryFallbackAvailable: true,
      controlledProductionMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      plannerEscalationAllowed: false,
      semanticProfilerOnly: true,
      failClosed: true,
    }),
  });
}

function providerCallCount(observation = {}) {
  const direct = Number(observation?.shadow?.providerCallCount);
  if (Number.isFinite(direct) && direct >= 0) return direct;
  const fallback = Number(observation?.providerCallCount);
  return Number.isFinite(fallback) && fallback >= 0 ? fallback : 0;
}

function safeRuntimeObservation({
  status,
  reason,
  authorization,
  canaryObservation = null,
  providerCalls = 0,
} = {}) {
  return Object.freeze({
    version: OBSERVATION_VERSION,
    status: String(status || "LIVE_BOOTSTRAP_BLOCKED"),
    reason: String(reason || "LIVE_BOOTSTRAP_RUNTIME_BLOCKED"),
    subjectTagSha256: String(authorization?.subject?.subjectTagSha256 || ""),
    responseSource: "PRIMARY",
    providerCallCount: Math.max(0, Number(providerCalls) || 0),
    comparisonStatus: String(
      canaryObservation?.comparison?.status ||
        canaryObservation?.comparison?.result ||
        "",
    ),
    latencyMs: Math.max(0, Number(canaryObservation?.latencyMs) || 0),
    legacyEvidenceSubstituted: false,
    merge: Object.freeze({
      applied: false,
      status: "NOT_APPLIED",
      reason: "LIVE_BOOTSTRAP_OBSERVE_ONLY_NO_MERGE",
      primaryPayloadUnchanged: true,
      productionReadyAssignment: false,
      productionRouteChanged: false,
    }),
    actualInternalUserExposureExecuted:
      status === "LIVE_BOOTSTRAP_OBSERVED" ||
      status === "LIVE_BOOTSTRAP_FALLBACK_SAFE",
    operationalTelemetryEvidenceEligible:
      status === "LIVE_BOOTSTRAP_OBSERVED" ||
      status === "LIVE_BOOTSTRAP_FALLBACK_SAFE",
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
    guardrails: Object.freeze({
      singleApprovedSubjectOnly: true,
      generalUsersBlocked: true,
      primaryResponseAuthority: true,
      responsePayloadMutation: false,
      semanticProfilerOnly: true,
      maxProviderCalls: 1,
      productionMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      failClosed: true,
    }),
    privacy: Object.freeze({
      rawIdentityIncluded: false,
      rawPrimaryResponseIncluded: false,
      rawShadowResolutionIncluded: false,
      rawRowsIncluded: false,
      fileNameIncluded: false,
      queryTablesKeyIncluded: false,
      emailIncluded: false,
    }),
  });
}

function blockedRuntimeResult(primaryPayload, authorization) {
  const observation = safeRuntimeObservation({
    status: "LIVE_BOOTSTRAP_BLOCKED",
    reason: authorization?.reason || "LIVE_BOOTSTRAP_RUNTIME_BLOCKED",
    authorization,
    providerCalls: 0,
  });

  return Object.freeze({
    version: RESULT_VERSION,
    status: "PRIMARY_BLOCKED_SAFE",
    reason: observation.reason,
    responseSource: "PRIMARY",
    responsePayload: primaryPayload,
    primaryPayload,
    authorization,
    observation,
    providerCallCount: 0,
    actualInternalUserExposureExecuted: false,
    operationalTelemetryEvidenceEligible: false,
    legacyEvidenceSubstituted: false,
    productionMergeApplied: false,
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
    failClosed: true,
  });
}

async function runQueryCandidatePlannerInternalCanaryLiveBootstrap({
  request = {},
  primaryPayload = {},
  env = process.env,
  featureControl = null,
  legacyPreflight = null,
  authorization = null,
  now = Date.now,
  canaryRunner = runQueryCandidatePlannerInternalAllowlistCanary,
} = {}) {
  const resolvedAuthorization =
    authorization ||
    evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapRuntime({
      request,
      env,
      featureControl,
      legacyPreflight,
    });

  if (
    resolvedAuthorization?.allowed !== true ||
    resolvedAuthorization?.runtimeExecutionEligible !== true
  ) {
    return blockedRuntimeResult(primaryPayload, resolvedAuthorization);
  }

  const preflight = buildBootstrapExecutionPreflight(resolvedAuthorization);
  if (!preflight) {
    return blockedRuntimeResult(primaryPayload, resolvedAuthorization);
  }

  try {
    const result = await canaryRunner({
      request,
      primaryPayload,
      env,
      featureControl,
      preflight,
      mergeAdapter: bootstrapObserveOnlyMergeAdapter,
      now,
    });

    const calls = providerCallCount(result?.observation);
    const guardrailViolation =
      calls > 1 ||
      result?.observation?.shadow?.plannerEscalationUsed === true ||
      result?.observation?.merge?.applied === true ||
      result?.mergeResult?.applied === true;

    const observation = safeRuntimeObservation({
      status: guardrailViolation
        ? "LIVE_BOOTSTRAP_GUARDRAIL_VIOLATION"
        : "LIVE_BOOTSTRAP_OBSERVED",
      reason: guardrailViolation
        ? "LIVE_BOOTSTRAP_RUNTIME_GUARDRAIL_VIOLATION"
        : "SINGLE_SUBJECT_LIVE_BOOTSTRAP_OBSERVED",
      authorization: resolvedAuthorization,
      canaryObservation: result?.observation,
      providerCalls: calls,
    });

    return Object.freeze({
      version: RESULT_VERSION,
      status: guardrailViolation
        ? "PRIMARY_FALLBACK_SAFE"
        : "LIVE_BOOTSTRAP_OBSERVED",
      reason: observation.reason,
      responseSource: "PRIMARY",
      responsePayload: primaryPayload,
      primaryPayload,
      authorization: resolvedAuthorization,
      canaryResultStatus: String(result?.status || ""),
      canaryObservation: result?.observation || null,
      observation,
      providerCallCount: calls,
      actualInternalUserExposureExecuted: true,
      operationalTelemetryEvidenceEligible: true,
      legacyEvidenceSubstituted: false,
      productionMergeApplied: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
      failClosed: true,
    });
  } catch (error) {
    const reason = String(
      error?.code || "LIVE_BOOTSTRAP_RUNTIME_EXECUTION_FAILED_SAFE",
    );
    const observation = safeRuntimeObservation({
      status: "LIVE_BOOTSTRAP_FALLBACK_SAFE",
      reason,
      authorization: resolvedAuthorization,
      providerCalls: 0,
    });

    return Object.freeze({
      version: RESULT_VERSION,
      status: "PRIMARY_FALLBACK_SAFE",
      reason,
      responseSource: "PRIMARY",
      responsePayload: primaryPayload,
      primaryPayload,
      authorization: resolvedAuthorization,
      observation,
      providerCallCount: 0,
      actualInternalUserExposureExecuted: true,
      operationalTelemetryEvidenceEligible: true,
      legacyEvidenceSubstituted: false,
      productionMergeApplied: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
      failClosed: true,
    });
  }
}

function defaultQueryCandidatePlannerLiveBootstrapObservationLogger(
  observation = {},
) {
  console.info("[query-candidate-live-bootstrap]", {
    version: String(observation.version || ""),
    status: String(observation.status || "UNKNOWN"),
    reason: String(observation.reason || ""),
    subjectTagSha256: String(observation.subjectTagSha256 || ""),
    responseSource: "PRIMARY",
    providerCallCount: Math.max(
      0,
      Number(observation.providerCallCount || 0),
    ),
    comparisonStatus: String(observation.comparisonStatus || ""),
    latencyMs: Math.max(0, Number(observation.latencyMs || 0)),
    legacyEvidenceSubstituted: false,
    productionMergeApplied: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    rawIdentityIncluded: false,
  });
}

module.exports = Object.freeze({
  RUNTIME_VERSION,
  OBSERVATION_VERSION,
  RESULT_VERSION,
  OBSERVE_ONLY_MERGE_ADAPTER_VERSION,
  parseQueryCandidatePlannerInternalCanaryLiveBootstrapRuntimeMode,
  evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapRuntime,
  bootstrapObserveOnlyMergeAdapter,
  buildBootstrapExecutionPreflight,
  safeRuntimeObservation,
  blockedRuntimeResult,
  runQueryCandidatePlannerInternalCanaryLiveBootstrap,
  defaultQueryCandidatePlannerLiveBootstrapObservationLogger,
});
