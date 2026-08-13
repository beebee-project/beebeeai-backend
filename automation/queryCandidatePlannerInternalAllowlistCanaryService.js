const {
  evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate,
} = require("./queryCandidatePlannerInternalCanaryApprovalBindingGate");
const { OPERATIONS } = require("./queryCandidatePlannerFeatureControl");
const {
  getQueryCandidatePlannerFeatureControl,
} = require("./queryCandidatePlannerFeatureControlRuntime");
const {
  compareCandidatePlannerShadow,
} = require("./queryCandidatePlannerShadowComparator");
const {
  buildSafeApiShadowContext,
  primaryResponseContractSha256,
  runQueryCandidatePlannerApiShadow,
} = require("./queryCandidatePlannerApiShadowRunner");
const {
  deriveQueryCandidatePlannerUploadIdentity,
  publicUploadIdentity,
} = require("./queryCandidatePlannerUploadLifecycle");
const {
  controlledProductionMergeAdapter,
} = require("./queryCandidatePlannerControlledProductionMergeAdapter");
const {
  AUDIENCE_MODES,
  evaluateControlledProductionPromotionGate,
  parsePromotionGateEnvironment,
} = require("./queryCandidatePlannerControlledProductionPromotionGate");
const {
  parseQueryCandidatePlannerInternalCanaryConfig,
  LLM_MODES,
} = require("./queryCandidatePlannerInternalAllowlistCanaryConfig");
const {
  parseEvidenceJson,
  validateQueryCandidatePlannerInternalCanaryEvidence,
} = require("./queryCandidatePlannerInternalCanaryEvidence");
const {
  deriveQueryCandidatePlannerInternalCanarySubject,
} = require("./queryCandidatePlannerInternalCanarySubject");

const SERVICE_VERSION =
  "query_candidate_planner_internal_allowlist_canary_service_v1";
const RESULT_VERSION =
  "query_candidate_planner_internal_allowlist_canary_result_v1";
const OBSERVATION_VERSION =
  "query_candidate_planner_internal_allowlist_canary_observation_v1";

function nowMs(now) {
  const value = Number(now());
  return Number.isFinite(value) ? value : Date.now();
}

function errorCode(error) {
  return String(error?.code || "INTERNAL_CANARY_EXECUTION_FAILED")
    .trim()
    .slice(0, 120);
}

async function runWithTimeout(task, timeoutMs, abortController) {
  let timer = null;
  try {
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(() => {
        abortController.abort();
        const error = new Error("Internal canary timeout");
        error.code = "INTERNAL_CANARY_TIMEOUT";
        reject(error);
      }, timeoutMs);
    });
    return await Promise.race([task, timeout]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function canaryGuardrails(overrides = {}) {
  return Object.freeze({
    allowlistOnly: true,
    generalUsersBlocked: true,
    deterministicRolloutEnabled: false,
    rolloutPercent: 0,
    primaryFallbackAvailable: true,
    primaryPayloadMutated: false,
    responseHeaderMutation: false,
    responseStatusMutation: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    plannerEscalationAllowed: false,
    semanticProfilerOnly: true,
    failClosed: true,
    ...overrides,
  });
}

function parseEvidence(config, evidenceBundle, now) {
  if (evidenceBundle && typeof evidenceBundle === "object") {
    return validateQueryCandidatePlannerInternalCanaryEvidence(evidenceBundle, {
      now,
    });
  }
  const parsed = parseEvidenceJson(config.evidenceJson);
  if (parsed.error) {
    return Object.freeze({
      valid: false,
      reason: parsed.error,
      errors: Object.freeze([parsed.error]),
      readiness: null,
      evidenceSha256: "",
      summary: Object.freeze({ rawEvidenceIncluded: false }),
      failClosed: true,
    });
  }
  return validateQueryCandidatePlannerInternalCanaryEvidence(parsed.value, {
    now,
  });
}

function blockedPreflight({
  reason,
  config,
  subject,
  evidence,
  promotionDecision = null,
}) {
  return Object.freeze({
    version: RESULT_VERSION,
    status: "BLOCKED",
    allowed: false,
    reason,
    config: Object.freeze({
      valid: config.configurationValid,
      enabled: config.enabled,
      killSwitch: config.killSwitch,
      llmMode: config.llmMode,
      plannerEscalationAllowed: false,
    }),
    subject: Object.freeze({
      complete: subject.complete === true,
      reason: subject.reason,
      subjectSha256: subject.subjectSha256 || "",
      subjectTagSha256: subject.subjectTagSha256 || "",
      rawIdentityIncluded: false,
    }),
    evidence: Object.freeze({
      valid: evidence?.valid === true,
      reason: evidence?.reason || "CANARY_EVIDENCE_NOT_EVALUATED",
      evidenceSha256: evidence?.evidenceSha256 || "",
      summary: evidence?.summary || null,
      rawEvidenceIncluded: false,
    }),
    promotionDecision,
    guardrails: canaryGuardrails(),
  });
}

function evaluateQueryCandidatePlannerInternalCanaryPreflight({
  request = {},
  env = process.env,
  featureControl = null,
  evidenceBundle = null,
  config = null,
  now = Date.now,
} = {}) {
  const resolvedConfig =
    config || parseQueryCandidatePlannerInternalCanaryConfig(env);
  const subject = deriveQueryCandidatePlannerInternalCanarySubject(request);

  // PATCH 15.3.2-F.1.6.1 APPROVAL BINDING COMPOSITION
  const approvalBindingGate =
    evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate({
      env,
      featureControl,
      subject,
    });

  // Approval is an additional prerequisite only.
  // BLOCK returns immediately; ALLOW must continue through every
  // existing Patch 15.3 preflight check below.
  if (!approvalBindingGate.allowed) {
    return approvalBindingGate.preflight;
  }

  // PATCH 15.3.2-F.1.6.1 LEGACY PREFLIGHT CONTINUES
  const evidence = parseEvidence(resolvedConfig, evidenceBundle, now);

  if (!resolvedConfig.configurationValid) {
    return blockedPreflight({
      reason: "INVALID_INTERNAL_CANARY_CONFIGURATION",
      config: resolvedConfig,
      subject,
      evidence,
    });
  }
  if (!resolvedConfig.enabled) {
    return blockedPreflight({
      reason: "INTERNAL_CANARY_DISABLED",
      config: resolvedConfig,
      subject,
      evidence,
    });
  }
  if (resolvedConfig.killSwitch) {
    return blockedPreflight({
      reason: "INTERNAL_CANARY_KILL_SWITCH_ACTIVE",
      config: resolvedConfig,
      subject,
      evidence,
    });
  }
  if (
    resolvedConfig.llmMode !== LLM_MODES.SEMANTIC_PROFILER_ONLY ||
    resolvedConfig.plannerEscalationAllowed !== false
  ) {
    return blockedPreflight({
      reason: "SEMANTIC_PROFILER_ONLY_POLICY_REQUIRED",
      config: resolvedConfig,
      subject,
      evidence,
    });
  }
  if (!subject.complete) {
    return blockedPreflight({
      reason: subject.reason,
      config: resolvedConfig,
      subject,
      evidence,
    });
  }
  if (!evidence.valid) {
    return blockedPreflight({
      reason: evidence.reason,
      config: resolvedConfig,
      subject,
      evidence,
    });
  }

  const promotionConfig = parsePromotionGateEnvironment(env);
  if (
    promotionConfig.audienceMode !== AUDIENCE_MODES.ALLOWLIST ||
    promotionConfig.rolloutPercent !== 0
  ) {
    return blockedPreflight({
      reason: "ALLOWLIST_ONLY_PROMOTION_CONFIGURATION_REQUIRED",
      config: resolvedConfig,
      subject,
      evidence,
    });
  }

  const control = featureControl || getQueryCandidatePlannerFeatureControl();
  const promotionDecision = evaluateControlledProductionPromotionGate({
    env,
    featureControl: control,
    readinessGate: evidence.readiness,
    subjectSha256: subject.subjectSha256,
  });
  if (!promotionDecision.allowed) {
    return blockedPreflight({
      reason: promotionDecision.reason,
      config: resolvedConfig,
      subject,
      evidence,
      promotionDecision,
    });
  }

  return Object.freeze({
    version: RESULT_VERSION,
    status: "ALLOWLIST_PREFLIGHT_ALLOWED",
    allowed: true,
    reason: "INTERNAL_ALLOWLIST_CANARY_PREFLIGHT_ALLOWED",
    config: Object.freeze({
      valid: true,
      enabled: true,
      killSwitch: false,
      timeoutMs: resolvedConfig.timeoutMs,
      llmMode: resolvedConfig.llmMode,
      plannerEscalationAllowed: false,
    }),
    subject: Object.freeze({
      complete: true,
      reason: subject.reason,
      subjectSha256: subject.subjectSha256,
      subjectTagSha256: subject.subjectTagSha256,
      rawIdentityIncluded: false,
    }),
    evidence: Object.freeze({
      valid: true,
      reason: evidence.reason,
      evidenceSha256: evidence.evidenceSha256,
      summary: evidence.summary,
      rawEvidenceIncluded: false,
    }),
    readinessGate: evidence.readiness,
    promotionDecision,
    guardrails: canaryGuardrails(),
  });
}

function providerCallCount(shadowResolution = {}) {
  return Number(
    shadowResolution?.plannerResolution?.invocation?.providerCallCount ??
      shadowResolution?.providerCallCount ??
      0,
  );
}

function plannerEscalationUsed(shadowResolution = {}) {
  return Boolean(
    shadowResolution?.plannerEscalationUsed === true ||
    shadowResolution?.plannerResolution?.invocation?.plannerEscalationUsed ===
      true ||
    shadowResolution?.policy?.plannerEscalationAllowed === true,
  );
}

function shadowGuardrailViolation(shadowResolution = {}) {
  return Boolean(
    shadowResolution?.policy?.productionCandidateMerge === true ||
    shadowResolution?.policy?.productionReadyAssignment === true ||
    shadowResolution?.policy?.productionRouteChanged === true ||
    providerCallCount(shadowResolution) > 1 ||
    plannerEscalationUsed(shadowResolution),
  );
}

function buildObservation({
  status,
  reason,
  preflight,
  primaryPayload,
  shadowResolution = null,
  comparison = null,
  mergeResult = null,
  lifecycleIdentity = null,
  latencyMs = 0,
}) {
  return Object.freeze({
    version: OBSERVATION_VERSION,
    serviceVersion: SERVICE_VERSION,
    status,
    reason,
    subjectTagSha256: preflight.subject.subjectTagSha256 || "",
    evidenceSha256: preflight.evidence.evidenceSha256 || "",
    primaryResponseSha256: primaryResponseContractSha256(primaryPayload),
    responseSource:
      mergeResult?.applied === true ? "CONTROLLED_PLANNER" : "PRIMARY_FALLBACK",
    promotion: Object.freeze({
      allowed: preflight.promotionDecision?.allowed === true,
      reason: preflight.promotionDecision?.reason || preflight.reason,
      audiencePath: preflight.promotionDecision?.audience?.path || "NONE",
      allowlistMatched:
        preflight.promotionDecision?.audience?.allowlistMatched === true,
      rolloutPercent: 0,
    }),
    shadow: Object.freeze({
      status: String(shadowResolution?.status || "NOT_RUN"),
      providerCallCount: Math.max(0, providerCallCount(shadowResolution)),
      plannerEscalationUsed: plannerEscalationUsed(shadowResolution),
      semanticProfilerOnly: true,
      rawResolutionIncluded: false,
    }),
    merge: Object.freeze({
      status: String(mergeResult?.status || "NOT_APPLIED"),
      reason: String(mergeResult?.reason || reason),
      applied: mergeResult?.applied === true,
      primaryPayloadUnchanged: mergeResult?.primaryPayloadUnchanged !== false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
    }),
    cacheLifecycle: Object.freeze({
      identity: publicUploadIdentity(lifecycleIdentity || {}),
      tenantIdIncluded: false,
      cacheSecretIncluded: false,
    }),
    comparison,
    latencyMs: Math.max(0, Number(latencyMs) || 0),
    guardrails: canaryGuardrails({
      controlledProductionMergeApplied: mergeResult?.applied === true,
    }),
    privacy: Object.freeze({
      rawPrimaryResponseIncluded: false,
      rawShadowResolutionIncluded: false,
      rawIdentityIncluded: false,
      rawEvidenceIncluded: false,
      rawRowsIncluded: false,
      fileNameIncluded: false,
      queryTablesKeyIncluded: false,
      tenantIdIncluded: false,
      emailIncluded: false,
    }),
  });
}

function blockedResult(primaryPayload, preflight) {
  const observation = buildObservation({
    status: "BLOCKED",
    reason: preflight.reason,
    preflight,
    primaryPayload,
  });
  return Object.freeze({
    version: RESULT_VERSION,
    status: "PRIMARY_BLOCKED",
    reason: preflight.reason,
    responseSource: "PRIMARY",
    responsePayload: primaryPayload,
    primaryPayload,
    mergeResult: null,
    observation,
    preflight,
    guardrails: canaryGuardrails(),
  });
}

async function runQueryCandidatePlannerInternalAllowlistCanary({
  request = {},
  primaryPayload = {},
  env = process.env,
  featureControl = null,
  evidenceBundle = null,
  config = null,
  preflight = null,
  shadowRunner = runQueryCandidatePlannerApiShadow,
  comparator = compareCandidatePlannerShadow,
  mergeAdapter = controlledProductionMergeAdapter,
  now = Date.now,
} = {}) {
  const control = featureControl || getQueryCandidatePlannerFeatureControl();
  const resolvedPreflight =
    preflight ||
    evaluateQueryCandidatePlannerInternalCanaryPreflight({
      request,
      env,
      featureControl: control,
      evidenceBundle,
      config,
      now,
    });

  if (!resolvedPreflight.allowed) {
    return blockedResult(primaryPayload, resolvedPreflight);
  }

  const startedAt = nowMs(now);
  const safeContext = buildSafeApiShadowContext({ request, primaryPayload });
  const lifecycleIdentity = deriveQueryCandidatePlannerUploadIdentity({
    request,
    primaryPayload,
  });
  const providerDecision = control.evaluate(OPERATIONS.PROVIDER_CALL);
  const cacheReadDecision = control.evaluate(OPERATIONS.CACHE_READ);
  const cacheWriteDecision = control.evaluate(OPERATIONS.CACHE_WRITE);
  const abortController = new AbortController();

  try {
    const shadowResolution = await runWithTimeout(
      Promise.resolve(
        shadowRunner({
          safeContext,
          lifecycleIdentity,
          providerDecision,
          cacheReadDecision,
          cacheWriteDecision,
          signal: abortController.signal,
          llmPolicy: Object.freeze({
            mode: LLM_MODES.SEMANTIC_PROFILER_ONLY,
            maxProviderCalls: 1,
            plannerEscalationAllowed: false,
          }),
        }),
      ),
      resolvedPreflight.config.timeoutMs,
      abortController,
    );

    if (shadowGuardrailViolation(shadowResolution)) {
      const observation = buildObservation({
        status: "FALLBACK_SAFE",
        reason: "SHADOW_GUARDRAIL_VIOLATION",
        preflight: resolvedPreflight,
        primaryPayload,
        shadowResolution,
        lifecycleIdentity,
        latencyMs: nowMs(now) - startedAt,
      });
      return Object.freeze({
        version: RESULT_VERSION,
        status: "PRIMARY_FALLBACK",
        reason: "SHADOW_GUARDRAIL_VIOLATION",
        responseSource: "PRIMARY",
        responsePayload: primaryPayload,
        primaryPayload,
        mergeResult: null,
        observation,
        preflight: resolvedPreflight,
        guardrails: canaryGuardrails(),
      });
    }

    const comparison = comparator({ primaryPayload, shadowResolution });
    const mergeResult = mergeAdapter({
      primaryPayload,
      shadowResolution,
      featureControl: control,
      readinessGate: resolvedPreflight.readinessGate,
      promotionGateDecision: resolvedPreflight.promotionDecision,
      apply: true,
    });

    const applied =
      mergeResult?.applied === true &&
      mergeResult?.status === "MERGED_COPY_READY" &&
      mergeResult?.primaryPayloadUnchanged === true &&
      mergeResult?.mergedPayload &&
      typeof mergeResult.mergedPayload === "object";
    const responsePayload = applied
      ? mergeResult.mergedPayload
      : primaryPayload;
    const observation = buildObservation({
      status: applied ? "CANARY_MERGED" : "FALLBACK_SAFE",
      reason: applied
        ? "INTERNAL_ALLOWLIST_CANARY_MERGED"
        : String(mergeResult?.reason || "MERGE_ADAPTER_BLOCKED"),
      preflight: resolvedPreflight,
      primaryPayload,
      shadowResolution,
      comparison,
      mergeResult,
      lifecycleIdentity,
      latencyMs: nowMs(now) - startedAt,
    });

    return Object.freeze({
      version: RESULT_VERSION,
      status: applied ? "CANARY_MERGED" : "PRIMARY_FALLBACK",
      reason: observation.reason,
      responseSource: applied ? "CONTROLLED_PLANNER" : "PRIMARY",
      responsePayload,
      primaryPayload,
      mergeResult,
      comparison,
      observation,
      preflight: resolvedPreflight,
      guardrails: canaryGuardrails({
        controlledProductionMergeApplied: applied,
      }),
    });
  } catch (error) {
    const reason = errorCode(error);
    const observation = buildObservation({
      status:
        reason === "INTERNAL_CANARY_TIMEOUT"
          ? "TIMEOUT_FALLBACK_SAFE"
          : "FAILED_FALLBACK_SAFE",
      reason,
      preflight: resolvedPreflight,
      primaryPayload,
      lifecycleIdentity,
      latencyMs: nowMs(now) - startedAt,
    });
    return Object.freeze({
      version: RESULT_VERSION,
      status: "PRIMARY_FALLBACK",
      reason,
      responseSource: "PRIMARY",
      responsePayload: primaryPayload,
      primaryPayload,
      mergeResult: null,
      observation,
      preflight: resolvedPreflight,
      guardrails: canaryGuardrails(),
    });
  }
}

module.exports = Object.freeze({
  SERVICE_VERSION,
  RESULT_VERSION,
  OBSERVATION_VERSION,
  canaryGuardrails,
  evaluateQueryCandidatePlannerInternalCanaryPreflight,
  runQueryCandidatePlannerInternalAllowlistCanary,
});
