const { OPERATIONS } = require("./queryCandidatePlannerFeatureControl");
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

const SERVICE_VERSION =
  "query_candidate_planner_api_shadow_service_v2_cache_lifecycle";
const OBSERVATION_VERSION = "query_candidate_planner_api_shadow_observation_v1";
const DEFAULT_TIMEOUT_MS = 30000;

function errorCode(error) {
  const code = String(error?.code || "").trim();
  if (code) return code.slice(0, 120);
  return "API_SHADOW_EXECUTION_FAILED";
}

function nowMs(now) {
  const value = Number(now());
  return Number.isFinite(value) ? value : Date.now();
}

async function runWithTimeout(task, timeoutMs, abortController) {
  let timer = null;
  try {
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(() => {
        abortController.abort();
        const error = new Error("API shadow timeout");
        error.code = "API_SHADOW_TIMEOUT";
        reject(error);
      }, timeoutMs);
    });
    return await Promise.race([task, timeout]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function guardrails() {
  return Object.freeze({
    shadowOnly: true,
    primaryResponseAuthority: true,
    responsePayloadMutation: false,
    responseHeaderMutation: false,
    responseStatusMutation: false,
    productionCandidateMerge: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
    sourceCandidateStatusMutation: false,
  });
}

function blockedObservation({ decision, primaryPayload }) {
  return Object.freeze({
    version: OBSERVATION_VERSION,
    serviceVersion: SERVICE_VERSION,
    status: "BLOCKED",
    reason: decision.reason,
    featureDecision: decision,
    primaryResponseSha256: primaryResponseContractSha256(primaryPayload),
    comparison: null,
    latencyMs: 0,
    guardrails: guardrails(),
    privacy: Object.freeze({ rawPrimaryResponseIncluded: false }),
  });
}

async function observeQueryCandidatePlannerApiShadow({
  request = {},
  primaryPayload = {},
  featureControl,
  shadowRunner = runQueryCandidatePlannerApiShadow,
  comparator = compareCandidatePlannerShadow,
  timeoutMs = DEFAULT_TIMEOUT_MS,
  now = Date.now,
} = {}) {
  const shadowDecision = featureControl.evaluate(OPERATIONS.SHADOW_EXECUTION);
  if (!shadowDecision.allowed) {
    return blockedObservation({
      decision: shadowDecision,
      primaryPayload,
    });
  }

  const providerDecision = featureControl.evaluate(OPERATIONS.PROVIDER_CALL);
  const cacheReadDecision = featureControl.evaluate(OPERATIONS.CACHE_READ);
  const cacheWriteDecision = featureControl.evaluate(OPERATIONS.CACHE_WRITE);

  const startedAt = nowMs(now);
  const primaryResponseSha256 = primaryResponseContractSha256(primaryPayload);
  const safeContext = buildSafeApiShadowContext({
    request,
    primaryPayload,
  });
  const lifecycleIdentity = deriveQueryCandidatePlannerUploadIdentity({
    request,
    primaryPayload,
  });
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
        }),
      ),
      timeoutMs,
      abortController,
    );

    const comparison = comparator({
      primaryPayload,
      shadowResolution,
    });
    const completedStatus =
      shadowResolution?.status === "SHADOW_COMPLETED"
        ? "COMPLETED"
        : "COMPLETED_SAFE";

    return Object.freeze({
      version: OBSERVATION_VERSION,
      serviceVersion: SERVICE_VERSION,
      status: completedStatus,
      reason: String(shadowResolution?.status || "SHADOW_RESULT_OBSERVED"),
      featureDecision: shadowDecision,
      providerDecision,
      cacheReadDecision,
      cacheWriteDecision,
      requestFingerprintSha256: safeContext.requestFingerprintSha256,
      cacheLifecycle: Object.freeze({
        identity: publicUploadIdentity(lifecycleIdentity),
        cacheReadAllowed: cacheReadDecision?.allowed === true,
        cacheWriteAllowed: cacheWriteDecision?.allowed === true,
        tenantIdIncluded: false,
        cacheSecretIncluded: false,
      }),
      primaryResponseSha256,
      primaryResponseUnchanged:
        primaryResponseSha256 === primaryResponseContractSha256(primaryPayload),
      shadow: Object.freeze({
        status: String(shadowResolution?.status || "UNKNOWN"),
        invocationStatus: String(
          shadowResolution?.plannerResolution?.invocation?.status ||
            shadowResolution?.invocationStatus ||
            "",
        ),
        providerCallCount: Number(
          shadowResolution?.plannerResolution?.invocation?.providerCallCount ??
            shadowResolution?.providerCallCount ??
            0,
        ),
        accepted: Number(
          shadowResolution?.plannerResolution?.counts?.accepted ??
            shadowResolution?.counts?.accepted ??
            0,
        ),
        productionCandidateMerge:
          shadowResolution?.policy?.productionCandidateMerge === true,
        productionReadyAssignment:
          shadowResolution?.policy?.productionReadyAssignment === true,
        productionRouteChanged:
          shadowResolution?.policy?.productionRouteChanged === true,
      }),
      comparison,
      latencyMs: Math.max(0, nowMs(now) - startedAt),
      guardrails: guardrails(),
      privacy: Object.freeze({
        safeContextVersion: safeContext.version,
        rawPrimaryResponseIncluded: false,
        rawRowsIncluded: false,
        sampleValuesIncluded: false,
        fileNameIncluded: false,
        queryTablesKeyIncluded: false,
        tenantIdIncluded: false,
      }),
    });
  } catch (error) {
    const code = errorCode(error);
    return Object.freeze({
      version: OBSERVATION_VERSION,
      serviceVersion: SERVICE_VERSION,
      status: code === "API_SHADOW_TIMEOUT" ? "TIMEOUT_SAFE" : "FAILED_SAFE",
      reason: code,
      featureDecision: shadowDecision,
      providerDecision,
      cacheReadDecision,
      cacheWriteDecision,
      requestFingerprintSha256: safeContext.requestFingerprintSha256,
      cacheLifecycle: Object.freeze({
        identity: publicUploadIdentity(lifecycleIdentity),
        cacheReadAllowed: cacheReadDecision?.allowed === true,
        cacheWriteAllowed: cacheWriteDecision?.allowed === true,
        tenantIdIncluded: false,
        cacheSecretIncluded: false,
      }),
      primaryResponseSha256,
      primaryResponseUnchanged:
        primaryResponseSha256 === primaryResponseContractSha256(primaryPayload),
      comparison: null,
      latencyMs: Math.max(0, nowMs(now) - startedAt),
      guardrails: guardrails(),
      privacy: Object.freeze({
        rawErrorMessageIncluded: false,
        rawPrimaryResponseIncluded: false,
        rawRowsIncluded: false,
        sampleValuesIncluded: false,
        fileNameIncluded: false,
      }),
    });
  }
}

module.exports = Object.freeze({
  SERVICE_VERSION,
  OBSERVATION_VERSION,
  DEFAULT_TIMEOUT_MS,
  observeQueryCandidatePlannerApiShadow,
});
