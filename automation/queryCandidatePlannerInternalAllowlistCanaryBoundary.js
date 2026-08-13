const {
  getQueryCandidatePlannerFeatureControl,
} = require("./queryCandidatePlannerFeatureControlRuntime");
const {
  observeQueryCandidatePlannerApiShadow,
} = require("./queryCandidatePlannerApiShadowService");
const {
  evaluateQueryCandidatePlannerInternalCanaryPreflight,
  runQueryCandidatePlannerInternalAllowlistCanary,
} = require("./queryCandidatePlannerInternalAllowlistCanaryService");
const {
  parseQueryCandidatePlannerInternalCanaryLiveBootstrapRuntimeMode,
  evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapRuntime,
  runQueryCandidatePlannerInternalCanaryLiveBootstrap,
  defaultQueryCandidatePlannerLiveBootstrapObservationLogger,
  safeRuntimeObservation,
} = require("./queryCandidatePlannerInternalCanaryLiveBootstrapRuntime");

const BOUNDARY_VERSION =
  "query_candidate_planner_internal_allowlist_canary_boundary_v1";

function defaultCanaryObservationLogger(observation = {}) {
  const summary = {
    version: observation.version || "",
    status: observation.status || "UNKNOWN",
    reason: observation.reason || "",
    subjectTagSha256: observation.subjectTagSha256 || "",
    evidenceSha256: observation.evidenceSha256 || "",
    responseSource: observation.responseSource || "PRIMARY",
    allowlistMatched: observation.promotion?.allowlistMatched === true,
    providerCallCount: Number(observation.shadow?.providerCallCount || 0),
    plannerEscalationUsed: observation.shadow?.plannerEscalationUsed === true,
    mergeApplied: observation.merge?.applied === true,
    latencyMs: Number(observation.latencyMs || 0),
    productionReadyAssignment: false,
    productionRouteChanged: false,
    rawIdentityIncluded: false,
  };
  console.info("[query-candidate-internal-canary]", summary);
}

function safeBlockedCanaryObservation(preflight = {}) {
  return Object.freeze({
    version: "query_candidate_planner_internal_allowlist_canary_observation_v1",
    status: "BLOCKED",
    reason: String(preflight.reason || "INTERNAL_CANARY_BLOCKED"),
    subjectTagSha256: String(preflight.subject?.subjectTagSha256 || ""),
    evidenceSha256: String(preflight.evidence?.evidenceSha256 || ""),
    responseSource: "PRIMARY",
    promotion: Object.freeze({
      allowed: false,
      reason: String(preflight.reason || "INTERNAL_CANARY_BLOCKED"),
      audiencePath: String(
        preflight.promotionDecision?.audience?.path || "NONE",
      ),
      allowlistMatched:
        preflight.promotionDecision?.audience?.allowlistMatched === true,
      rolloutPercent: 0,
    }),
    shadow: Object.freeze({
      status: "ASYNC_SHADOW_ONLY",
      providerCallCount: 0,
      plannerEscalationUsed: false,
      semanticProfilerOnly: true,
      rawResolutionIncluded: false,
    }),
    merge: Object.freeze({
      status: "NOT_APPLIED",
      reason: String(preflight.reason || "INTERNAL_CANARY_BLOCKED"),
      applied: false,
      primaryPayloadUnchanged: true,
      productionReadyAssignment: false,
      productionRouteChanged: false,
    }),
    comparison: null,
    latencyMs: 0,
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
    privacy: Object.freeze({
      rawIdentityIncluded: false,
      rawEvidenceIncluded: false,
      rawPrimaryResponseIncluded: false,
      rawShadowResolutionIncluded: false,
    }),
  });
}

function createQueryCandidatePlannerInternalAllowlistCanaryBoundary({
  handler,
  featureControl = null,
  env = process.env,
  evidenceBundle = null,
  preflightEvaluator = evaluateQueryCandidatePlannerInternalCanaryPreflight,
  canaryRunner = runQueryCandidatePlannerInternalAllowlistCanary,
  shadowObserve = observeQueryCandidatePlannerApiShadow,
  shadowRunner,
  comparator,
  onObservation = null,
  onCanaryObservation = defaultCanaryObservationLogger,
  bootstrapModeParser =
    parseQueryCandidatePlannerInternalCanaryLiveBootstrapRuntimeMode,
  bootstrapAuthorize =
    evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapRuntime,
  bootstrapRunner = runQueryCandidatePlannerInternalCanaryLiveBootstrap,
  onLiveBootstrapObservation =
    defaultQueryCandidatePlannerLiveBootstrapObservationLogger,
  now = Date.now,
} = {}) {
  if (typeof handler !== "function") {
    throw new TypeError("Internal canary boundary handler must be a function");
  }

  return async function queryCandidatePlannerInternalAllowlistCanaryBoundary(
    req,
    res,
    next,
  ) {
    const originalJson = res.json.bind(res);
    let responseCaptured = false;
    let responseTask = null;

    res.json = function internalCanaryObservedJson(primaryPayload) {
      if (responseCaptured) return originalJson(primaryPayload);
      responseCaptured = true;

      const control =
        featureControl || getQueryCandidatePlannerFeatureControl();
      const preflight = preflightEvaluator({
        request: req,
        env,
        featureControl: control,
        evidenceBundle,
        now,
      });

      res.locals = res.locals || {};
      res.locals.queryCandidatePlannerCanaryPreflight = preflight;

      const bootstrapMode = bootstrapModeParser(env);
      res.locals.queryCandidatePlannerLiveBootstrapMode = bootstrapMode;

      if (bootstrapMode?.active === true) {
        const authorization = bootstrapAuthorize({
          request: req,
          env,
          featureControl: control,
          legacyPreflight: preflight,
          mode: bootstrapMode,
        });

        res.locals.queryCandidatePlannerLiveBootstrapAuthorization =
          authorization;

        // Patch 15.3.3-B live bootstrap never owns the HTTP response.
        // The Primary payload is returned immediately and unchanged.
        const response = originalJson(primaryPayload);

        if (
          authorization?.allowed !== true ||
          authorization?.runtimeExecutionEligible !== true
        ) {
          const observation = safeRuntimeObservation({
            status: "LIVE_BOOTSTRAP_BLOCKED",
            reason:
              authorization?.reason || "LIVE_BOOTSTRAP_RUNTIME_BLOCKED",
            authorization,
            providerCalls: 0,
          });
          res.locals.queryCandidatePlannerLiveBootstrapObservation =
            observation;
          if (typeof onLiveBootstrapObservation === "function") {
            onLiveBootstrapObservation(observation, { req, res });
          }
          return response;
        }

        const bootstrapTask = Promise.resolve()
          .then(() =>
            bootstrapRunner({
              request: req,
              primaryPayload,
              env,
              featureControl: control,
              legacyPreflight: preflight,
              authorization,
              now,
            }),
          )
          .then((result) => {
            const observation =
              result?.observation ||
              safeRuntimeObservation({
                status: "LIVE_BOOTSTRAP_FALLBACK_SAFE",
                reason: "LIVE_BOOTSTRAP_RUNTIME_RESULT_INVALID",
                authorization,
                providerCalls: 0,
              });

            res.locals.queryCandidatePlannerLiveBootstrapResult = result;
            res.locals.queryCandidatePlannerLiveBootstrapObservation =
              observation;

            if (typeof onLiveBootstrapObservation === "function") {
              onLiveBootstrapObservation(observation, { req, res });
            }

            return result;
          })
          .catch((error) => {
            const observation = safeRuntimeObservation({
              status: "LIVE_BOOTSTRAP_FALLBACK_SAFE",
              reason: String(
                error?.code || "LIVE_BOOTSTRAP_BOUNDARY_FAILED_SAFE",
              ),
              authorization,
              providerCalls: 0,
            });
            res.locals.queryCandidatePlannerLiveBootstrapObservation =
              observation;
            if (typeof onLiveBootstrapObservation === "function") {
              onLiveBootstrapObservation(observation, { req, res });
            }
            return null;
          });

        res.locals.queryCandidatePlannerLiveBootstrapTask = bootstrapTask;
        return response;
      }

      if (!preflight.allowed) {
        const response = originalJson(primaryPayload);
        const canaryObservation = safeBlockedCanaryObservation(preflight);
        res.locals.queryCandidatePlannerCanaryObservation = canaryObservation;
        if (typeof onCanaryObservation === "function") {
          onCanaryObservation(canaryObservation, { req, res });
        }

        const shadowTask = Promise.resolve()
          .then(() =>
            shadowObserve({
              request: req,
              primaryPayload,
              featureControl: control,
              shadowRunner,
              comparator,
            }),
          )
          .then((observation) => {
            res.locals.queryCandidatePlannerShadowObservation = observation;
            if (typeof onObservation === "function") {
              onObservation(observation, { req, res });
            }
            return observation;
          })
          .catch(() => null);
        res.locals.queryCandidatePlannerShadowTask = shadowTask;
        return response;
      }

      responseTask = Promise.resolve()
        .then(() =>
          canaryRunner({
            request: req,
            primaryPayload,
            env,
            featureControl: control,
            evidenceBundle,
            preflight,
            shadowRunner,
            comparator,
            now,
          }),
        )
        .then((result) => {
          const observation = result.observation;
          res.locals.queryCandidatePlannerCanaryResult = result;
          res.locals.queryCandidatePlannerCanaryObservation = observation;
          res.locals.queryCandidatePlannerShadowObservation = observation;
          if (typeof onCanaryObservation === "function") {
            onCanaryObservation(observation, { req, res });
          }
          if (typeof onObservation === "function") {
            onObservation(observation, { req, res });
          }
          return originalJson(result.responsePayload || primaryPayload);
        })
        .catch((error) => {
          const observation = Object.freeze({
            version:
              "query_candidate_planner_internal_allowlist_canary_observation_v1",
            status: "BOUNDARY_FALLBACK_SAFE",
            reason: String(
              error?.code || "INTERNAL_CANARY_BOUNDARY_FAILED_SAFE",
            ),
            responseSource: "PRIMARY",
            merge: Object.freeze({ applied: false }),
            guardrails: Object.freeze({
              primaryFallbackAvailable: true,
              controlledProductionMergeApplied: false,
              productionReadyAssignment: false,
              productionRouteChanged: false,
              failClosed: true,
            }),
            privacy: Object.freeze({
              rawErrorMessageIncluded: false,
              rawIdentityIncluded: false,
              rawPrimaryResponseIncluded: false,
            }),
          });
          res.locals.queryCandidatePlannerCanaryObservation = observation;
          if (typeof onCanaryObservation === "function") {
            onCanaryObservation(observation, { req, res });
          }
          return originalJson(primaryPayload);
        });

      res.locals.queryCandidatePlannerCanaryTask = responseTask;
      return responseTask;
    };

    try {
      const handlerResult = await handler(req, res, next);
      if (responseTask) await responseTask;
      return handlerResult;
    } catch (error) {
      res.json = originalJson;
      throw error;
    }
  };
}

module.exports = Object.freeze({
  BOUNDARY_VERSION,
  defaultCanaryObservationLogger,
  safeBlockedCanaryObservation,
  createQueryCandidatePlannerInternalAllowlistCanaryBoundary,
});