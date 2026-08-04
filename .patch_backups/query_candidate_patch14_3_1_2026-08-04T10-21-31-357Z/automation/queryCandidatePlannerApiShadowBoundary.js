const {
  getQueryCandidatePlannerFeatureControl,
} = require("./queryCandidatePlannerFeatureControlRuntime");
const {
  observeQueryCandidatePlannerApiShadow,
} = require("./queryCandidatePlannerApiShadowService");

const BOUNDARY_VERSION = "query_candidate_planner_api_shadow_boundary_v1";

function defaultObservationLogger(observation = {}) {
  if (observation.status === "BLOCKED") return;
  const summary = {
    version: observation.version,
    status: observation.status,
    reason: observation.reason,
    requestFingerprintSha256: observation.requestFingerprintSha256 || "",
    latencyMs: observation.latencyMs || 0,
    comparisonVerdict: observation.comparison?.verdict || "NOT_AVAILABLE",
    primaryCount: observation.comparison?.counts?.primary || 0,
    shadowCount: observation.comparison?.counts?.shadow || 0,
    providerCallCount: observation.shadow?.providerCallCount || 0,
    productionCandidateMerge: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
  };
  console.info("[query-candidate-api-shadow]", summary);
}

function createQueryCandidatePlannerApiShadowBoundary({
  handler,
  featureControl = null,
  shadowRunner,
  comparator,
  observe = observeQueryCandidatePlannerApiShadow,
  onObservation = defaultObservationLogger,
  timeoutMs,
} = {}) {
  if (typeof handler !== "function") {
    throw new TypeError("API shadow boundary handler must be a function");
  }

  return async function queryCandidatePlannerApiShadowBoundary(req, res, next) {
    const originalJson = res.json.bind(res);
    let responseObserved = false;

    res.json = function apiShadowObservedJson(primaryPayload) {
      const response = originalJson(primaryPayload);
      if (responseObserved) return response;
      responseObserved = true;

      const control =
        featureControl || getQueryCandidatePlannerFeatureControl();
      const task = Promise.resolve()
        .then(() =>
          observe({
            request: req,
            primaryPayload,
            featureControl: control,
            shadowRunner,
            comparator,
            timeoutMs,
          }),
        )
        .then((observation) => {
          res.locals = res.locals || {};
          res.locals.queryCandidatePlannerShadowObservation = observation;
          if (typeof onObservation === "function") {
            onObservation(observation, { req, res });
          }
          return observation;
        })
        .catch((error) => {
          const observation = Object.freeze({
            version: "query_candidate_planner_api_shadow_observation_v1",
            status: "BOUNDARY_FAILED_SAFE",
            reason: String(error?.code || "BOUNDARY_OBSERVER_FAILED"),
            guardrails: Object.freeze({
              primaryResponseAuthority: true,
              responsePayloadMutation: false,
              responseHeaderMutation: false,
              responseStatusMutation: false,
              productionCandidateMerge: false,
              productionReadyAssignment: false,
              productionRouteChanged: false,
            }),
          });
          res.locals = res.locals || {};
          res.locals.queryCandidatePlannerShadowObservation = observation;
          return observation;
        });

      res.locals = res.locals || {};
      res.locals.queryCandidatePlannerShadowTask = task;
      return response;
    };

    try {
      return await handler(req, res, next);
    } catch (error) {
      res.json = originalJson;
      throw error;
    }
  };
}

module.exports = Object.freeze({
  BOUNDARY_VERSION,
  defaultObservationLogger,
  createQueryCandidatePlannerApiShadowBoundary,
});
