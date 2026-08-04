"use strict";

const {
  createQueryCandidatePlannerFeatureControl,
} = require("./queryCandidatePlannerFeatureControl");

const RUNTIME_VERSION =
  "query_candidate_planner_feature_control_runtime_v1";

let runtimeControl = null;

function getQueryCandidatePlannerFeatureControl() {
  if (!runtimeControl) {
    runtimeControl = createQueryCandidatePlannerFeatureControl({
      env: process.env,
    });
  }
  return runtimeControl;
}

function activateQueryCandidatePlannerKillSwitch(options = {}) {
  return getQueryCandidatePlannerFeatureControl().activateKillSwitch(options);
}

function releaseQueryCandidatePlannerRuntimeKillSwitch(options = {}) {
  return getQueryCandidatePlannerFeatureControl().releaseRuntimeKillSwitch(
    options,
  );
}

function getQueryCandidatePlannerFeatureControlSnapshot() {
  return getQueryCandidatePlannerFeatureControl().snapshot();
}

function resetQueryCandidatePlannerFeatureControlForTests({
  control = null,
} = {}) {
  runtimeControl = control;
  return runtimeControl;
}

module.exports = Object.freeze({
  RUNTIME_VERSION,
  getQueryCandidatePlannerFeatureControl,
  activateQueryCandidatePlannerKillSwitch,
  releaseQueryCandidatePlannerRuntimeKillSwitch,
  getQueryCandidatePlannerFeatureControlSnapshot,
  resetQueryCandidatePlannerFeatureControlForTests,
});
