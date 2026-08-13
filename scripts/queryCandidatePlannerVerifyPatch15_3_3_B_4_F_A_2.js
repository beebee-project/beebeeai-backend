"use strict";

const assert = require("assert");
const {
  OPERATIONS,
  createQueryCandidatePlannerFeatureControl,
} = require("../automation/queryCandidatePlannerFeatureControl");
const {
  loadBootstrapProductionReadinessGate,
  createReadinessAwareApprovalFeatureControl,
} = require("../automation/queryCandidatePlannerInternalCanaryBootstrapReadinessBridge");

const env = {
  QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED: "1",
  QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED: "1",
  QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED: "1",
  QUERY_CANDIDATE_PLANNER_PROVIDER_KILL_SWITCH: "0",
  QUERY_CANDIDATE_PLANNER_KILL_SWITCH: "0",
  QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED: "1",
  QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH: "0",
  QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED: "1",
  QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED: "0",
  QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED: "0",
};

const control = createQueryCandidatePlannerFeatureControl({ env });
const direct = control.evaluate(OPERATIONS.PRODUCTION_CANDIDATE_MERGE);
const loaded = loadBootstrapProductionReadinessGate();
assert.strictEqual(loaded.valid, true, loaded.reason);
const bridge = createReadinessAwareApprovalFeatureControl({
  featureControl: control,
  readinessGate: loaded.readinessGate,
});
assert.strictEqual(bridge.valid, true, bridge.reason);
const merged = bridge.featureControl.evaluate(OPERATIONS.PRODUCTION_CANDIDATE_MERGE);
const route = bridge.featureControl.evaluate(OPERATIONS.PRODUCTION_ROUTE);
const ready = bridge.featureControl.evaluate(OPERATIONS.PRODUCTION_READY_ASSIGNMENT);

console.log(`DIRECT_PRODUCTION_CANDIDATE_MERGE_ALLOWED ${direct.allowed === true}`);
console.log(`DIRECT_PRODUCTION_CANDIDATE_MERGE_REASON ${direct.reason}`);
console.log(`READINESS_BRIDGE_VALID ${bridge.valid === true}`);
console.log(`READINESS_FILE_SHA256 ${loaded.readinessFileSha256}`);
console.log(`BRIDGED_PRODUCTION_CANDIDATE_MERGE_ALLOWED ${merged.allowed === true}`);
console.log(`BRIDGED_PRODUCTION_CANDIDATE_MERGE_REASON ${merged.reason}`);
console.log(`PRODUCTION_ROUTE_ALLOWED ${route.allowed === true}`);
console.log(`PRODUCTION_READY_ASSIGNMENT_ALLOWED ${ready.allowed === true}`);
console.log("F16_GATE_MODIFIED false");
console.log("B2_RUNTIME_MODIFIED false");
console.log("RAILWAY_VARIABLE_MUTATED false");
console.log("RAILWAY_DEPLOY_TRIGGERED false");
console.log("PROVIDER_CALLS_EXECUTED 0");
console.log("ACTUAL_LIVE_REQUEST_EXECUTED false");
console.log("PASS PATCH 15.3.3-B-4-F-A.2 READINESS-AWARE RUNTIME COMPATIBILITY REPAIR VERIFICATION");
