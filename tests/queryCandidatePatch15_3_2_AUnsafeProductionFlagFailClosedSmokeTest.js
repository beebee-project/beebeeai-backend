"use strict";

const assert = require("assert");
const crypto = require("crypto");
const dataset = require("../evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json");
const {
  buildRealShadowCaseRegistryScaffold,
  buildRealShadowCaseRegistry,
  verifyRealShadowPreparation,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

const hash = (value) =>
  crypto.createHash("sha256").update(String(value)).digest("hex");
const scaffold = buildRealShadowCaseRegistryScaffold(dataset);
const draft = {
  ...scaffold,
  cases: scaffold.cases.map((item) => ({
    ...item,
    requestFingerprintSha256: hash(`${item.caseId}:request`),
    uploadFingerprintSha256: hash(`${item.caseId}:upload`),
  })),
};
const registry = buildRealShadowCaseRegistry({
  accuracyDataset: dataset,
  draft,
}).registry;
const result = verifyRealShadowPreparation({
  accuracyDataset: dataset,
  registry,
  secret: Buffer.alloc(48, 0x43).toString("base64url"),
  allowlistSha256: hash("internal-account"),
  env: {
    QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED: "1",
    QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH: "0",
  },
});
assert.strictEqual(result.ready, false);
assert(result.errors.some((error) => error.includes("productionDisabled")));
assert(result.errors.some((error) => error.includes("productionKillSwitchActive")));
assert.strictEqual(result.productionPromotionAuthorized, false);
console.log("PASS query candidate patch15.3.2-A unsafe flag fail-closed smoke");
