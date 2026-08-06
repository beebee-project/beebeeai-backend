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
  secret: Buffer.alloc(48, 0x42).toString("base64url"),
  allowlistSha256: hash("internal-account"),
  env: {},
});
assert.strictEqual(result.ready, true);
assert.strictEqual(result.collectorConfigurationWouldBeValid, true);
assert.strictEqual(result.collectorEnabledByThisOperation, false);
assert.strictEqual(result.internalCanaryEnabledByThisOperation, false);
assert.strictEqual(result.productionPromotionAuthorized, false);
assert.strictEqual(result.rawSecretIncluded, false);
console.log("PASS query candidate patch15.3.2-A preparation verify smoke");
