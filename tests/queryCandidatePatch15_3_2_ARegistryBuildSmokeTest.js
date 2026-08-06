"use strict";

const assert = require("assert");
const crypto = require("crypto");
const dataset = require("../evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json");
const {
  buildRealShadowCaseRegistryScaffold,
  buildRealShadowCaseRegistry,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");
const {
  parseRegistry,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceConfig");

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
const result = buildRealShadowCaseRegistry({ accuracyDataset: dataset, draft });
assert.strictEqual(result.valid, true);
assert.strictEqual(result.caseCount, 10);
assert.strictEqual(result.requestFingerprintCount, 10);
assert.strictEqual(result.uploadFingerprintCount, 10);
assert.match(result.registrySha256, /^[a-f0-9]{64}$/);
assert.strictEqual(parseRegistry(JSON.stringify(result.registry)).valid, true);
console.log("PASS query candidate patch15.3.2-A registry build smoke");
