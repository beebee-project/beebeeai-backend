"use strict";

const assert = require("assert");
const crypto = require("crypto");
const dataset = require("../evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json");
const {
  buildRealShadowCaseRegistryScaffold,
  buildRealShadowCaseRegistry,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

const hash = (value) =>
  crypto.createHash("sha256").update(String(value)).digest("hex");
const sameRequest = hash("same-request");
const scaffold = buildRealShadowCaseRegistryScaffold(dataset);
const draft = {
  ...scaffold,
  cases: scaffold.cases.map((item) => ({
    ...item,
    requestFingerprintSha256: sameRequest,
    uploadFingerprintSha256: hash(`${item.caseId}:upload`),
  })),
};
const result = buildRealShadowCaseRegistry({ accuracyDataset: dataset, draft });
assert.strictEqual(result.valid, false);
assert(result.errors.some((error) => error.includes("duplicate request fingerprint")));
console.log("PASS query candidate patch15.3.2-A duplicate fingerprint fail-closed smoke");
