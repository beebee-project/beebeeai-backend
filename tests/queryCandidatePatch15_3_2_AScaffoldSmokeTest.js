"use strict";

const assert = require("assert");
const dataset = require("../evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json");
const {
  buildRealShadowCaseRegistryScaffold,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

const scaffold = buildRealShadowCaseRegistryScaffold(dataset);
assert.strictEqual(scaffold.cases.length, dataset.cases.length);
assert.deepStrictEqual(
  scaffold.cases.map((item) => item.caseId),
  dataset.cases.map((item) => item.caseId),
);
assert(scaffold.cases.every((item) => item.requestFingerprintSha256 === ""));
assert(scaffold.cases.every((item) => item.uploadFingerprintSha256 === ""));
assert.strictEqual(scaffold.actualTrafficOnly, true);
assert.strictEqual(scaffold.syntheticFingerprintForbidden, true);
console.log("PASS query candidate patch15.3.2-A scaffold smoke");
