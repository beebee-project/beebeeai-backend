"use strict";

const assert = require("assert");
const dataset = require("../evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json");
const {
  buildRealShadowCaseRegistryScaffold,
  buildRealShadowCaseRegistry,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

const draft = buildRealShadowCaseRegistryScaffold(dataset);
const result = buildRealShadowCaseRegistry({ accuracyDataset: dataset, draft });
assert.strictEqual(result.valid, false);
assert(result.errors.some((error) => error.includes("actual request fingerprint required")));
assert(result.errors.some((error) => error.includes("actual upload fingerprint required")));
assert.strictEqual(result.registry, null);
console.log("PASS query candidate patch15.3.2-A incomplete fail-closed smoke");
