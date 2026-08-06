"use strict";
const assert = require("assert");
const { accuracyDataset } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, finalizeRealShadowCaseRegistry } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const dataset = accuracyDataset();
const result = finalizeRealShadowCaseRegistry({
  accuracyDataset: dataset,
  ledger: buildRealShadowFingerprintLedgerScaffold(dataset),
  now: Date.parse("2026-08-06T05:30:00.000Z"),
});
assert.strictEqual(result.valid, false);
assert.strictEqual(result.registry, null);
console.log("PASS query candidate patch15.3.2-B incomplete finalization fail-closed smoke");
