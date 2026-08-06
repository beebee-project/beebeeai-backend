"use strict";
const assert = require("assert");
const { completeLedger } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { finalizeRealShadowCaseRegistry } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const { dataset, ledger } = completeLedger();
const result = finalizeRealShadowCaseRegistry({
  accuracyDataset: dataset,
  ledger,
  now: Date.parse("2026-08-06T05:30:00.000Z"),
});
assert.strictEqual(result.valid, true);
assert.strictEqual(result.caseCount, 10);
assert.strictEqual(result.requestFingerprintCount, 10);
assert.strictEqual(result.uploadFingerprintCount, 10);
assert.strictEqual(result.actualTraffic, true);
assert.strictEqual(result.synthetic, false);
assert.strictEqual(result.rawIdentityIncluded, false);
console.log("PASS query candidate patch15.3.2-B complete registry finalization smoke");
