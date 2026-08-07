"use strict";
const assert = require("assert");
const {
  finalizeRealShadowCaseRegistry,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const { completeLedger } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const support = completeLedger();
const result = finalizeRealShadowCaseRegistry({
  accuracyDataset: support.dataset,
  sourceCatalog: support.catalog,
  ledger: support.ledger,
  now: Date.parse("2026-08-06T08:30:00.000Z"),
});
assert.strictEqual(result.valid, true);
assert.strictEqual(result.caseCount, 10);
assert(/^[a-f0-9]{64}$/.test(result.sourceCatalogSha256));
console.log("PASS query candidate patch15.3.2-B.1 complete finalization smoke");
