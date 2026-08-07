"use strict";
const assert = require("assert");
const {
  finalizeRealShadowCaseRegistry,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const { completeLedger } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const support = completeLedger();
const input = {
  accuracyDataset: support.dataset,
  sourceCatalog: support.catalog,
  ledger: support.ledger,
  now: Date.parse("2026-08-06T08:30:00.000Z"),
};
const a = finalizeRealShadowCaseRegistry(input);
const b = finalizeRealShadowCaseRegistry(input);
assert.strictEqual(a.registrySha256, b.registrySha256);
assert.strictEqual(a.sourceCatalogSha256, b.sourceCatalogSha256);
console.log("PASS query candidate patch15.3.2-B.1 determinism smoke");
