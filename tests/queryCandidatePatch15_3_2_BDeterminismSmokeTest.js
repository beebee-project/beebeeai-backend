"use strict";
const assert = require("assert");
const { completeLedger } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { finalizeRealShadowCaseRegistry } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const support = completeLedger();
const options = { accuracyDataset: support.dataset, sourceCatalog: support.catalog, ledger: support.ledger, now: Date.parse("2026-08-06T08:30:00.000Z") };
assert.deepStrictEqual(finalizeRealShadowCaseRegistry(options), finalizeRealShadowCaseRegistry(options));
console.log("PASS query candidate patch15.3.2-B deterministic finalization smoke superseded=B.1");
