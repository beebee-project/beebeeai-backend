"use strict";
const assert = require("assert");
const { completeSourceCatalog } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, finalizeRealShadowCaseRegistry } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const support = completeSourceCatalog();
const result = finalizeRealShadowCaseRegistry({ accuracyDataset: support.dataset, sourceCatalog: support.catalog, ledger: buildRealShadowFingerprintLedgerScaffold(support.dataset, support.catalog) });
assert.strictEqual(result.valid, false);
console.log("PASS query candidate patch15.3.2-B incomplete finalization fail-closed smoke superseded=B.1");
