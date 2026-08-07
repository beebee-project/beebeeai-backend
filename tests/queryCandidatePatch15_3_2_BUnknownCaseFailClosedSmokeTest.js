"use strict";
const assert = require("assert");
const { completeSourceCatalog, hash } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, upsertRealShadowFingerprintCapture } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const support = completeSourceCatalog();
assert.throws(() => upsertRealShadowFingerprintCapture({ accuracyDataset: support.dataset, sourceCatalog: support.catalog, ledger: buildRealShadowFingerprintLedgerScaffold(support.dataset, support.catalog), sourceFilePath: support.sources.byCaseId.get(support.dataset.cases[0].caseId), caseId: "not_in_dataset", requestFingerprintSha256: hash("a"), uploadFingerprintSha256: hash("b"), captureSource: "INTERNAL_PREVIEW" }), (error) => error.code === "REAL_SHADOW_CASE_ID_NOT_IN_ACCURACY_DATASET");
console.log("PASS query candidate patch15.3.2-B unknown case fail-closed smoke superseded=B.1");
