"use strict";
const assert = require("assert");
const { completeSourceCatalog, hash } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, upsertRealShadowFingerprintCapture } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const support = completeSourceCatalog();
const item = support.dataset.cases[0];
assert.throws(() => upsertRealShadowFingerprintCapture({ accuracyDataset: support.dataset, sourceCatalog: support.catalog, ledger: buildRealShadowFingerprintLedgerScaffold(support.dataset, support.catalog), sourceFilePath: support.sources.byCaseId.get(item.caseId), caseId: item.caseId, requestFingerprintSha256: "invalid", uploadFingerprintSha256: hash("b"), captureSource: "INTERNAL_PREVIEW" }), (error) => error.code === "REAL_SHADOW_REQUEST_FINGERPRINT_INVALID");
console.log("PASS query candidate patch15.3.2-B invalid fingerprint fail-closed smoke superseded=B.1");
