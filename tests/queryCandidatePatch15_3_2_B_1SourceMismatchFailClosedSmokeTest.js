"use strict";
const assert = require("assert");
const {
  buildRealShadowFingerprintLedgerScaffold,
  upsertRealShadowFingerprintCapture,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const { completeSourceCatalog, hash } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const support = completeSourceCatalog();
const ledger = buildRealShadowFingerprintLedgerScaffold(support.dataset, support.catalog);
const first = support.dataset.cases[0];
const wrong = support.dataset.cases[1];
assert.throws(
  () => upsertRealShadowFingerprintCapture({
    accuracyDataset: support.dataset,
    sourceCatalog: support.catalog,
    ledger,
    sourceFilePath: support.sources.byCaseId.get(wrong.caseId),
    caseId: first.caseId,
    requestFingerprintSha256: hash("a"),
    uploadFingerprintSha256: hash("b"),
    captureSource: "INTERNAL_PREVIEW",
  }),
  (error) => error.code === "REAL_SHADOW_CAPTURE_SOURCE_ARTIFACT_MISMATCH",
);
console.log("PASS query candidate patch15.3.2-B.1 source mismatch fail-closed smoke");
