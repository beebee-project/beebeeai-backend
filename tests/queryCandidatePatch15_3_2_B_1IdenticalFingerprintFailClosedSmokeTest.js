"use strict";
const assert = require("assert");
const {
  buildRealShadowFingerprintLedgerScaffold,
  upsertRealShadowFingerprintCapture,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const { completeSourceCatalog, hash } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const support = completeSourceCatalog();
const ledger = buildRealShadowFingerprintLedgerScaffold(support.dataset, support.catalog);
const item = support.dataset.cases[0];
assert.throws(
  () => upsertRealShadowFingerprintCapture({
    accuracyDataset: support.dataset,
    sourceCatalog: support.catalog,
    ledger,
    sourceFilePath: support.sources.byCaseId.get(item.caseId),
    caseId: item.caseId,
    requestFingerprintSha256: hash("a"),
    uploadFingerprintSha256: hash("a"),
    captureSource: "INTERNAL_PREVIEW",
  }),
  (error) => error.code === "REAL_SHADOW_REQUEST_UPLOAD_FINGERPRINT_IDENTICAL",
);
console.log("PASS query candidate patch15.3.2-B.1 identical fingerprint fail-closed smoke");
