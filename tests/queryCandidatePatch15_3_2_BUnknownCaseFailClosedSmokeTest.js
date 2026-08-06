"use strict";
const assert = require("assert");
const { accuracyDataset, hash } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, upsertRealShadowFingerprintCapture } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const dataset = accuracyDataset();
assert.throws(() => upsertRealShadowFingerprintCapture({
  accuracyDataset: dataset,
  ledger: buildRealShadowFingerprintLedgerScaffold(dataset),
  caseId: "not_in_dataset",
  requestFingerprintSha256: hash("a"),
  uploadFingerprintSha256: hash("b"),
  captureSource: "INTERNAL_PREVIEW",
  capturedAt: "2026-08-06T05:00:00.000Z",
  now: Date.parse("2026-08-06T05:30:00.000Z"),
}), (error) => error.code === "REAL_SHADOW_CASE_ID_NOT_IN_ACCURACY_DATASET");
console.log("PASS query candidate patch15.3.2-B unknown case fail-closed smoke");
