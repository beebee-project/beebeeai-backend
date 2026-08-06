"use strict";
const assert = require("assert");
const { accuracyDataset, hash } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, upsertRealShadowFingerprintCapture } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const dataset = accuracyDataset();
assert.throws(() => upsertRealShadowFingerprintCapture({
  accuracyDataset: dataset,
  ledger: buildRealShadowFingerprintLedgerScaffold(dataset),
  caseId: dataset.cases[0].caseId,
  requestFingerprintSha256: "invalid",
  uploadFingerprintSha256: hash("b"),
  captureSource: "INTERNAL_PREVIEW",
  capturedAt: "2026-08-06T05:00:00.000Z",
  now: Date.parse("2026-08-06T05:30:00.000Z"),
}), (error) => error.code === "REAL_SHADOW_REQUEST_FINGERPRINT_INVALID");
console.log("PASS query candidate patch15.3.2-B invalid fingerprint fail-closed smoke");
