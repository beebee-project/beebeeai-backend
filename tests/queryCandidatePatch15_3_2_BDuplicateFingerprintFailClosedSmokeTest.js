"use strict";
const assert = require("assert");
const { accuracyDataset, hash } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, upsertRealShadowFingerprintCapture } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const dataset = accuracyDataset();
let ledger = buildRealShadowFingerprintLedgerScaffold(dataset);
ledger = upsertRealShadowFingerprintCapture({
  accuracyDataset: dataset,
  ledger,
  caseId: dataset.cases[0].caseId,
  requestFingerprintSha256: hash("a"),
  uploadFingerprintSha256: hash("b"),
  captureSource: "API_SHADOW_OBSERVATION",
  capturedAt: "2026-08-06T05:00:00.000Z",
  now: Date.parse("2026-08-06T05:30:00.000Z"),
}).ledger;
assert.throws(() => upsertRealShadowFingerprintCapture({
  accuracyDataset: dataset,
  ledger,
  caseId: dataset.cases[1].caseId,
  requestFingerprintSha256: hash("a"),
  uploadFingerprintSha256: hash("c"),
  captureSource: "INTERNAL_PREVIEW",
  capturedAt: "2026-08-06T05:01:00.000Z",
  now: Date.parse("2026-08-06T05:30:00.000Z"),
}), (error) => error.code === "REAL_SHADOW_REQUEST_FINGERPRINT_DUPLICATE");
console.log("PASS query candidate patch15.3.2-B duplicate fingerprint fail-closed smoke");
