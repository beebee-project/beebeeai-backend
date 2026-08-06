"use strict";
const assert = require("assert");
const { accuracyDataset, hash } = require("./queryCandidatePatch15_3_2_BTestSupport");
const { buildRealShadowFingerprintLedgerScaffold, upsertRealShadowFingerprintCapture } = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const dataset = accuracyDataset();
const ledger = buildRealShadowFingerprintLedgerScaffold(dataset);
const result = upsertRealShadowFingerprintCapture({
  accuracyDataset: dataset,
  ledger,
  caseId: dataset.cases[0].caseId,
  requestFingerprintSha256: hash("a"),
  uploadFingerprintSha256: hash("b"),
  captureSource: "API_SHADOW_OBSERVATION",
  capturedAt: "2026-08-06T05:00:00.000Z",
  now: Date.parse("2026-08-06T05:30:00.000Z"),
});
assert.strictEqual(result.completedCount, 1);
assert.strictEqual(result.remainingCount, 9);
assert.strictEqual(result.ledger.cases[0].actualTraffic, true);
assert.strictEqual(result.ledger.cases[0].synthetic, false);
console.log("PASS query candidate patch15.3.2-B actual fingerprint record smoke");
