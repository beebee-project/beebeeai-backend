"use strict";

const fs = require("fs");
const path = require("path");
const {
  buildRealShadowFingerprintLedgerScaffold,
  upsertRealShadowFingerprintCapture,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");

function accuracyDataset() {
  return JSON.parse(
    fs.readFileSync(
      path.join(
        __dirname,
        "../evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
}

function hash(character) {
  return String(character).repeat(64).slice(0, 64);
}

function completeLedger() {
  const dataset = accuracyDataset();
  let ledger = buildRealShadowFingerprintLedgerScaffold(dataset);
  dataset.cases.forEach((item, index) => {
    const requestCharacter = (index + 1).toString(16).slice(-1);
    const uploadCharacter = (index + 11).toString(16).slice(-1);
    ledger = upsertRealShadowFingerprintCapture({
      accuracyDataset: dataset,
      ledger,
      caseId: item.caseId,
      requestFingerprintSha256: hash(requestCharacter),
      uploadFingerprintSha256: hash(uploadCharacter),
      captureSource: index % 2 === 0
        ? "API_SHADOW_OBSERVATION"
        : "INTERNAL_PREVIEW",
      capturedAt: "2026-08-06T05:00:00.000Z",
      now: Date.parse("2026-08-06T05:30:00.000Z"),
    }).ledger;
  });
  return { dataset, ledger };
}

module.exports = Object.freeze({
  accuracyDataset,
  hash,
  completeLedger,
});
