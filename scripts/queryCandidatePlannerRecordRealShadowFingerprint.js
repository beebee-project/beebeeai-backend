#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  upsertRealShadowFingerprintCapture,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

function required(name) {
  const value = arg(name);
  if (!value) throw new Error(`${name} is required`);
  return value;
}

function atomicWrite(filePath, value) {
  const temporary = `${filePath}.${process.pid}.tmp`;
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(temporary, value, { encoding: "utf8", mode: 0o600 });
  fs.renameSync(temporary, filePath);
}

try {
  const root = path.resolve(__dirname, "..");
  const ledgerPath = path.resolve(required("--ledger"));
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
  const ledger = JSON.parse(fs.readFileSync(ledgerPath, "utf8"));
  const result = upsertRealShadowFingerprintCapture({
    accuracyDataset,
    ledger,
    caseId: required("--case-id"),
    requestFingerprintSha256: required("--request-fingerprint"),
    uploadFingerprintSha256: required("--upload-fingerprint"),
    captureSource: required("--capture-source"),
    capturedAt: arg("--captured-at", new Date().toISOString()),
    expectedColdCostMicrousd: Number(arg("--expected-cold-cost-microusd", "0")),
    modelId: arg("--model-id", "semantic_profiler_default"),
    operatorNote: arg("--operator-note", "actual internal request captured"),
  });
  atomicWrite(ledgerPath, `${JSON.stringify(result.ledger, null, 2)}\n`);
  console.log(`PASS recorded case=${result.recordedCaseId}`);
  console.log(`PROGRESS ${result.completedCount}/10`);
  console.log(`REMAINING ${result.remainingCount}`);
  console.log(`LEDGER_SHA256 ${result.ledgerSha256}`);
  console.log("RAW_FINGERPRINTS_LOGGED false");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
