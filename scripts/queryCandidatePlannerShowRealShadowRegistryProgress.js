#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  validateRealShadowFingerprintLedger,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

try {
  const root = path.resolve(__dirname, "..");
  const ledgerArg = arg("--ledger");
  if (!ledgerArg) throw new Error("--ledger is required");
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
  const ledger = JSON.parse(fs.readFileSync(path.resolve(ledgerArg), "utf8"));
  const result = validateRealShadowFingerprintLedger({
    accuracyDataset,
    ledger,
    requireComplete: false,
  });
  if (!result.valid) {
    result.errors.forEach((error) => console.error(`BLOCKED ${error}`));
    process.exitCode = 2;
  } else {
    for (const item of ledger.cases) {
      const complete =
        /^[a-f0-9]{64}$/i.test(String(item.requestFingerprintSha256 || "")) &&
        /^[a-f0-9]{64}$/i.test(String(item.uploadFingerprintSha256 || "")) &&
        Boolean(item.captureSource) &&
        Boolean(item.capturedAt);
      console.log(`${complete ? "READY" : "PENDING"} ${item.caseId}`);
    }
    console.log(`PROGRESS ${result.completedCount}/${result.expectedCaseCount}`);
    console.log(`REMAINING ${result.remainingCount}`);
    console.log(`COMPLETE ${result.complete}`);
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
