#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  upsertRealShadowFingerprintCapture,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const {
  buildExpectedRejectionAttestation,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceFoundation");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}
function required(name) {
  const value = arg(name);
  if (!value) throw new Error(`${name} is required`);
  return value;
}
function privateOutput(name, fallback) {
  const output = path.resolve(arg(name, fallback));
  if (!/\.private\./i.test(path.basename(output))) {
    const error = new Error(`${name} must use a .private. filename`);
    error.code = "REAL_SHADOW_PRIVATE_OUTPUT_NAME_REQUIRED";
    throw error;
  }
  return output;
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
  const sourceCatalogPath = path.resolve(required("--source-catalog"));
  const attestationPath = privateOutput(
    "--attestation-output",
    "queryCandidatePlannerRealShadowExpectedRejectionAttestation.private.json",
  );
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(root, "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json"),
      "utf8",
    ),
  );
  const sourceCatalog = JSON.parse(fs.readFileSync(sourceCatalogPath, "utf8"));
  const ledger = JSON.parse(fs.readFileSync(ledgerPath, "utf8"));
  const caseId = required("--case-id");
  const requestFingerprintSha256 = required("--request-fingerprint");
  const uploadFingerprintSha256 = required("--upload-fingerprint");
  const captureSource = arg("--capture-source", "INTERNAL_PREVIEW");
  const observedAt = arg("--observed-at", new Date().toISOString());

  const recorded = upsertRealShadowFingerprintCapture({
    accuracyDataset,
    sourceCatalog,
    ledger,
    sourceFilePath: required("--source-file"),
    caseId,
    requestFingerprintSha256,
    uploadFingerprintSha256,
    captureSource,
    capturedAt: observedAt,
    expectedColdCostMicrousd: Number(arg("--expected-cold-cost-microusd", "0")),
    modelId: arg("--model-id", "semantic_profiler_default"),
    operatorNote: "actual source-bound expected rejection verified",
  });

  const attested = buildExpectedRejectionAttestation({
    accuracyDataset,
    sourceCatalog,
    ledger: recorded.ledger,
    caseId,
    requestFingerprintSha256,
    uploadFingerprintSha256,
    observationStatus: required("--observation-status"),
    observationReason: required("--observation-reason"),
    shadowAccepted: Number(required("--shadow-accepted")),
    captureSource,
    observedAt,
  });

  atomicWrite(ledgerPath, `${JSON.stringify(recorded.ledger, null, 2)}\n`);
  atomicWrite(attestationPath, `${JSON.stringify(attested.attestation, null, 2)}\n`);

  console.log(`PASS recorded expected-rejection case=${recorded.recordedCaseId}`);
  console.log(`PROGRESS ${recorded.completedCount}/10`);
  console.log(`REMAINING ${recorded.remainingCount}`);
  console.log(`LEDGER_SHA256 ${recorded.ledgerSha256}`);
  console.log(`ATTESTATION_SHA256 ${attested.attestationSha256}`);
  console.log("EXPECTED_REJECTION_VERIFIED true");
  console.log("SHADOW_ACCEPTED 0");
  console.log("RAW_FINGERPRINTS_LOGGED false");
  console.log("SOURCE_PATH_LOGGED false");
  console.log("RAW_FILE_CONTENT_LOGGED false");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
