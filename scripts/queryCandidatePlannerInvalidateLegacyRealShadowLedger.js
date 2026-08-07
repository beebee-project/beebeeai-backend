#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  LEGACY_LEDGER_VERSION,
  buildRealShadowFingerprintLedgerScaffold,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}
function required(name) {
  const value = arg(name);
  if (!value) throw new Error(`${name} is required`);
  return path.resolve(value);
}

try {
  const root = path.resolve(__dirname, "..");
  const oldLedgerPath = required("--legacy-ledger");
  const sourceCatalogPath = required("--source-catalog");
  const outputPath = required("--output");
  if (!/\.private\./i.test(path.basename(outputPath))) {
    const error = new Error("--output must use a .private. filename");
    error.code = "REAL_SHADOW_PRIVATE_OUTPUT_NAME_REQUIRED";
    throw error;
  }
  const oldLedger = JSON.parse(fs.readFileSync(oldLedgerPath, "utf8"));
  if (oldLedger.version !== LEGACY_LEDGER_VERSION) {
    const error = new Error("legacy v1 ledger required");
    error.code = "REAL_SHADOW_LEGACY_LEDGER_REQUIRED";
    throw error;
  }
  const backupPath = `${oldLedgerPath}.invalidated-${Date.now()}.backup.private.json`;
  fs.copyFileSync(oldLedgerPath, backupPath);
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(root, "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json"),
      "utf8",
    ),
  );
  const sourceCatalog = JSON.parse(fs.readFileSync(sourceCatalogPath, "utf8"));
  const ledger = buildRealShadowFingerprintLedgerScaffold(
    accuracyDataset,
    sourceCatalog,
    {
      registryId: arg("--registry-id", "internal_real_shadow_2026_08_v2"),
    },
  );
  fs.writeFileSync(outputPath, `${JSON.stringify(ledger, null, 2)}\n`, {
    encoding: "utf8",
    mode: 0o600,
  });
  console.log(`PASS legacy ledger invalidated backup=${backupPath}`);
  console.log(`OUTPUT ${outputPath}`);
  console.log("PROGRESS 0/10");
  console.log("LEGACY_CAPTURES_PRESERVED false");
  console.log("RERUN_ALL_CASES_REQUIRED true");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
