#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  buildRealShadowFingerprintLedgerScaffold,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

try {
  const root = path.resolve(__dirname, "..");
  const dataset = JSON.parse(
    fs.readFileSync(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
  const output = path.resolve(
    arg(
      "--output",
      "queryCandidatePlannerRealShadowFingerprintLedger.private.json",
    ),
  );
  const registryId = arg("--registry-id", "internal_real_shadow_2026_08_v1");
  const ledger = buildRealShadowFingerprintLedgerScaffold(dataset, {
    registryId,
  });
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, `${JSON.stringify(ledger, null, 2)}\n`, "utf8");
  console.log(
    `PASS fingerprint ledger scaffold cases=${ledger.cases.length} output=${output}`,
  );
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
