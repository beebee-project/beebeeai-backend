#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  buildRealShadowFingerprintLedgerScaffold,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}
function required(name) {
  const value = arg(name);
  if (!value) throw new Error(`${name} is required`);
  return value;
}

try {
  const root = path.resolve(__dirname, "..");
  const dataset = JSON.parse(
    fs.readFileSync(
      path.join(root, "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json"),
      "utf8",
    ),
  );
  const sourceCatalog = JSON.parse(
    fs.readFileSync(path.resolve(required("--source-catalog")), "utf8"),
  );
  const output = path.resolve(
    arg("--output", "queryCandidatePlannerRealShadowFingerprintLedger.private.json"),
  );
  if (!/\.private\./i.test(path.basename(output))) {
    const error = new Error("ledger output must use a .private. filename");
    error.code = "REAL_SHADOW_PRIVATE_OUTPUT_NAME_REQUIRED";
    throw error;
  }
  const ledger = buildRealShadowFingerprintLedgerScaffold(
    dataset,
    sourceCatalog,
    {
      registryId: arg("--registry-id", "internal_real_shadow_2026_08_v2"),
    },
  );
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, `${JSON.stringify(ledger, null, 2)}\n`, {
    encoding: "utf8",
    mode: 0o600,
  });
  console.log(
    `PASS source-bound fingerprint ledger scaffold cases=${ledger.cases.length} output=${output}`,
  );
  console.log(`SOURCE_CATALOG_SHA256 ${ledger.sourceCatalogSha256}`);
  console.log("LEGACY_CAPTURES_PRESERVED false");
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
