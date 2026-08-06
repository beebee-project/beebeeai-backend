#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  finalizeRealShadowCaseRegistry,
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
  return path.resolve(value);
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

try {
  const root = path.resolve(__dirname, "..");
  const ledgerPath = required("--ledger");
  const outputPath = privateOutput(
    "--output",
    "queryCandidatePlannerRealShadowCaseRegistry.private.json",
  );
  const railwayOutputPath = privateOutput(
    "--railway-output",
    "queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt",
  );
  const summaryOutputPath = privateOutput(
    "--summary-output",
    "queryCandidatePlannerRealShadowCaseRegistry.summary.private.json",
  );
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
  const result = finalizeRealShadowCaseRegistry({ accuracyDataset, ledger });
  if (!result.valid) {
    for (const error of result.errors) console.error(`BLOCKED ${error}`);
    process.exitCode = 2;
  } else {
    fs.writeFileSync(
      outputPath,
      `${JSON.stringify(result.registry, null, 2)}\n`,
      { encoding: "utf8", mode: 0o600 },
    );
    fs.writeFileSync(
      railwayOutputPath,
      `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON=${JSON.stringify(result.registry)}\n`,
      { encoding: "utf8", mode: 0o600 },
    );
    fs.writeFileSync(
      summaryOutputPath,
      `${JSON.stringify(
        {
          version: result.version,
          decision: "REAL_SHADOW_REGISTRY_FINALIZATION_PASS",
          registryId: result.registry.registryId,
          registrySha256: result.registrySha256,
          ledgerSha256: result.ledgerSha256,
          caseCount: result.caseCount,
          requestFingerprintCount: result.requestFingerprintCount,
          uploadFingerprintCount: result.uploadFingerprintCount,
          source: result.source,
          actualTraffic: true,
          synthetic: false,
          rawIdentityIncluded: false,
          collectorEnabledByThisOperation: false,
          internalCanaryEnabledByThisOperation: false,
          productionPromotionAuthorized: false,
        },
        null,
        2,
      )}\n`,
      { encoding: "utf8", mode: 0o600 },
    );
    console.log(
      `PASS real shadow registry finalized sha256=${result.registrySha256} cases=${result.caseCount}`,
    );
    console.log(`OUTPUT ${outputPath}`);
    console.log(`RAILWAY ${railwayOutputPath}`);
    console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
    console.log("COLLECTOR_ENABLED_BY_THIS_OPERATION false");
    console.log("INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
