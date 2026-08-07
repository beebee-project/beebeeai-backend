#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  evaluateRealShadowEvidenceFoundation,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceFoundation");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}
function args(name) {
  const values = [];
  for (let index = 0; index < process.argv.length; index += 1) {
    if (process.argv[index] === name && process.argv[index + 1]) {
      values.push(process.argv[index + 1]);
    }
  }
  return values;
}
function requiredPath(name) {
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
function writePrivate(filePath, value) {
  fs.writeFileSync(filePath, value, { encoding: "utf8", mode: 0o600 });
}

try {
  const root = path.resolve(__dirname, "..");
  const ledgerPath = requiredPath("--ledger");
  const sourceCatalogPath = requiredPath("--source-catalog");
  const attestationPaths = args("--expected-rejection-attestation").map((item) =>
    path.resolve(item),
  );
  if (attestationPaths.length === 0) {
    throw new Error("--expected-rejection-attestation is required");
  }

  const registryOutputPath = privateOutput(
    "--registry-output",
    "queryCandidatePlannerRealShadowCaseRegistry.private.json",
  );
  const railwayOutputPath = privateOutput(
    "--railway-output",
    "queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt",
  );
  const registrySummaryPath = privateOutput(
    "--registry-summary-output",
    "queryCandidatePlannerRealShadowCaseRegistry.summary.private.json",
  );
  const foundationSummaryPath = privateOutput(
    "--foundation-summary-output",
    "queryCandidatePlannerRealShadowEvidenceFoundation.summary.private.json",
  );

  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(root, "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json"),
      "utf8",
    ),
  );
  const sourceCatalog = JSON.parse(fs.readFileSync(sourceCatalogPath, "utf8"));
  const ledger = JSON.parse(fs.readFileSync(ledgerPath, "utf8"));
  const expectedRejectionAttestations = attestationPaths.map((item) =>
    JSON.parse(fs.readFileSync(item, "utf8")),
  );

  const result = evaluateRealShadowEvidenceFoundation({
    accuracyDataset,
    sourceCatalog,
    ledger,
    expectedRejectionAttestations,
  });

  if (!result.valid) {
    result.errors.forEach((error) => console.error(`BLOCKED ${error}`));
    process.exitCode = 2;
  } else {
    writePrivate(registryOutputPath, `${JSON.stringify(result.registry, null, 2)}\n`);
    writePrivate(
      railwayOutputPath,
      `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON=${JSON.stringify(result.registry)}\n`,
    );
    writePrivate(
      registrySummaryPath,
      `${JSON.stringify(
        {
          decision: "REAL_SHADOW_REGISTRY_FINALIZATION_PASS",
          registrySha256: result.registrySha256,
          sourceCatalogSha256: result.summary.sourceCatalogSha256,
          ledgerSha256: result.summary.ledgerSha256,
          caseCount: result.summary.completedCaseCount,
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
    );
    writePrivate(
      foundationSummaryPath,
      `${JSON.stringify(result.summary, null, 2)}\n`,
    );

    console.log("PASS phase 15.3-A real shadow evidence foundation finalized");
    console.log(`REGISTRY_SHA256 ${result.registrySha256}`);
    console.log(`SOURCE_CATALOG_SHA256 ${result.summary.sourceCatalogSha256}`);
    console.log(`LEDGER_SHA256 ${result.summary.ledgerSha256}`);
    console.log(`EXPECTED_REJECTION_CASES ${result.summary.expectedRejectionCaseCount}`);
    console.log(`EXPECTED_REJECTION_EVIDENCE ${result.summary.expectedRejectionEvidenceCount}`);
    console.log("READY_FOR_PATCH_15_3_2_C true");
    console.log("COLLECTOR_ENABLED_BY_THIS_OPERATION false");
    console.log("INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
    console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
