#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  validateUploadableSourceCatalog,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}

try {
  const root = path.resolve(__dirname, "..");
  const catalogArg = arg("--catalog");
  if (!catalogArg) throw new Error("--catalog is required");
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(root, "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json"),
      "utf8",
    ),
  );
  const catalog = JSON.parse(fs.readFileSync(path.resolve(catalogArg), "utf8"));
  const result = validateUploadableSourceCatalog({
    accuracyDataset,
    catalog,
    requireComplete: false,
    verifyFiles: true,
  });
  if (!result.valid) {
    result.errors.forEach((error) => console.error(`BLOCKED ${error}`));
    process.exitCode = 2;
  } else {
    catalog.cases.forEach((item) => {
      const ready =
        Boolean(item.sourcePath) &&
        /^[a-f0-9]{64}$/i.test(String(item.sourceArtifactSha256 || "")) &&
        item.uploadable === true &&
        item.semanticCompatibilityConfirmed === true;
      console.log(`${ready ? "READY" : "PENDING"} ${item.caseId}`);
    });
    console.log(`PROGRESS ${result.completedCount}/${result.expectedCaseCount}`);
    console.log(`REMAINING ${result.remainingCount}`);
    console.log(`COMPLETE ${result.complete}`);
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
