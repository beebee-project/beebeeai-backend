#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  buildUploadableSourceCatalogScaffold,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}

try {
  const root = path.resolve(__dirname, "..");
  const dataset = JSON.parse(
    fs.readFileSync(
      path.join(root, "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json"),
      "utf8",
    ),
  );
  const output = path.resolve(
    arg(
      "--output",
      "queryCandidatePlannerRealShadowUploadableSourceCatalog.draft.private.json",
    ),
  );
  if (!/\.private\./i.test(path.basename(output))) {
    const error = new Error("source catalog output must use a .private. filename");
    error.code = "REAL_SHADOW_PRIVATE_OUTPUT_NAME_REQUIRED";
    throw error;
  }
  const catalog = buildUploadableSourceCatalogScaffold(dataset, {
    catalogId: arg(
      "--catalog-id",
      "internal_real_shadow_uploadable_sources_2026_08_v1",
    ),
  });
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, `${JSON.stringify(catalog, null, 2)}\n`, {
    encoding: "utf8",
    mode: 0o600,
  });
  console.log(`PASS uploadable source catalog scaffold cases=${catalog.cases.length}`);
  console.log(`OUTPUT ${output}`);
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
  console.log("ACTUAL_SOURCE_FILES_INCLUDED false");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
