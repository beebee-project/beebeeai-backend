"use strict";
const assert = require("assert");
const {
  buildUploadableSourceCatalogScaffold,
  bindUploadableSource,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");
const { accuracyDataset, createSourceFiles } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const dataset = accuracyDataset();
const sources = createSourceFiles(dataset);
let catalog = buildUploadableSourceCatalogScaffold(dataset);
const sameFile = sources.byCaseId.get(dataset.cases[0].caseId);
catalog = bindUploadableSource({
  accuracyDataset: dataset,
  catalog,
  caseId: dataset.cases[0].caseId,
  sourcePath: sameFile,
  sourceKind: "PUBLIC_DATASET",
  semanticCompatibilityConfirmed: true,
}).catalog;
assert.throws(
  () => bindUploadableSource({
    accuracyDataset: dataset,
    catalog,
    caseId: dataset.cases[1].caseId,
    sourcePath: sameFile,
    sourceKind: "PUBLIC_DATASET",
    semanticCompatibilityConfirmed: true,
  }),
  (error) => [
    "REAL_SHADOW_SOURCE_ARTIFACT_DUPLICATE",
    "REAL_SHADOW_SOURCE_PATH_DUPLICATE",
  ].includes(error.code),
);
console.log("PASS query candidate patch15.3.2-B.1 duplicate source fail-closed smoke");
