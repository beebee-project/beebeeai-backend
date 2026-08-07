"use strict";
const assert = require("assert");
const {
  buildUploadableSourceCatalogScaffold,
  bindUploadableSource,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");
const { accuracyDataset, createSourceFiles } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const dataset = accuracyDataset();
const sources = createSourceFiles(dataset);
const catalog = buildUploadableSourceCatalogScaffold(dataset);
assert.throws(
  () => bindUploadableSource({
    accuracyDataset: dataset,
    catalog,
    caseId: dataset.cases[0].caseId,
    sourcePath: sources.byCaseId.get(dataset.cases[0].caseId),
    sourceKind: "SYNTHETIC",
    semanticCompatibilityConfirmed: true,
  }),
  (error) => error.code === "REAL_SHADOW_SOURCE_KIND_INVALID",
);
console.log("PASS query candidate patch15.3.2-B.1 synthetic source fail-closed smoke");
