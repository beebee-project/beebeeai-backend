"use strict";
const assert = require("assert");
const {
  buildUploadableSourceCatalogScaffold,
  finalizeUploadableSourceCatalog,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");
const { accuracyDataset } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const dataset = accuracyDataset();
const result = finalizeUploadableSourceCatalog({
  accuracyDataset: dataset,
  catalog: buildUploadableSourceCatalogScaffold(dataset),
});
assert.strictEqual(result.valid, false);
assert(result.errors.some((item) => item.includes("sourceKind")));
console.log("PASS query candidate patch15.3.2-B.1 incomplete catalog fail-closed smoke");
