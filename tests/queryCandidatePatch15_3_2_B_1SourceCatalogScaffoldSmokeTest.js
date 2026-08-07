"use strict";
const assert = require("assert");
const {
  buildUploadableSourceCatalogScaffold,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");
const { accuracyDataset } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const catalog = buildUploadableSourceCatalogScaffold(accuracyDataset());
assert.strictEqual(catalog.cases.length, 10);
assert(catalog.cases.every((item) => item.uploadable === false));
assert.strictEqual(catalog.sourcePolicy.syntheticForbidden, true);
console.log("PASS query candidate patch15.3.2-B.1 source catalog scaffold smoke");
