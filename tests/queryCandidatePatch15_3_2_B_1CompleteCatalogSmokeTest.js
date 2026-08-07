"use strict";
const assert = require("assert");
const { completeSourceCatalog } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const result = completeSourceCatalog();
assert.strictEqual(result.finalized.valid, true);
assert.strictEqual(result.finalized.completedCount, 10);
assert(/^[a-f0-9]{64}$/.test(result.finalized.sourceCatalogSha256));
assert(result.finalized.publicCatalog.cases.every((item) => !("sourcePath" in item)));
console.log("PASS query candidate patch15.3.2-B.1 complete source catalog smoke");
