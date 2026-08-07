"use strict";
const assert = require("assert");
const {
  validateRealShadowFingerprintLedger,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");
const { completeSourceCatalog } = require("./queryCandidatePatch15_3_2_B_1TestSupport");
const support = completeSourceCatalog();
const result = validateRealShadowFingerprintLedger({
  accuracyDataset: support.dataset,
  sourceCatalog: support.catalog,
  ledger: {
    version: "query_candidate_planner_real_shadow_fingerprint_ledger_v1",
    cases: [],
  },
  requireComplete: false,
});
assert.strictEqual(result.valid, false);
assert(result.errors.includes("REAL_SHADOW_LEGACY_LEDGER_REJECTED"));
console.log("PASS query candidate patch15.3.2-B.1 legacy ledger fail-closed smoke");
