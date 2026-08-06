"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const root = path.resolve(__dirname, "..");
const files = [
  "automation/queryCandidatePlannerRealShadowPreparation.js",
  "scripts/queryCandidatePlannerGenerateRealShadowEvidenceSecret.js",
  "scripts/queryCandidatePlannerScaffoldRealShadowCaseRegistry.js",
  "scripts/queryCandidatePlannerPrepareRealShadowCaseRegistry.js",
  "scripts/queryCandidatePlannerVerifyRealShadowPreparation.js",
];
const source = files
  .map((file) => fs.readFileSync(path.join(root, file), "utf8"))
  .join("\n");
for (const token of [
  "syntheticFingerprintAllowed: false",
  "collectorEnabledByThisOperation: false",
  "internalCanaryEnabledByThisOperation: false",
  "productionPromotionAuthorized: false",
  "rawSecretIncluded: false",
  "rawIdentityIncluded: false",
  "REAL_SHADOW_CASE_REGISTRY_READY",
]) {
  assert(source.includes(token), `required source token missing: ${token}`);
}
for (const token of [
  "collectorEnabledByThisOperation: true",
  "internalCanaryEnabledByThisOperation: true",
  "productionPromotionAuthorized: true",
  "process.env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED =",
]) {
  assert(!source.includes(token), `forbidden source token found: ${token}`);
}
console.log("PASS query candidate patch15.3.2-A source integrity smoke");
