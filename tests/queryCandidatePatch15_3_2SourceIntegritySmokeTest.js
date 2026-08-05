"use strict";
const assert=require("assert"),fs=require("fs"),path=require("path");
const ROOT=path.resolve(__dirname,"..");
const files=[
  "automation/queryCandidatePlannerRealShadowEvidenceConfig.js",
  "automation/queryCandidatePlannerRealShadowEvidenceCrypto.js",
  "automation/queryCandidatePlannerRealShadowCaptureBridge.js",
  "automation/queryCandidatePlannerRealShadowEvidenceStore.js",
  "automation/queryCandidatePlannerRealShadowEvidenceCollector.js",
  "automation/queryCandidatePlannerRealShadowEvidenceBundleBuilder.js",
  "models/QueryCandidatePlannerRealShadowEvidenceObservation.js",
  "routes/automationRoutes.js",
  "routes/fileRoutes.js",
  "scripts/queryCandidatePlannerExportRealShadowEvidence.js",
  "scripts/queryCandidatePlannerBuildRealShadowEvidenceBundle.js",
];
const source=files.map(f=>fs.readFileSync(path.join(ROOT,f),"utf8")).join("\n");
for(const token of [
  "REAL_SHADOW_TRAFFIC",
  "actualTraffic: true",
  "synthetic: false",
  "SEMANTIC_PROFILER_ONLY",
  "plannerEscalationAllowed: false",
  "APPROVED_ACTUAL",
  "aes-256-gcm",
  "primaryResponseUnchanged",
  "promotionAuthorized: false",
  "productionCandidateMerge: false",
  "productionReadyAssignment: false",
  "productionRouteChanged: false",
]) assert(source.includes(token),`required source token missing: ${token}`);
for(const token of [
  "plannerEscalationAllowed: true",
  "promotionAuthorized: true",
  "productionCandidateMerge: true",
  "productionReadyAssignment: true",
  "productionRouteChanged: true",
  "rawRowsIncluded: true",
  "userIdentityIncluded: true",
]) assert(!source.includes(token),`forbidden source token found: ${token}`);
console.log("PASS query candidate patch15.3.2 source integrity smoke");
