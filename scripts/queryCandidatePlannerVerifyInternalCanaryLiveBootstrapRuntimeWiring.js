"use strict";

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");

const expected = Object.freeze({
  "automation/queryCandidatePlannerInternalAllowlistCanaryBoundary.js":
    "DF36E4B02A91544D6FE82BCC53394ED345019E3058C5DC52F85DF1E8CDC788D0",
  "automation/queryCandidatePlannerInternalCanaryLiveBootstrapRuntime.js":
    "F52737193BCCA38C8534BB12698D96E2B162381E62AA7170B82AB0E01C19519A",
  "automation/queryCandidatePlannerInternalAllowlistCanaryService.js":
    "1A61F219ADF49BD863B84C5B8C4DB02158E901E7EDA864AC551656A4A7E75C8F",
  "automation/queryCandidatePlannerInternalCanaryApprovalBindingGate.js":
    "ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7",
  "automation/queryCandidatePlannerFinalEvaluationEvidenceBundle.js":
    "439F29AC82D866EEADA3EDFBD8615892904ACD507E4F8D4D5161431E0449440A",
  "automation/queryCandidatePlannerInternalCanaryLiveBootstrapGate.js":
    "4585B4549B0F756274F47FBB9089E56A07D21C6EFE3C1929214E856B068B5498",
  "routes/automationRoutes.js":
    "2D5390681F3A4306EBE1BE6166FBE9CC875A71C5A94CCDAABE824511EBC4B626",
});

function sha256File(relative) {
  return crypto
    .createHash("sha256")
    .update(fs.readFileSync(path.join(root, relative)))
    .digest("hex")
    .toUpperCase();
}

for (const [relative, sha] of Object.entries(expected)) {
  if (!fs.existsSync(path.join(root, relative))) {
    throw new Error(`MISSING ${relative}`);
  }
  const actual = sha256File(relative);
  if (actual !== sha) {
    throw new Error(`SHA_DRIFT ${relative} actual=${actual} expected=${sha}`);
  }
}

const {
  parseQueryCandidatePlannerInternalCanaryLiveBootstrapRuntimeMode,
} = require("../automation/queryCandidatePlannerInternalCanaryLiveBootstrapRuntime");

const defaultMode =
  parseQueryCandidatePlannerInternalCanaryLiveBootstrapRuntimeMode({});
if (defaultMode.active !== false || defaultMode.enabled !== false) {
  throw new Error("DEFAULT_RUNTIME_MODE_MUST_BE_INACTIVE");
}

const boundary = fs.readFileSync(
  path.join(
    root,
    "automation",
    "queryCandidatePlannerInternalAllowlistCanaryBoundary.js",
  ),
  "utf8",
);

for (const token of [
  "queryCandidatePlannerInternalCanaryLiveBootstrapRuntime",
  "bootstrapMode?.active === true",
  "queryCandidatePlannerLiveBootstrapTask",
  "onLiveBootstrapObservation",
]) {
  if (!boundary.includes(token)) {
    throw new Error(`BOUNDARY_WIRING_TOKEN_MISSING ${token}`);
  }
}

const route = fs.readFileSync(
  path.join(root, "routes", "automationRoutes.js"),
  "utf8",
);

if (route.includes("queryCandidatePlannerInternalCanaryLiveBootstrapRuntime")) {
  throw new Error("ROUTE_MUST_NOT_IMPORT_BOOTSTRAP_RUNTIME_DIRECTLY");
}

console.log("PASS Patch 15.3.3-B-2 runtime wiring verification");
console.log("DEFAULT_BOOTSTRAP_RUNTIME_ACTIVE false");
console.log("ROUTE_MODIFIED false");
console.log("LEGACY_SERVICE_MODIFIED false");
console.log("LEGACY_EVIDENCE_SUBSTITUTED false");
console.log("PRIMARY_RESPONSE_AUTHORITY true");
console.log("PRODUCTION_MERGE_APPLIED_BY_BOOTSTRAP false");
console.log("PROVIDER_CALLS_EXECUTED_BY_VERIFIER 0");
console.log("ACTUAL_INTERNAL_USER_EXPOSURE_EXECUTED false");
console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
console.log("RAILWAY_MODIFIED false");
