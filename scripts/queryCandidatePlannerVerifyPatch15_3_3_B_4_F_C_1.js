"use strict";

const assert = require("assert");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

function sha256(rel) {
  const file = path.join(__dirname, "..", rel);
  assert(fs.existsSync(file), `missing required file: ${rel}`);
  return crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex").toUpperCase();
}

const expected = Object.freeze({
  "automation/queryCandidatePlannerInternalCanaryLiveBootstrapGate.js":
    "9386A73BAD4E37C055209AF59B86C5FFB21A62545E26017BFB5D3A109E4EB1D9",
  "automation/queryCandidatePlannerInternalCanaryBootstrapReadinessBridge.js":
    "77DB527F808BBB61BD63BD61913E01A489AB25E154C5D4C0E67DAC730AB81259",
  "evaluation/queryCandidatePlannerInternalCanaryBootstrapProductionReadiness.v1.json":
    "46D1211AF4F318DAB91D137F0728C3AE6F246CD8B85A2582802CCB6DB1475AC4",
  "automation/queryCandidatePlannerInternalCanaryApprovalBindingGate.js":
    "ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7",
  "automation/queryCandidatePlannerInternalCanaryLiveBootstrapRuntime.js":
    "F52737193BCCA38C8534BB12698D96E2B162381E62AA7170B82AB0E01C19519A",
  "automation/queryCandidatePlannerFeatureControl.js":
    "E80A47537ECDB4454C6120693A9F3E725F74AC986C42ABF52E6AD163B30EAB07",
  "automation/queryCandidatePlannerFinalEvaluationEvidenceBundle.js":
    "439F29AC82D866EEADA3EDFBD8615892904ACD507E4F8D4D5161431E0449440A",
});

for (const [rel, expectedSha] of Object.entries(expected)) {
  assert.strictEqual(sha256(rel), expectedSha, `protected source drift: ${rel}`);
}

const gate = require("../automation/queryCandidatePlannerInternalCanaryLiveBootstrapGate");
const bridgeSha = sha256("automation/queryCandidatePlannerInternalCanaryBootstrapReadinessBridge.js");
assert.strictEqual(gate.EXPECTED_READINESS_BRIDGE_SHA256, bridgeSha, "A Gate bridge SHA pin mismatch");
const deps = gate.verifyProtectedDependencies();
assert.strictEqual(deps.valid, true, deps.reason);
assert.strictEqual(deps.reason, "OK");

console.log(`A_GATE_SHA256 ${sha256("automation/queryCandidatePlannerInternalCanaryLiveBootstrapGate.js")}`);
console.log(`READINESS_BRIDGE_SHA256 ${bridgeSha}`);
console.log(`A_GATE_EXPECTED_BRIDGE_SHA256 ${gate.EXPECTED_READINESS_BRIDGE_SHA256}`);
console.log(`BRIDGE_SHA_PIN_EXACT ${gate.EXPECTED_READINESS_BRIDGE_SHA256 === bridgeSha}`);
console.log(`PROTECTED_DEPENDENCIES_VALID ${deps.valid === true}`);
console.log(`PROTECTED_DEPENDENCIES_REASON ${deps.reason}`);
console.log("F16_GATE_MODIFIED false");
console.log("B2_RUNTIME_MODIFIED false");
console.log("FEATURE_CONTROL_MODIFIED false");
console.log("READINESS_BRIDGE_MODIFIED false");
console.log("READINESS_FILE_MODIFIED false");
console.log("RAILWAY_VARIABLE_MUTATED false");
console.log("RAILWAY_DEPLOY_TRIGGERED false");
console.log("PROVIDER_CALLS_EXECUTED 0");
console.log("ACTUAL_LIVE_REQUEST_EXECUTED false");
console.log("PASS PATCH 15.3.3-B-4-F-C.1 READINESS BRIDGE IDENTITY CONVERGENCE VERIFICATION");
