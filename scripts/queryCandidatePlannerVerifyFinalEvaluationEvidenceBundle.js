const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const {
  verifyFinalEvaluationEvidenceBundle,
} = require("../automation/queryCandidatePlannerFinalEvaluationEvidenceBundle");

function arg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? String(process.argv[index + 1] || "").trim() : "";
}

try {
  const file = arg("--bundle");
  if (!file) throw new Error("--bundle is required");
  const resolved = path.resolve(file);
  const bytes = fs.readFileSync(resolved);
  const bundle = JSON.parse(bytes.toString("utf8"));
  verifyFinalEvaluationEvidenceBundle(bundle);
  const fileSha256 = crypto
    .createHash("sha256")
    .update(bytes)
    .digest("hex")
    .toUpperCase();
  console.log(
    "PASS Patch 15.3.2-G final evaluation evidence bundle verification",
  );
  console.log(`READINESS_DECISION ${bundle.decision}`);
  console.log(`BUNDLE_PAYLOAD_SHA256 ${bundle.bundlePayloadSha256}`);
  console.log(`BUNDLE_FILE_SHA256 ${fileSha256}`);
  console.log("EVALUATION_EVIDENCE_FINALIZED true");
  console.log("INTERNAL_CANARY_BOOTSTRAP_READINESS true");
  console.log("BOOTSTRAP_ONLY true");
  console.log("LEGACY_EVIDENCE_SUBSTITUTION_FORBIDDEN true");
  console.log("LEGACY_15_3_REAL_SHADOW_CONTRACT_SATISFIED false");
  console.log("ACTUAL_TRAFFIC_EVIDENCE_REQUIRED_FOR_15_3_4 true");
  console.log("RUNTIME_AUTO_ACTIVATION_AUTHORIZED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_VERIFIER 0");
  console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
} catch (error) {
  console.error(`BLOCKED ${error.code || error.message}`);
  process.exitCode = 1;
}
