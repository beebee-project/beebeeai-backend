"use strict";

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const {
  buildFinalEvaluationEvidenceBundle,
} = require("../automation/queryCandidatePlannerFinalEvaluationEvidenceBundle");

function arg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? String(process.argv[index + 1] || "").trim() : "";
}

function main() {
  const inputs = {
    candidateFile: arg("--candidate-bundle"),
    rotationFile: arg("--rotation-plan"),
    receiptFile: arg("--approval-receipt"),
    approvalBindingGateFile: arg("--approval-gate"),
    composedServiceFile: arg("--composed-service"),
  };
  const output = arg("--output");
  if (!output) throw new Error("--output is required");

  const bundle = buildFinalEvaluationEvidenceBundle(inputs);
  const target = path.resolve(output);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const bytes = Buffer.from(`${JSON.stringify(bundle, null, 2)}\n`, "utf8");
  fs.writeFileSync(target, bytes, { mode: 0o600 });
  const fileSha256 = crypto.createHash("sha256").update(bytes).digest("hex").toUpperCase();

  console.log("PASS Patch 15.3.2-G final evaluation evidence bundle built");
  console.log(`READINESS_DECISION ${bundle.decision}`);
  console.log(`BUNDLE_PAYLOAD_SHA256 ${bundle.bundlePayloadSha256}`);
  console.log(`BUNDLE_FILE_SHA256 ${fileSha256}`);
  console.log(`FINAL_BASELINE_SHA256 ${bundle.immutableBindings.finalBaselineSha256}`);
  console.log(`CANDIDATE_PAYLOAD_SHA256 ${bundle.immutableBindings.candidatePayloadSha256}`);
  console.log(`ALLOWLIST_SHA256 ${bundle.immutableBindings.allowlistSha256}`);
  console.log(`APPROVAL_RECEIPT_PAYLOAD_SHA256 ${bundle.immutableBindings.approvalReceiptPayloadSha256}`);
  console.log("BOOTSTRAP_ONLY true");
  console.log("LEGACY_15_3_REAL_SHADOW_CONTRACT_SATISFIED false");
  console.log("ACTUAL_TRAFFIC_EVIDENCE_REQUIRED_FOR_15_3_4 true");
  console.log("RUNTIME_AUTO_ACTIVATION_AUTHORIZED false");
  console.log("ACTUAL_INTERNAL_USER_EXPOSURE_AUTHORIZED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_BUNDLE_BUILDER 0");
  console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
  console.log("RAILWAY_MODIFIED false");
  console.log("ENVIRONMENT_MODIFIED false");
  console.log("ROUTE_MODIFIED false");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
}

try {
  main();
} catch (error) {
  console.error(`BLOCKED ${error.code || error.message}`);
  process.exitCode = 1;
}
