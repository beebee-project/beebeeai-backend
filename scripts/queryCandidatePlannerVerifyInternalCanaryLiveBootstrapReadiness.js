"use strict";

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const {
  evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapGate,
  ENV,
  EXPECTED_G_BUNDLE_PAYLOAD_SHA256,
} = require("../automation/queryCandidatePlannerInternalCanaryLiveBootstrapGate");

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) out[argv[i].slice(2)] = argv[++i] || "";
  }
  return out;
}

function required(args, name) {
  const value = String(args[name] || "").trim();
  if (!value) throw new Error(`--${name} is required`);
  return value;
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(path.resolve(file), "utf8"));
}

function sha256File(file) {
  return crypto
    .createHash("sha256")
    .update(fs.readFileSync(path.resolve(file)))
    .digest("hex")
    .toUpperCase();
}

function normalizeSha256(value) {
  const normalized = String(value || "").trim().toUpperCase();
  if (!/^[A-F0-9]{64}$/.test(normalized)) throw new Error("SHA256_INVALID");
  return normalized;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const gBundleFile = required(args, "g-bundle");
  const receiptFile = required(args, "approval-receipt");
  const subjectSha256 = normalizeSha256(required(args, "subject-sha256"));

  const gBundle = readJson(gBundleFile);
  const receipt = readJson(receiptFile);
  const approvalReceiptPayloadSha256 = normalizeSha256(
    receipt.approvalReceiptPayloadSha256,
  );
  const allowlistSha256 = normalizeSha256(
    receipt.immutableBindings?.allowlistSha256,
  );

  // Entire environment is local to this verification object. process.env is
  // not mutated and no Provider/route/merge runner is invoked.
  const env = {
    [ENV.enabled]: "true",
    [ENV.killSwitch]: "false",
    [ENV.bundleJson]: JSON.stringify(gBundle),
    [ENV.bundleSha256]: EXPECTED_G_BUNDLE_PAYLOAD_SHA256,

    QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_RECEIPT_JSON:
      JSON.stringify(receipt),
    QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_BUNDLE_SHA256:
      approvalReceiptPayloadSha256,
    QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256: allowlistSha256,

    QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED: "true",
    QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH: "false",
    QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_LLM_MODE:
      "SEMANTIC_PROFILER_ONLY",

    QUERY_CANDIDATE_PLANNER_KILL_SWITCH: "false",
    QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED: "true",
    QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED: "true",
    QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED: "true",
    QUERY_CANDIDATE_PLANNER_PROVIDER_KILL_SWITCH: "false",

    QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED: "true",
    QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED: "true",
    QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED: "false",
    QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED: "false",
    QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH: "false",

    QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED: "true",
    QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE: "ALLOWLIST",
    QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT: "0",
  };

  const featureControl = {
    evaluate(operation) {
      return { allowed: true, reason: `LOCAL_READINESS_${operation}` };
    },
  };

  const legacyPreflight = {
    allowed: false,
    status: "BLOCKED",
    reason: "READINESS_EVIDENCE_INVALID",
    evidence: {
      valid: false,
      reason: "READINESS_EVIDENCE_INVALID",
    },
  };

  const result =
    evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapGate({
      env,
      featureControl,
      subject: { complete: true, subjectSha256 },
      legacyPreflight,
    });

  console.log(
    "PASS Patch 15.3.3-A local live-bootstrap readiness verification executed",
  );
  console.log(`BOOTSTRAP_DECISION ${result.decision}`);
  console.log(`BOOTSTRAP_ALLOWED ${result.allowed === true}`);
  console.log(`BOOTSTRAP_REASON ${result.reason}`);
  console.log(`G_BUNDLE_FILE_SHA256 ${sha256File(gBundleFile)}`);
  console.log(
    `G_BUNDLE_PAYLOAD_SHA256 ${result.bootstrapReadiness?.gBundlePayloadSha256 || ""}`,
  );
  console.log(
    `APPROVAL_RECEIPT_PAYLOAD_SHA256 ${approvalReceiptPayloadSha256}`,
  );
  console.log(`ALLOWLIST_SHA256 ${allowlistSha256}`);
  console.log(`SUBJECT_ALLOWLIST_MATCH ${subjectSha256 === allowlistSha256}`);
  console.log(
    `LEGACY_EVIDENCE_VALID ${result.legacyEvidence?.valid === true}`,
  );
  console.log(
    `LEGACY_EVIDENCE_SUBSTITUTED ${result.legacyEvidence?.substituted === true}`,
  );
  console.log(
    `RUNTIME_BOOTSTRAP_EXECUTION_ELIGIBLE ${result.runtimeBootstrapExecutionEligible === true}`,
  );
  console.log(
    `ACTUAL_INTERNAL_USER_EXPOSURE_EXECUTED ${result.actualInternalUserExposureExecuted === true}`,
  );
  console.log(
    `PROVIDER_CALLS_EXECUTED_BY_GATE ${Number(result.providerCallsExecutedByGate || 0)}`,
  );
  console.log(
    `ACTUAL_OPERATIONAL_TELEMETRY ${result.actualOperationalTelemetry === true}`,
  );
  console.log(
    `PERCENTAGE_ROLLOUT_AUTHORIZED ${result.percentageRolloutAuthorized === true}`,
  );
  console.log(
    `PRODUCTION_PROMOTION_AUTHORIZED ${result.productionPromotionAuthorized === true}`,
  );
  console.log("PROCESS_ENV_MUTATED false");
  console.log("RAILWAY_MODIFIED false");

  if (!result.allowed) process.exitCode = 1;
}

try {
  main();
} catch (error) {
  console.error(`BLOCKED ${error.code || error.message}`);
  process.exitCode = 1;
}
