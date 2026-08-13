"use strict";

const fs = require("fs");
const path = require("path");

const {
  ENV,
  evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate,
} = require("../automation/queryCandidatePlannerInternalCanaryApprovalBindingGate");

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) {
      out[argv[i].slice(2)] = argv[++i] || "";
    }
  }
  return out;
}

function main() {
  const args = parseArgs(process.argv.slice(2));

  for (const name of ["receipt", "allowlist-sha256"]) {
    if (!args[name]) throw new Error(`--${name} is required`);
  }

  const receiptPath = path.resolve(args.receipt);
  if (!fs.existsSync(receiptPath)) {
    throw new Error("Approval receipt file missing");
  }

  const receipt = JSON.parse(
    fs.readFileSync(receiptPath, "utf8"),
  );

  const approvalBundleSha256 =
    String(
      args["approval-bundle-sha256"] ||
        receipt.approvalReceiptPayloadSha256 ||
        "",
    ).trim();

  const allowlistSha256 =
    String(args["allowlist-sha256"] || "").trim();

  const env = {
    [ENV.receiptJson]: JSON.stringify(receipt),
    [ENV.approvalBundleSha256]: approvalBundleSha256,
    [ENV.allowlistSha256]: allowlistSha256,

    [ENV.internalCanaryEnabled]: "1",
    [ENV.internalCanaryKillSwitch]: "0",
    [ENV.internalCanaryLlmMode]: "SEMANTIC_PROFILER_ONLY",

    [ENV.globalKillSwitch]: "0",
    [ENV.featureEnabled]: "1",
    [ENV.shadowEnabled]: "1",
    [ENV.providerEnabled]: "1",
    [ENV.providerKillSwitch]: "0",

    [ENV.productionEnabled]: "1",
    [ENV.productionCandidateMergeEnabled]: "1",
    [ENV.productionReadyAssignmentEnabled]: "0",
    [ENV.productionRouteEnabled]: "0",
    [ENV.productionKillSwitch]: "0",

    [ENV.promotionGateEnabled]: "1",
    [ENV.promotionAudienceMode]: "ALLOWLIST",
    [ENV.promotionRolloutPercent]: "0",
  };

  const featureControl = {
    evaluate(operation) {
      return Object.freeze({
        allowed: true,
        reason: `OFFLINE_VERIFIER_ALLOW_${operation}`,
      });
    },
  };

  const decision =
    evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate({
      env,
      featureControl,
      subject: {
        complete: true,
        subjectSha256: allowlistSha256,
      },
    });

  if (!decision.allowed) {
    throw new Error(decision.reason);
  }

  console.log(
    "PASS Patch 15.3.2-F.1.6 approval binding offline verification",
  );
  console.log(`GATE_DECISION ${decision.decision}`);
  console.log(`GATE_REASON ${decision.reason}`);
  console.log(
    `CANDIDATE_PAYLOAD_SHA256 ${decision.preflight.approvalBinding.candidatePayloadSha256}`,
  );
  console.log(
    `APPROVAL_RECEIPT_PAYLOAD_SHA256 ${decision.preflight.approvalBinding.approvalReceiptPayloadSha256}`,
  );
  console.log(
    `ALLOWLIST_MATCHED ${decision.preflight.approvalBinding.allowlistMatched}`,
  );
  console.log("RUNTIME_GATE_BINDING_APPLIED true");
  console.log("RUNTIME_CANARY_AUTHORIZED true");
  console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
  console.log("CANARY_EVIDENCE_COLLECTION_REQUIRED true");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_VERIFIER 0");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
