const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const {
  buildManualApprovalReceipt,
} = require("../automation/queryCandidatePlannerInternalCanaryManualApprovalReceipt");

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

  if (!args["candidate-bundle"])
    throw new Error("--candidate-bundle is required");
  if (!args["allowlist-sha256"])
    throw new Error("--allowlist-sha256 is required");
  if (!args.output) throw new Error("--output is required");

  const receipt = buildManualApprovalReceipt({
    candidateBundleFile: args["candidate-bundle"],
    allowlistSha256: args["allowlist-sha256"],
    approve: args.approve,
  });

  const target = path.resolve(args.output);
  fs.mkdirSync(path.dirname(target), { recursive: true });

  const bytes = Buffer.from(`${JSON.stringify(receipt, null, 2)}\n`, "utf8");
  fs.writeFileSync(target, bytes);

  const fileSha256 = crypto
    .createHash("sha256")
    .update(bytes)
    .digest("hex")
    .toUpperCase();

  console.log(
    "PASS patch 15.3.2-F.1.5 internal canary manual approval receipt built",
  );
  console.log(`APPROVAL_DECISION ${receipt.decision}`);
  console.log(
    `CANDIDATE_PAYLOAD_SHA256 ${receipt.immutableBindings.candidatePayloadSha256}`,
  );
  console.log(
    `CANDIDATE_BUNDLE_FILE_SHA256 ${receipt.immutableBindings.candidateBundleFileSha256}`,
  );
  console.log(`ALLOWLIST_SHA256 ${receipt.immutableBindings.allowlistSha256}`);
  console.log(
    `ALLOWLIST_ENV_NAME ${receipt.immutableBindings.allowlistEnvironmentVariableName}`,
  );
  console.log(
    `APPROVAL_RECEIPT_PAYLOAD_SHA256 ${receipt.approvalReceiptPayloadSha256}`,
  );
  console.log(`APPROVAL_RECEIPT_FILE_SHA256 ${fileSha256}`);
  console.log("INTERNAL_CANARY_APPROVAL_GRANTED true");
  console.log("RUNTIME_GATE_BINDING_APPLIED false");
  console.log("RUNTIME_CANARY_AUTHORIZED false");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_RECEIPT_BUILDER 0");
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.code || error.message}`);
    process.exitCode = 1;
  }
}
