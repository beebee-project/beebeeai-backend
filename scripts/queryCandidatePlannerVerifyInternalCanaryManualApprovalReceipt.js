const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const {
  verifyManualApprovalReceipt,
  sha256File,
  normalizeSha256,
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

  if (!args.receipt) throw new Error("--receipt is required");
  if (!args["candidate-bundle"])
    throw new Error("--candidate-bundle is required");
  if (!args["allowlist-sha256"])
    throw new Error("--allowlist-sha256 is required");

  const receiptPath = path.resolve(args.receipt);
  const candidatePath = path.resolve(args["candidate-bundle"]);

  if (!fs.existsSync(receiptPath)) throw new Error("Receipt file missing");
  if (!fs.existsSync(candidatePath))
    throw new Error("Candidate bundle file missing");

  const receipt = JSON.parse(fs.readFileSync(receiptPath, "utf8"));
  verifyManualApprovalReceipt(receipt);

  const expectedAllowlist = normalizeSha256(
    args["allowlist-sha256"],
    "VERIFICATION_ALLOWLIST_SHA_INVALID",
  );

  if (receipt.immutableBindings.allowlistSha256 !== expectedAllowlist) {
    const error = new Error(
      "Approval receipt allowlist SHA does not match verification input.",
    );
    error.code = "ALLOWLIST_BINDING_MISMATCH";
    throw error;
  }

  if (
    receipt.immutableBindings.candidateBundleFileSha256 !==
    sha256File(candidatePath)
  ) {
    const error = new Error(
      "Approval receipt candidate bundle file SHA does not match current bundle.",
    );
    error.code = "CANDIDATE_FILE_BINDING_MISMATCH";
    throw error;
  }

  const receiptFileSha = crypto
    .createHash("sha256")
    .update(fs.readFileSync(receiptPath))
    .digest("hex")
    .toUpperCase();

  console.log("PASS internal canary manual approval receipt verification");
  console.log(`APPROVAL_DECISION ${receipt.decision}`);
  console.log(
    `CANDIDATE_PAYLOAD_SHA256 ${receipt.immutableBindings.candidatePayloadSha256}`,
  );
  console.log(`ALLOWLIST_SHA256 ${receipt.immutableBindings.allowlistSha256}`);
  console.log(
    `APPROVAL_RECEIPT_PAYLOAD_SHA256 ${receipt.approvalReceiptPayloadSha256}`,
  );
  console.log(`APPROVAL_RECEIPT_FILE_SHA256 ${receiptFileSha}`);
  console.log("INTERNAL_CANARY_APPROVAL_GRANTED true");
  console.log("RUNTIME_GATE_BINDING_APPLIED false");
  console.log("RUNTIME_CANARY_AUTHORIZED false");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.code || error.message}`);
    process.exitCode = 1;
  }
}
