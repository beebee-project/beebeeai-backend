"use strict";

const fs = require("fs");
const path = require("path");

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) out[argv[i].slice(2)] = argv[++i] || "";
  }
  return out;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.receipt) throw new Error("--receipt is required");

  const target = path.resolve(args.receipt);
  if (!fs.existsSync(target)) throw new Error("Receipt file missing");

  const receipt = JSON.parse(fs.readFileSync(target, "utf8"));
  const payloadSha = String(
    receipt.approvalReceiptPayloadSha256 || "",
  ).trim();

  if (!/^[A-Fa-f0-9]{64}$/.test(payloadSha)) {
    throw new Error("Receipt payload SHA invalid");
  }

  const compact = JSON.stringify(receipt);

  console.log(
    `QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_BUNDLE_SHA256=${payloadSha}`,
  );
  console.log(
    `QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_RECEIPT_JSON=${compact}`,
  );
  console.log("RAW_IMMUTABLE_ACCOUNT_ID_INCLUDED false");
  console.log("RAW_TENANT_ID_INCLUDED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
