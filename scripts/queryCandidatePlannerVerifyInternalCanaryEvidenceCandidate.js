const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const {
  verifyCandidate,
} = require("../automation/queryCandidatePlannerInternalCanaryEvidenceCandidate");

function args(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) out[argv[i].slice(2)] = argv[++i] || "";
  }
  return out;
}

function main() {
  const input = args(process.argv.slice(2));
  if (!input.bundle) throw new Error("--bundle is required");

  const target = path.resolve(input.bundle);
  const bytes = fs.readFileSync(target);
  const bundle = JSON.parse(bytes.toString("utf8"));

  verifyCandidate(bundle);

  const fileSha256 = crypto
    .createHash("sha256")
    .update(bytes)
    .digest("hex")
    .toUpperCase();

  console.log("PASS internal canary evidence candidate verification");
  console.log(`REVIEW_DECISION ${bundle.eligibility.decision}`);
  console.log(
    `INTERNAL_CANARY_REVIEW_ELIGIBLE ${bundle.eligibility.internalCanaryReviewEligible}`,
  );
  console.log("MANUAL_OPERATOR_APPROVAL_REQUIRED true");
  console.log("INTERNAL_CANARY_AUTHORIZED false");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
  console.log(`CANDIDATE_PAYLOAD_SHA256 ${bundle.candidatePayloadSha256}`);
  console.log(`CANDIDATE_FILE_SHA256 ${fileSha256}`);
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.code || error.message}`);
    process.exitCode = 1;
  }
}
