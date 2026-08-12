const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const {
  buildCandidate,
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

  const map = {
    pricingFile: "pricing",
    readinessFile: "readiness",
    consistentInputFile: "consistent-input",
    sourceThresholdFile: "source-threshold",
    recalibratedThresholdFile: "recalibrated-threshold",
    recalibrationEvidenceFile: "recalibration-evidence",
    operationalReportFile: "operational-report",
    assessmentFile: "assessment",
    baselineFile: "baseline",
  };

  const options = {};
  for (const [target, arg] of Object.entries(map)) {
    if (!input[arg]) throw new Error(`--${arg} is required`);
    options[target] = input[arg];
  }

  if (!input.output) throw new Error("--output is required");

  const bundle = buildCandidate(options);
  verifyCandidate(bundle);

  const target = path.resolve(input.output);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const bytes = Buffer.from(`${JSON.stringify(bundle, null, 2)}\n`, "utf8");
  fs.writeFileSync(target, bytes);

  const fileSha256 = crypto
    .createHash("sha256")
    .update(bytes)
    .digest("hex")
    .toUpperCase();

  console.log(
    "PASS patch 15.3.2-F.1.4 internal canary evidence candidate built",
  );
  console.log(`REVIEW_DECISION ${bundle.eligibility.decision}`);
  console.log(
    `INTERNAL_CANARY_REVIEW_ELIGIBLE ${bundle.eligibility.internalCanaryReviewEligible}`,
  );
  console.log(
    `INTERNAL_CANARY_AUTHORIZED ${bundle.eligibility.internalCanaryAuthorized}`,
  );
  console.log(
    `PERCENTAGE_ROLLOUT_AUTHORIZED ${bundle.eligibility.percentageRolloutAuthorized}`,
  );
  console.log(
    `PRODUCTION_PROMOTION_AUTHORIZED ${bundle.eligibility.productionPromotionAuthorized}`,
  );
  console.log(
    `ACTUAL_OPERATIONAL_TELEMETRY ${bundle.methodology.actualOperationalTelemetry}`,
  );
  console.log(`EVALUATOR_SHA256 ${bundle.integrity.evaluator.worktreeSha256}`);
  console.log(
    `WORKTREE_EQUALS_HEAD ${bundle.integrity.evaluator.worktreeEqualsHead}`,
  );
  console.log(`FINAL_BASELINE_SHA256 ${bundle.integrity.finalBaselineSha256}`);
  console.log(`CANDIDATE_PAYLOAD_SHA256 ${bundle.candidatePayloadSha256}`);
  console.log(`CANDIDATE_FILE_SHA256 ${fileSha256}`);
  console.log(
    `PROVIDER_CALLS_EXECUTED_BY_BUNDLE_BUILDER ${bundle.guardrails.providerCallsExecutedByBundleBuilder}`,
  );
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
