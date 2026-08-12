const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const {
  invokeExistingCostCacheLatencyEvaluator,
} = require("../automation/queryCandidatePlannerCostCacheLatencyEvaluatorAdapter");
const {
  buildCanonicalEvaluationBaseline,
} = require("../automation/queryCandidatePlannerCanonicalEvaluationBaseline");

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) out[argv[i].slice(2)] = argv[++i] || "";
  }
  return out;
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(path.resolve(file), "utf8"));
}

function sha(data) {
  return crypto.createHash("sha256").update(data).digest("hex").toUpperCase();
}

function evaluatorIdentity() {
  const relative =
    "automation/queryCandidatePlannerCostCacheLatencyEvaluator.js";
  const worktree = fs.readFileSync(path.resolve(relative));
  let head = null;
  try {
    head = execFileSync("git", ["show", `HEAD:${relative}`], {
      encoding: null,
    });
  } catch {
    head = null;
  }
  return {
    worktreeSha256: sha(worktree),
    headSha256: head ? sha(head) : "",
    worktreeEqualsHead: head ? sha(worktree) === sha(head) : false,
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  for (const name of [
    "input",
    "pricing",
    "readiness",
    "threshold-policy",
    "report-output",
    "baseline-output",
  ]) {
    if (!args[name]) throw new Error(`--${name} is required`);
  }

  const canonicalInput = readJson(args.input);
  const pricingPolicy = readJson(args.pricing);
  const liveParityReadiness = readJson(args.readiness);
  const thresholdPolicy = readJson(args["threshold-policy"]);
  const identity = evaluatorIdentity();

  if (!identity.worktreeEqualsHead) {
    const error = new Error(
      "Cost/Cache/Latency evaluator WORKTREE differs from Git HEAD.",
    );
    error.code = "EVALUATOR_WORKTREE_DRIFT";
    throw error;
  }

  const invocation = await invokeExistingCostCacheLatencyEvaluator({
    dataset: canonicalInput.dataset,
    pricingPolicy,
    thresholdPolicy,
  });

  const reportPath = path.resolve(args["report-output"]);
  fs.mkdirSync(path.dirname(reportPath), { recursive: true });
  fs.writeFileSync(
    reportPath,
    `${JSON.stringify(invocation.report, null, 2)}\n`,
    "utf8",
  );

  const baseline = buildCanonicalEvaluationBaseline({
    canonicalInput,
    pricingPolicy,
    liveParityReadiness,
    thresholdPolicy,
    operationalReport: invocation.report,
    evaluatorIdentity: identity,
  });

  const baselinePath = path.resolve(args["baseline-output"]);
  fs.mkdirSync(path.dirname(baselinePath), { recursive: true });
  fs.writeFileSync(
    baselinePath,
    `${JSON.stringify(baseline, null, 2)}\n`,
    "utf8",
  );

  console.log("PASS patch 15.3.2-F.1 canonical Cost/Cache/Latency evaluation");
  console.log(`EVALUATOR_EXPORT ${invocation.exportName}`);
  console.log(`INVOCATION_SHAPE ${invocation.invocationShape}`);
  console.log(`EVALUATOR_SHA256 ${identity.worktreeSha256}`);
  console.log(`WORKTREE_EQUALS_HEAD ${identity.worktreeEqualsHead}`);
  console.log(`OPERATIONAL_DECISION ${invocation.report.decision}`);
  console.log(`BASELINE_DECISION ${baseline.decision}`);
  console.log(
    `COST_THRESHOLD_RECALIBRATION_REQUIRED ${baseline.operationalEvaluation.costThresholdRecalibrationRequired}`,
  );
  console.log(
    `NON_COST_FAILURE_COUNT ${baseline.operationalEvaluation.nonCostFailureCount}`,
  );
  console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_EVALUATOR 0");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log(`BASELINE_SHA256 ${baseline.baselineSha256}`);
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
}

if (require.main === module) {
  main().catch((error) => {
    console.error(`BLOCKED ${error.code || error.message}`);
    process.exitCode = 1;
  });
}
