const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { execFileSync } = require("child_process");

const {
  invokeExistingCostCacheLatencyEvaluator,
} = require("../automation/queryCandidatePlannerCostCacheLatencyEvaluatorAdapter");

const {
  assessEvaluation,
} = require("../automation/queryCandidatePlannerActualPricingAbsoluteCostThresholdRecalibration");

const {
  buildCanonicalEvaluationBaseline,
} = require("../automation/queryCandidatePlannerCanonicalEvaluationBaseline");

function args(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) out[argv[i].slice(2)] = argv[++i] || "";
  }
  return out;
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(path.resolve(file), "utf8"));
}

function writeJson(file, value) {
  const target = path.resolve(file);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function shaFile(file) {
  return crypto
    .createHash("sha256")
    .update(fs.readFileSync(path.resolve(file)))
    .digest("hex")
    .toUpperCase();
}

function gitHeadSha(file) {
  const rel = path
    .relative(process.cwd(), path.resolve(file))
    .replace(/\\/g, "/");
  const data = execFileSync("git", ["show", `HEAD:${rel}`], { encoding: null });
  return crypto.createHash("sha256").update(data).digest("hex").toUpperCase();
}

async function main() {
  const input = args(process.argv.slice(2));
  for (const key of [
    "input",
    "pricing",
    "readiness",
    "threshold-policy",
    "report-output",
    "assessment-output",
    "baseline-output",
  ]) {
    if (!input[key]) throw new Error(`--${key} is required`);
  }

  const evaluatorPath =
    "automation/queryCandidatePlannerCostCacheLatencyEvaluator.js";

  const worktreeSha256 = shaFile(evaluatorPath);
  const headSha256 = gitHeadSha(evaluatorPath);
  const worktreeEqualsHead = worktreeSha256 === headSha256;

  if (!worktreeEqualsHead) {
    const error = new Error("Existing evaluator differs from Git HEAD.");
    error.code = "EVALUATOR_WORKTREE_DRIFT";
    throw error;
  }

  const canonicalInput = readJson(input.input);
  const pricingPolicy = readJson(input.pricing);
  const liveParityReadiness = readJson(input.readiness);
  const thresholdPolicy = readJson(input["threshold-policy"]);

  const invocation = await invokeExistingCostCacheLatencyEvaluator({
    dataset: canonicalInput.dataset,
    pricingPolicy,
    thresholdPolicy,
  });

  const report = invocation.report;
  const assessment = assessEvaluation(report, thresholdPolicy);

  writeJson(input["report-output"], report);
  writeJson(input["assessment-output"], assessment);

  const baseline = buildCanonicalEvaluationBaseline({
    canonicalInput,
    pricingPolicy,
    liveParityReadiness,
    thresholdPolicy,
    operationalReport: report,
    evaluatorIdentity: {
      worktreeSha256,
      headSha256,
      worktreeEqualsHead,
    },
  });

  writeJson(input["baseline-output"], baseline);

  console.log(
    "PASS patch 15.3.2-F.1.3 actual-pricing absolute Cost threshold re-evaluation",
  );
  console.log(`EVALUATOR_EXPORT ${invocation.exportName}`);
  console.log(`INVOCATION_SHAPE ${invocation.invocationShape}`);
  console.log(`EVALUATOR_SHA256 ${worktreeSha256}`);
  console.log(`WORKTREE_EQUALS_HEAD ${worktreeEqualsHead}`);
  console.log(`OPERATIONAL_DECISION ${report.decision}`);
  console.log(`ASSESSMENT_DECISION ${assessment.decision}`);
  console.log(`FAILED_CHECK_COUNT ${assessment.failedCheckCount}`);
  console.log(`ABSOLUTE_COST_PASS_COUNT ${assessment.absoluteCostPassCount}`);
  console.log(
    `ABSOLUTE_COST_FAILURE_COUNT ${assessment.absoluteCostFailureCount}`,
  );
  console.log(
    `CACHE_COST_AVOIDANCE_PASSED ${assessment.cacheCostAvoidancePassed}`,
  );
  console.log(
    `CACHE_COST_AVOIDANCE_ACTUAL ${assessment.cacheCostAvoidanceActual}`,
  );
  console.log(
    `CACHE_COST_AVOIDANCE_THRESHOLD ${assessment.cacheCostAvoidanceThreshold}`,
  );
  console.log(
    `AVERAGE_COST_MAX_MICROUSD ${assessment.thresholds.averageCostMicrousdMax}`,
  );
  console.log(
    `PROVIDER_CALL_AVERAGE_COST_MAX_MICROUSD ${assessment.thresholds.providerCallAverageCostMicrousdMax}`,
  );
  console.log(
    `MONTHLY_PROJECTED_COST_MAX_MICROUSD ${assessment.thresholds.monthlyProjectedCostMicrousdMax}`,
  );
  console.log("SOURCE_THRESHOLD_POLICY_MODIFIED false");
  console.log("EVALUATOR_MODIFIED false");
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
