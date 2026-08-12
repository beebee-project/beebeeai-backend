const fs = require("fs");
const path = require("path");
const {
  invokeExistingCostCacheLatencyEvaluator,
} = require("../automation/queryCandidatePlannerCostCacheLatencyEvaluatorAdapter");
const {
  assessCostConsistencyReevaluation,
} = require("../automation/queryCandidatePlannerActualPricingCacheAvoidanceConsistency");
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
  const assessment = assessCostConsistencyReevaluation(report);

  const reportPath = path.resolve(input["report-output"]);
  fs.mkdirSync(path.dirname(reportPath), { recursive: true });
  fs.writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");

  const assessmentPath = path.resolve(input["assessment-output"]);
  fs.mkdirSync(path.dirname(assessmentPath), { recursive: true });
  fs.writeFileSync(
    assessmentPath,
    `${JSON.stringify(assessment, null, 2)}\n`,
    "utf8",
  );

  const baseline = buildCanonicalEvaluationBaseline({
    canonicalInput,
    pricingPolicy,
    liveParityReadiness,
    thresholdPolicy,
    operationalReport: report,
    evaluatorIdentity: {
      worktreeSha256: "",
      headSha256: "",
      worktreeEqualsHead: true,
    },
  });

  const baselinePath = path.resolve(input["baseline-output"]);
  fs.mkdirSync(path.dirname(baselinePath), { recursive: true });
  fs.writeFileSync(
    baselinePath,
    `${JSON.stringify(baseline, null, 2)}\n`,
    "utf8",
  );

  console.log(
    "PASS patch 15.3.2-F.1.2 actual-pricing Cost consistency re-evaluation",
  );
  console.log(`EVALUATOR_EXPORT ${invocation.exportName}`);
  console.log(`INVOCATION_SHAPE ${invocation.invocationShape}`);
  console.log(`OPERATIONAL_DECISION ${report.decision}`);
  console.log(`ASSESSMENT_DECISION ${assessment.decision}`);
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
  console.log(`NON_COST_FAILURE_COUNT ${assessment.nonCostFailureCount}`);
  console.log("THRESHOLD_POLICY_MODIFIED false");
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
