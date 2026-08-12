const fs = require("fs");
const path = require("path");
const {
  buildRealShadowEvaluationBaseline,
} = require("../automation/queryCandidatePlannerRealShadowEvaluationBaseline");

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

function main() {
  const args = parseArgs(process.argv.slice(2));
  for (const key of [
    "collection-summary",
    "pricing",
    "operational-report",
    "output",
  ]) {
    if (!args[key]) throw new Error(`--${key} is required`);
  }

  const baseline = buildRealShadowEvaluationBaseline({
    collectionSummary: readJson(args["collection-summary"]),
    pricingPolicy: readJson(args.pricing),
    operationalReport: readJson(args["operational-report"]),
    exportedRecords: args.records ? readJson(args.records) : null,
  });

  const target = path.resolve(args.output);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, `${JSON.stringify(baseline, null, 2)}\n`, "utf8");

  console.log("PASS patch 15.3.2-F evaluation baseline built");
  console.log(`DECISION ${baseline.decision}`);
  console.log(`BASELINE_SHA256 ${baseline.baselineSha256}`);
  console.log(`EXECUTIONS ${baseline.coverage.executionCount}`);
  console.log(`LIFECYCLE ${baseline.coverage.lifecycleCount}`);
  console.log(`CASES ${baseline.coverage.caseCount}`);
  console.log(`OPERATIONAL_DECISION ${baseline.operational.decision}`);
  console.log("PRICING_MODE APPROVED_ACTUAL");
  console.log("EVALUATION_ONLY true");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
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
