const fs = require("fs");
const path = require("path");
const {
  deriveCanonicalEvaluationDataset,
} = require("../automation/queryCandidatePlannerCanonicalEvaluationInput");

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
  for (const name of ["dataset", "pricing", "readiness", "output"]) {
    if (!args[name]) throw new Error(`--${name} is required`);
  }

  const result = deriveCanonicalEvaluationDataset({
    dataset: readJson(args.dataset),
    pricingPolicy: readJson(args.pricing),
    liveParityReadiness: readJson(args.readiness),
  });

  const target = path.resolve(args.output);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, `${JSON.stringify(result, null, 2)}\n`, "utf8");

  console.log("PASS patch 15.3.2-F.1 canonical evaluation input prepared");
  console.log(`MODE ${result.benchmarkMode}`);
  console.log(`EXECUTIONS ${result.dataset.executions.length}`);
  console.log(`LIFECYCLE ${result.dataset.lifecycleEvents.length}`);
  console.log(
    `STRIPPED_SYNTHETIC_OBSERVED_COST ${result.dataset.compatibility.strippedObservedCostCount}`,
  );
  console.log("APPROVED_ACTUAL_PRICING true");
  console.log("ACTUAL_LIVE_PROVIDER_PARITY_EVIDENCE true");
  console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
  console.log("PATCH_E_SUMMARY_USED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_PREPARATION 0");
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
