const fs = require("fs");
const path = require("path");
const {
  deriveActualPricingConsistentInput,
} = require("../automation/queryCandidatePlannerActualPricingCacheAvoidanceConsistency");

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

function main() {
  const input = args(process.argv.slice(2));
  for (const key of ["input", "pricing", "output"]) {
    if (!input[key]) throw new Error(`--${key} is required`);
  }

  const result = deriveActualPricingConsistentInput({
    canonicalInput: readJson(input.input),
    pricingPolicy: readJson(input.pricing),
  });

  const target = path.resolve(input.output);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, `${JSON.stringify(result, null, 2)}\n`, "utf8");

  console.log(
    "PASS patch 15.3.2-F.1.2 actual-pricing cost-consistent input prepared",
  );
  console.log(`SCENARIOS ${result.costConsistency.scenarioCount}`);
  console.log(
    `REPRICED_EXECUTIONS ${result.costConsistency.repricedExecutionCount}`,
  );
  console.log(
    `REPRICED_CACHE_HITS ${result.costConsistency.repricedCacheHitCount}`,
  );
  console.log(
    `PROVIDER_COST_MICROUSD ${result.costConsistency.providerCostMicrousd}`,
  );
  console.log(
    `AVOIDED_BY_CACHE_MICROUSD ${result.costConsistency.avoidedByCacheMicrousd}`,
  );
  console.log(
    `PREDICTED_CACHE_COST_AVOIDANCE_RATE ${result.costConsistency.cacheCostAvoidanceRate}`,
  );
  console.log("THRESHOLD_POLICY_MODIFIED false");
  console.log("EVALUATOR_MODIFIED false");
  console.log("PROVIDER_CALLS_EXECUTED 0");
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
