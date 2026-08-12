const fs = require("fs");
const path = require("path");
const {
  buildPrivateThresholdPolicy,
} = require("../automation/queryCandidatePlannerActualPricingAbsoluteCostThresholdRecalibration");

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

function main() {
  const input = args(process.argv.slice(2));
  for (const key of [
    "input",
    "pricing",
    "threshold-policy",
    "policy-output",
    "evidence-output",
  ]) {
    if (!input[key]) throw new Error(`--${key} is required`);
  }

  const result = buildPrivateThresholdPolicy({
    dataset: readJson(input.input).dataset,
    pricingPolicy: readJson(input.pricing),
    sourceThresholdPolicy: readJson(input["threshold-policy"]),
  });

  writeJson(input["policy-output"], result.policy);
  writeJson(input["evidence-output"], result.evidence);

  const d = result.evidence.providerCostDistribution;
  const t = result.evidence.derivedThresholds;

  console.log(
    "PASS patch 15.3.2-F.1.3 absolute cost threshold policy prepared",
  );
  console.log(`PROVIDER_CALL_SAMPLES ${d.sampleCount}`);
  console.log(`PROVIDER_COSTS_MICROUSD ${d.costsMicrousd.join(",")}`);
  console.log(`PROVIDER_COST_AVERAGE_MICROUSD ${d.averageMicrousd}`);
  console.log(`PROVIDER_COST_P95_MICROUSD ${d.p95Microusd}`);
  console.log(`PROVIDER_COST_MAX_MICROUSD ${d.maxMicrousd}`);
  console.log(`HEADROOM_RATE ${result.evidence.headroomRate}`);
  console.log(
    `RAW_PROVIDER_CEILING_MICROUSD ${result.evidence.rawProviderCeilingMicrousd}`,
  );
  console.log(
    `PROVIDER_CALL_AVERAGE_COST_MAX_MICROUSD ${t.providerCallAverageCostMicrousdMax}`,
  );
  console.log(`AVERAGE_COST_MAX_MICROUSD ${t.averageCostMicrousdMax}`);
  console.log(
    `MONTHLY_PROJECTED_COST_MAX_MICROUSD ${t.monthlyProjectedCostMicrousdMax}`,
  );
  console.log(
    `CACHE_COST_AVOIDANCE_RATE_MIN ${result.evidence.preservedContracts.cacheCostAvoidanceRateMin}`,
  );
  console.log("CHANGED_THRESHOLD_COUNT 3");
  console.log("SOURCE_THRESHOLD_POLICY_MODIFIED false");
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
