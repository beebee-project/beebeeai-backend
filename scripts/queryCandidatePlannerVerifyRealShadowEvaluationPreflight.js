const fs = require("fs");
const path = require("path");
const {
  validateCollectionSummary,
  validateApprovedActualPricing,
} = require("../automation/queryCandidatePlannerRealShadowEvaluationBaseline");

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

function enabled(value) {
  return String(value || "").trim() === "1";
}

function main() {
  const input = args(process.argv.slice(2));
  if (!input["collection-summary"])
    throw new Error("--collection-summary is required");
  if (!input.pricing) throw new Error("--pricing is required");

  const summary = readJson(input["collection-summary"]);
  const pricing = readJson(input.pricing);
  const collection = validateCollectionSummary(summary);
  const priceFacts = validateApprovedActualPricing(pricing);

  const collectorRequested = enabled(
    process.env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED,
  );
  const collectorKillSwitch = enabled(
    process.env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH,
  );
  if (collectorRequested || !collectorKillSwitch) {
    const error = new Error(
      "Collector must be frozen: ENABLED=0 and KILL_SWITCH=1",
    );
    error.code = "COLLECTOR_NOT_FROZEN";
    throw error;
  }

  const internalCanary = enabled(
    process.env.QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED,
  );
  const promotionGate = enabled(
    process.env.QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED,
  );
  const productionEnabled = enabled(
    process.env.QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED,
  );

  if (internalCanary || promotionGate || productionEnabled) {
    const error = new Error("Canary/Promotion/Production must remain blocked");
    error.code = "PROMOTION_STATE_NOT_FAIL_CLOSED";
    throw error;
  }

  console.log("PASS patch 15.3.2-F evaluation preflight");
  console.log(`EXECUTIONS ${collection.executionCount}`);
  console.log(`LIFECYCLE ${collection.lifecycleCount}`);
  console.log(`CASES ${collection.caseCount}`);
  console.log(`EXPORT_TO ${collection.to}`);
  console.log(`PRICING_POLICY_ID ${priceFacts.policyId}`);
  console.log("PRICING_MODE APPROVED_ACTUAL");
  console.log("COLLECTOR_FROZEN true");
  console.log("INTERNAL_CANARY_ENABLED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log("READY_FOR_FROZEN_EXPORT true");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.code || error.message}`);
    process.exitCode = 1;
  }
}
