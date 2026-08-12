const fs = require("fs");
const path = require("path");

function parseArgs(argv) {
  const out = { rates: [] };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--model-rate") out.rates.push(argv[++i] || "");
    else if (arg.startsWith("--")) out[arg.slice(2)] = argv[++i] || "";
  }
  return out;
}

function positiveNumber(value, label) {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) throw new Error(`${label} must be > 0`);
  return n;
}

function parseRate(spec) {
  const parts = String(spec || "").split(":");
  if (parts.length !== 3 || !parts[0].trim()) {
    throw new Error(
      "--model-rate must be modelId:inputMicrousdPerMillion:outputMicrousdPerMillion",
    );
  }
  return {
    modelId: parts[0].trim(),
    input: positiveNumber(parts[1], "input rate"),
    output: positiveNumber(parts[2], "output rate"),
  };
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args["policy-id"] || args["policy-id"].includes("replace_with")) {
    throw new Error("--policy-id is required");
  }
  if (
    !args["effective-at"] ||
    !Number.isFinite(Date.parse(args["effective-at"]))
  ) {
    throw new Error("--effective-at must be an ISO date-time");
  }
  if (String(args["approve"] || "").toLowerCase() !== "true") {
    throw new Error("--approve true is required");
  }
  if (!args.output) throw new Error("--output is required");
  if (!args.rates.length)
    throw new Error("at least one --model-rate is required");

  const models = {};
  for (const spec of args.rates) {
    const rate = parseRate(spec);
    if (models[rate.modelId])
      throw new Error(`duplicate model rate: ${rate.modelId}`);
    models[rate.modelId] = {
      inputMicrousdPerMillionTokens: rate.input,
      outputMicrousdPerMillionTokens: rate.output,
    };
  }

  const policy = {
    version: "query_candidate_planner_cost_pricing_policy_v1",
    policyId: args["policy-id"],
    currency: "USD",
    mode: "APPROVED_ACTUAL",
    rateUnit: "MICROUSD_PER_MILLION_TOKENS",
    models,
    guardrails: {
      providerInvoice: false,
      productionBillingAuthority: false,
      approvedByOperator: true,
      effectiveAt: new Date(args["effective-at"]).toISOString(),
    },
  };

  const target = path.resolve(args.output);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, `${JSON.stringify(policy, null, 2)}\n`, "utf8");
  console.log("PASS approved actual pricing policy prepared");
  console.log(`POLICY_ID ${policy.policyId}`);
  console.log(`MODE ${policy.mode}`);
  console.log(`MODEL_COUNT ${Object.keys(models).length}`);
  console.log("PRODUCTION_BILLING_AUTHORITY false");
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
