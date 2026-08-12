const fs = require("fs");
const path = require("path");

const {
  buildQueryCandidatePlannerInternalCanarySubjectRotation,
} = require("../automation/queryCandidatePlannerInternalCanarySubjectRotation");

function arg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? String(process.argv[index + 1] || "").trim() : "";
}

function requiredEnv(name) {
  const value = String(process.env[name] || "").trim();
  if (!value) throw new Error(`${name}_REQUIRED`);
  return value;
}

function main() {
  const accountId = requiredEnv("BEEBEE_CANARY_ROTATION_ACCOUNT_ID");
  const tenantId = String(
    process.env.BEEBEE_CANARY_ROTATION_TENANT_ID || "",
  ).trim();

  const currentAllowlistSha256 =
    arg("--current-allowlist-sha256") ||
    requiredEnv("QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256");

  const output = arg("--output");
  if (!output) throw new Error("--output is required");

  const request = {
    user: {
      accountId,
      ...(tenantId ? { tenantId } : {}),
    },
  };

  const plan = buildQueryCandidatePlannerInternalCanarySubjectRotation({
    currentAllowlistSha256,
    request,
  });

  if (!plan.valid) {
    console.error(`BLOCKED ${plan.reason}`);
    process.exitCode = 1;
    return;
  }

  const resolvedOutput = path.resolve(output);
  fs.mkdirSync(path.dirname(resolvedOutput), { recursive: true });
  fs.writeFileSync(resolvedOutput, `${JSON.stringify(plan, null, 2)}\n`, {
    encoding: "utf8",
    mode: 0o600,
  });

  console.log("PASS Patch 15.3.2-F.1.7 subject rotation prepared");
  console.log(`ROTATION_DECISION ${plan.decision}`);
  console.log(`CURRENT_ALLOWLIST_SHA256 ${plan.currentAllowlistSha256}`);
  console.log(`PROPOSED_ALLOWLIST_SHA256 ${plan.proposedAllowlistSha256}`);
  console.log(`SUBJECT_SOURCE ${plan.subject.source}`);
  console.log("RAW_IDENTITY_INCLUDED false");
  console.log("F_1_4_CANDIDATE_PRESERVED true");
  console.log("F_1_5_RECEIPT_REISSUE_REQUIRED true");
  console.log("RAILWAY_MODIFIED false");
  console.log("ENVIRONMENT_MODIFIED false");
  console.log("ROUTE_MODIFIED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_ROTATION 0");
  console.log("ACTUAL_OPERATIONAL_TELEMETRY false");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
}

try {
  main();
} catch (error) {
  console.error(`BLOCKED ${error.code || error.message}`);
  process.exitCode = 1;
}
