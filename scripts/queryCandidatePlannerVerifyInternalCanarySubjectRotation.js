const fs = require("fs");
const crypto = require("crypto");

const {
  ROTATION_VERSION,
  EXPECTED_CANDIDATE_PAYLOAD_SHA256,
  normalizeSha256,
} = require("../automation/queryCandidatePlannerInternalCanarySubjectRotation");

function arg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? String(process.argv[index + 1] || "").trim() : "";
}

function sha256File(filePath) {
  return crypto
    .createHash("sha256")
    .update(fs.readFileSync(filePath))
    .digest("hex")
    .toUpperCase();
}

function main() {
  const planPath = arg("--plan");
  if (!planPath) throw new Error("--plan is required");
  if (!fs.existsSync(planPath)) throw new Error("ROTATION_PLAN_MISSING");

  const plan = JSON.parse(fs.readFileSync(planPath, "utf8"));
  if (plan.version !== ROTATION_VERSION)
    throw new Error("ROTATION_VERSION_INVALID");
  if (plan.valid !== true) throw new Error("ROTATION_PLAN_NOT_VALID");
  if (plan.decision !== "READY_FOR_LOCAL_APPROVAL_REBINDING") {
    throw new Error("ROTATION_DECISION_INVALID");
  }

  const oldSha = normalizeSha256(plan.currentAllowlistSha256);
  const newSha = normalizeSha256(plan.proposedAllowlistSha256);
  if (!oldSha || !newSha || oldSha === newSha)
    throw new Error("ROTATION_SHA_CONTRACT_INVALID");

  if (
    normalizeSha256(plan.immutableBindings?.candidatePayloadSha256) !==
    EXPECTED_CANDIDATE_PAYLOAD_SHA256
  ) {
    throw new Error("F_1_4_CANDIDATE_BINDING_DRIFT");
  }

  if (plan.approvalRebinding?.f15ReceiptReissueRequired !== true) {
    throw new Error("F_1_5_REISSUE_REQUIREMENT_MISSING");
  }

  const serialized = JSON.stringify(plan);
  for (const forbidden of [
    "immutableAccountId",
    "accountId",
    "tenantId",
    "email",
    "name",
    "userId",
    "rawAccountId",
    "rawTenantId",
  ]) {
    if (serialized.includes(`\"${forbidden}\"`)) {
      throw new Error(`RAW_IDENTITY_FIELD_FORBIDDEN_${forbidden}`);
    }
  }

  if (plan.guardrails?.railwayModified !== false)
    throw new Error("RAILWAY_MUTATION_FORBIDDEN");
  if (plan.guardrails?.routeModified !== false)
    throw new Error("ROUTE_MUTATION_FORBIDDEN");
  if (Number(plan.guardrails?.providerCallsExecutedByRotation) !== 0) {
    throw new Error("PROVIDER_CALL_FORBIDDEN");
  }
  if (plan.guardrails?.percentageRolloutAuthorized !== false) {
    throw new Error("ROLLOUT_AUTHORIZATION_FORBIDDEN");
  }
  if (plan.guardrails?.productionPromotionAuthorized !== false) {
    throw new Error("PRODUCTION_PROMOTION_FORBIDDEN");
  }

  console.log("PASS Patch 15.3.2-F.1.7 subject rotation verification");
  console.log(`ROTATION_PLAN_FILE_SHA256 ${sha256File(planPath)}`);
  console.log(`CURRENT_ALLOWLIST_SHA256 ${oldSha}`);
  console.log(`PROPOSED_ALLOWLIST_SHA256 ${newSha}`);
  console.log("F_1_4_CANDIDATE_PRESERVED true");
  console.log("F_1_5_RECEIPT_REISSUE_REQUIRED true");
  console.log("RAW_IDENTITY_INCLUDED false");
  console.log("RAILWAY_MODIFIED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_VERIFIER 0");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
}

try {
  main();
} catch (error) {
  console.error(`BLOCKED ${error.code || error.message}`);
  process.exitCode = 1;
}
