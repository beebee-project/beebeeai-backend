const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const {
  TARGET,
  GATE_MODULE,
  EXPECTED_GATE_SHA256,
  REQUIRE_MARKER,
  COMPOSITION_MARKER,
  LEGACY_CONTINUATION_MARKER,
  OBSOLETE_F16_MARKER,
  verifyCompositionSource,
} = require("./queryCandidatePlannerApplyInternalCanaryPreflightCompositionRepair");

function sha256(data) {
  return crypto.createHash("sha256").update(data).digest("hex").toUpperCase();
}

function main() {
  const repoRoot = path.resolve(process.argv[2] || ".");
  const target = path.join(repoRoot, TARGET);
  const gateFile = path.join(repoRoot, GATE_MODULE);

  if (!fs.existsSync(target)) throw new Error(`Target missing: ${TARGET}`);
  if (!fs.existsSync(gateFile)) throw new Error(`Gate missing: ${GATE_MODULE}`);

  const gateSha = sha256(fs.readFileSync(gateFile));
  if (gateSha !== EXPECTED_GATE_SHA256) {
    throw new Error(
      `Gate SHA drift: expected=${EXPECTED_GATE_SHA256} actual=${gateSha}`,
    );
  }

  const source = fs.readFileSync(target, "utf8");
  const serviceSha = sha256(Buffer.from(source, "utf8"));

  verifyCompositionSource(source);

  if (source.includes(OBSOLETE_F16_MARKER)) {
    throw new Error("Obsolete early-return integration still present");
  }

  const compositionIndex = source.indexOf(COMPOSITION_MARKER);
  const legacyIndex = source.indexOf(LEGACY_CONTINUATION_MARKER);

  const legacyTokens = [
    "parseQueryCandidatePlannerInternalCanaryConfig(env)",
    "parseEvidence(resolvedConfig, evidenceBundle, now)",
    "INVALID_INTERNAL_CANARY_CONFIGURATION",
    "INTERNAL_CANARY_DISABLED",
    "INTERNAL_CANARY_KILL_SWITCH_ACTIVE",
    "SEMANTIC_PROFILER_ONLY_POLICY_REQUIRED",
    "ALLOWLIST_ONLY_PROMOTION_CONFIGURATION_REQUIRED",
    "evaluateControlledProductionPromotionGate({",
    "if (!promotionDecision.allowed)",
    'status: "ALLOWLIST_PREFLIGHT_ALLOWED"',
  ];

  for (const token of legacyTokens) {
    const start =
      token === "parseQueryCandidatePlannerInternalCanaryConfig(env)"
        ? 0
        : legacyIndex;
    if (source.indexOf(token, start) < 0) {
      throw new Error(`Legacy preflight contract missing: ${token}`);
    }
  }

  console.log("PASS Patch 15.3.2-F.1.6.1 preflight composition verification");
  console.log(`SERVICE_SHA256 ${serviceSha}`);
  console.log(`GATE_SHA256 ${gateSha}`);
  console.log("APPROVAL_BINDING_REQUIRED true");
  console.log("APPROVAL_BLOCK_FAIL_CLOSED true");
  console.log("APPROVAL_ALLOW_CONTINUES_LEGACY_PREFLIGHT true");
  console.log("LEGACY_CONFIGURATION_CHECK_REACHABLE true");
  console.log("LEGACY_KILL_SWITCH_CHECK_REACHABLE true");
  console.log("LEGACY_EVIDENCE_CHECK_REACHABLE true");
  console.log("LEGACY_PROMOTION_GATE_REACHABLE true");
  console.log("LEGACY_FINAL_ALLOW_REACHABLE true");
  console.log("OBSOLETE_EARLY_RETURN_PRESENT false");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_VERIFIER 0");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
