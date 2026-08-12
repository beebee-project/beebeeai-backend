const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const TARGET =
  "automation/queryCandidatePlannerInternalAllowlistCanaryService.js";

const GATE_MODULE =
  "automation/queryCandidatePlannerInternalCanaryApprovalBindingGate.js";

const EXPECTED_SERVICE_SHA256 =
  "089E260D90625E068769F3D3538FAC198B4EAB3CEC4D864EF8CA9A747123E561";

const EXPECTED_GATE_SHA256 =
  "ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7";

const REQUIRE_MARKER =
  "// PATCH 15.3.2-F.1.6.1 APPROVAL BINDING COMPOSITION REQUIRE";

const COMPOSITION_MARKER =
  "// PATCH 15.3.2-F.1.6.1 APPROVAL BINDING COMPOSITION";

const LEGACY_CONTINUATION_MARKER =
  "// PATCH 15.3.2-F.1.6.1 LEGACY PREFLIGHT CONTINUES";

const OBSOLETE_F16_MARKER =
  "// PATCH 15.3.2-F.1.6 APPROVAL BINDING GATE INTEGRATION";

function sha256(data) {
  return crypto.createHash("sha256").update(data).digest("hex").toUpperCase();
}

function fileSha(file) {
  return sha256(fs.readFileSync(file));
}

function fail(message) {
  throw new Error(message);
}

function requireExactOne(source, token, label) {
  const count = source.split(token).length - 1;
  if (count !== 1) {
    fail(`${label} anchor count invalid: ${count}`);
  }
}

function insertRequire(source) {
  const requireBlock = [
    REQUIRE_MARKER,
    "const {",
    "  evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate,",
    '} = require("./queryCandidatePlannerInternalCanaryApprovalBindingGate");',
    "",
  ].join("\n");

  if (source.startsWith('"use strict";')) {
    const strictLine = '"use strict";';
    return (
      strictLine +
      "\n\n" +
      requireBlock +
      source.slice(strictLine.length).replace(/^\s*/, "")
    );
  }

  return requireBlock + source;
}

function buildComposedSource(source) {
  if (source.includes(OBSOLETE_F16_MARKER)) {
    fail(
      "Obsolete Patch 15.3.2-F.1.6 early-return integration is present. " +
        "Restore service to pre-integration SHA before F.1.6.1.",
    );
  }

  if (
    source.includes(REQUIRE_MARKER) ||
    source.includes(COMPOSITION_MARKER) ||
    source.includes(LEGACY_CONTINUATION_MARKER)
  ) {
    fail("F.1.6.1 composition markers already present");
  }

  const subjectAnchor =
    "  const subject = deriveQueryCandidatePlannerInternalCanarySubject(request);";

  const evidenceAnchor =
    "  const evidence = parseEvidence(resolvedConfig, evidenceBundle, now);";

  requireExactOne(source, subjectAnchor, "subject");
  requireExactOne(source, evidenceAnchor, "evidence");

  const subjectIndex = source.indexOf(subjectAnchor);
  const evidenceIndex = source.indexOf(evidenceAnchor);

  if (!(subjectIndex >= 0 && evidenceIndex > subjectIndex)) {
    fail("Legacy preflight subject/evidence order invalid");
  }

  source = insertRequire(source);

  const composition = [
    subjectAnchor,
    "",
    `  ${COMPOSITION_MARKER}`,
    "  const approvalBindingGate =",
    "    evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate({",
    "      env,",
    "      featureControl,",
    "      subject,",
    "    });",
    "",
    "  // Approval is an additional prerequisite only.",
    "  // BLOCK returns immediately; ALLOW must continue through every",
    "  // existing Patch 15.3 preflight check below.",
    "  if (!approvalBindingGate.allowed) {",
    "    return approvalBindingGate.preflight;",
    "  }",
    "",
    `  ${LEGACY_CONTINUATION_MARKER}`,
    evidenceAnchor,
  ].join("\n");

  source = source.replace(`${subjectAnchor}\n${evidenceAnchor}`, composition);

  if (!source.includes(COMPOSITION_MARKER)) {
    fail("Composition insertion failed");
  }

  const compositionIndex = source.indexOf(COMPOSITION_MARKER);
  const legacyMarkerIndex = source.indexOf(LEGACY_CONTINUATION_MARKER);

  if (!(legacyMarkerIndex > compositionIndex)) {
    fail("Legacy continuation marker order invalid");
  }

  const legacyRequiredAfterComposition = [
    "const evidence = parseEvidence(resolvedConfig, evidenceBundle, now);",
    "if (!resolvedConfig.configurationValid)",
    "if (!resolvedConfig.enabled)",
    "if (resolvedConfig.killSwitch)",
    "SEMANTIC_PROFILER_ONLY_POLICY_REQUIRED",
    "if (!subject.complete)",
    "if (!evidence.valid)",
    "const promotionConfig = parsePromotionGateEnvironment(env);",
    "const promotionDecision = evaluateControlledProductionPromotionGate({",
    "if (!promotionDecision.allowed)",
    'status: "ALLOWLIST_PREFLIGHT_ALLOWED"',
    'reason: "INTERNAL_ALLOWLIST_CANARY_PREFLIGHT_ALLOWED"',
  ];

  for (const token of legacyRequiredAfterComposition) {
    const index = source.indexOf(token, legacyMarkerIndex);
    if (index < 0) {
      fail(
        `Legacy preflight continuation token missing after approval gate: ${token}`,
      );
    }
  }

  const obsoleteEarlyReturnComment = "prior evidence path remains";
  if (source.includes(obsoleteEarlyReturnComment)) {
    fail("Obsolete early-return F.1.6 integration text detected");
  }

  return source;
}

function verifyCompositionSource(source) {
  const required = [
    REQUIRE_MARKER,
    COMPOSITION_MARKER,
    LEGACY_CONTINUATION_MARKER,
    "if (!approvalBindingGate.allowed) {",
    "return approvalBindingGate.preflight;",
  ];

  for (const token of required) {
    if (!source.includes(token)) {
      fail(`F.1.6.1 composition token missing: ${token}`);
    }
  }

  if (source.includes(OBSOLETE_F16_MARKER)) {
    fail("Obsolete F.1.6 integration marker remains");
  }

  const compositionIndex = source.indexOf(COMPOSITION_MARKER);
  const blockIndex = source.indexOf(
    "if (!approvalBindingGate.allowed) {",
    compositionIndex,
  );
  const conditionalReturnIndex = source.indexOf(
    "return approvalBindingGate.preflight;",
    blockIndex,
  );
  const legacyIndex = source.indexOf(
    LEGACY_CONTINUATION_MARKER,
    conditionalReturnIndex,
  );
  const evidenceIndex = source.indexOf(
    "const evidence = parseEvidence(resolvedConfig, evidenceBundle, now);",
    legacyIndex,
  );
  const promotionIndex = source.indexOf(
    "const promotionDecision = evaluateControlledProductionPromotionGate({",
    legacyIndex,
  );
  const finalAllowIndex = source.indexOf(
    'status: "ALLOWLIST_PREFLIGHT_ALLOWED"',
    legacyIndex,
  );

  if (
    !(
      compositionIndex >= 0 &&
      blockIndex > compositionIndex &&
      conditionalReturnIndex > blockIndex &&
      legacyIndex > conditionalReturnIndex &&
      evidenceIndex > legacyIndex &&
      promotionIndex > evidenceIndex &&
      finalAllowIndex > promotionIndex
    )
  ) {
    fail("F.1.6.1 approval/legacy composition order invalid");
  }

  const betweenReturnAndLegacy = source.slice(
    conditionalReturnIndex,
    legacyIndex,
  );

  if (!betweenReturnAndLegacy.includes("  }")) {
    fail("Approval BLOCK return is not closed before legacy continuation");
  }

  return {
    compositionIndex,
    legacyIndex,
    evidenceIndex,
    promotionIndex,
    finalAllowIndex,
  };
}

function main() {
  const repoRoot = path.resolve(process.argv[2] || ".");
  const target = path.join(repoRoot, TARGET);
  const gateFile = path.join(repoRoot, GATE_MODULE);

  if (!fs.existsSync(target)) {
    fail(`Target service missing: ${TARGET}`);
  }
  if (!fs.existsSync(gateFile)) {
    fail(`F.1.6 approval binding gate missing: ${GATE_MODULE}`);
  }

  const beforeSha = fileSha(target);
  const gateSha = fileSha(gateFile);

  if (gateSha !== EXPECTED_GATE_SHA256) {
    fail(
      `F.1.6 approval binding gate SHA drift: expected=${EXPECTED_GATE_SHA256} actual=${gateSha}`,
    );
  }

  const currentSource = fs.readFileSync(target, "utf8");

  if (
    currentSource.includes(REQUIRE_MARKER) &&
    currentSource.includes(COMPOSITION_MARKER) &&
    currentSource.includes(LEGACY_CONTINUATION_MARKER)
  ) {
    verifyCompositionSource(currentSource);
    console.log("PASS Patch 15.3.2-F.1.6.1 composition already applied");
    console.log(`SERVICE_SHA256 ${beforeSha}`);
    console.log(`GATE_SHA256 ${gateSha}`);
    console.log("IDEMPOTENT true");
    console.log("LEGACY_PREFLIGHT_CONTINUES true");
    console.log("OBSOLETE_EARLY_RETURN_PRESENT false");
    console.log("PROVIDER_CALLS_EXECUTED_BY_PATCH 0");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
    return;
  }

  if (beforeSha !== EXPECTED_SERVICE_SHA256) {
    fail(
      `Pre-integration service SHA mismatch: expected=${EXPECTED_SERVICE_SHA256} actual=${beforeSha}`,
    );
  }

  const nextSource = buildComposedSource(currentSource);
  verifyCompositionSource(nextSource);

  const backupRoot = path.join(repoRoot, ".patch_backups");
  fs.mkdirSync(backupRoot, { recursive: true });

  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

  const backupDir = path.join(
    backupRoot,
    `query_candidate_patch15_3_2_F_1_6_1_${timestamp}`,
  );
  fs.mkdirSync(backupDir, { recursive: true });

  const backupFile = path.join(backupDir, path.basename(target));

  fs.copyFileSync(target, backupFile);
  fs.writeFileSync(target, nextSource, "utf8");

  const afterSha = fileSha(target);

  console.log(
    "PASS Patch 15.3.2-F.1.6.1 existing canary preflight composition repair applied",
  );
  console.log(`TARGET ${TARGET}`);
  console.log(`BEFORE_SHA256 ${beforeSha}`);
  console.log(`AFTER_SHA256 ${afterSha}`);
  console.log(`GATE_SHA256 ${gateSha}`);
  console.log(`BACKUP ${path.relative(repoRoot, backupFile)}`);
  console.log("APPROVAL_BINDING_IS_ADDITIONAL_PREREQUISITE true");
  console.log("APPROVAL_BLOCK_RETURNS_IMMEDIATELY true");
  console.log("APPROVAL_ALLOW_CONTINUES_LEGACY_PREFLIGHT true");
  console.log("LEGACY_CONFIG_CHECK_PRESERVED true");
  console.log("LEGACY_KILL_SWITCH_CHECK_PRESERVED true");
  console.log("LEGACY_EVIDENCE_CHECK_PRESERVED true");
  console.log("LEGACY_PROMOTION_GATE_PRESERVED true");
  console.log("LEGACY_FEATURE_CONTROL_PRESERVED true");
  console.log("ROUTE_MODIFIED false");
  console.log("ENVIRONMENT_MODIFIED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_PATCH 0");
  console.log("PERCENTAGE_ROLLOUT_AUTHORIZED false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}

module.exports = Object.freeze({
  TARGET,
  GATE_MODULE,
  EXPECTED_SERVICE_SHA256,
  EXPECTED_GATE_SHA256,
  REQUIRE_MARKER,
  COMPOSITION_MARKER,
  LEGACY_CONTINUATION_MARKER,
  OBSOLETE_F16_MARKER,
  sha256,
  buildComposedSource,
  verifyCompositionSource,
});
