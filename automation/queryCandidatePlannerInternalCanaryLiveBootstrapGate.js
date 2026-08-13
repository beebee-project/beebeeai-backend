"use strict";

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const {
  evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate,
} = require("./queryCandidatePlannerInternalCanaryApprovalBindingGate");
const {
  verifyFinalEvaluationEvidenceBundle,
} = require("./queryCandidatePlannerFinalEvaluationEvidenceBundle");
const {
  BRIDGE_VERSION: READINESS_BRIDGE_VERSION,
  EXPECTED_READINESS_FILE_SHA256,
  resolveReadinessAwareApprovalFeatureControl,
} = require("./queryCandidatePlannerInternalCanaryBootstrapReadinessBridge");

const GATE_VERSION =
  "query_candidate_planner_internal_canary_live_bootstrap_gate_v1";
const EXPECTED_G_BUNDLE_PAYLOAD_SHA256 =
  "918F5760F4869581842B7A778AD647E3D88775FDEB58677F80CADC3639C3928B";
const EXPECTED_ALLOWLIST_SHA256 =
  "35D88A2074548BB9A6DB6BD3415CEE3CD2024BE9896AE6EC23260DB9B859AB95";
const EXPECTED_APPROVAL_RECEIPT_PAYLOAD_SHA256 =
  "4F5BA14C79CCB0ADD5DED476335729E59546A244A301FDF12897B14A7A09EF81";
const EXPECTED_F16_GATE_SHA256 =
  "ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7";
const EXPECTED_G_MODULE_SHA256 =
  "439F29AC82D866EEADA3EDFBD8615892904ACD507E4F8D4D5161431E0449440A";
const EXPECTED_READINESS_BRIDGE_SHA256 =
  "97C8F0CE5742873B16D05411B645C60B68B85625B308186E88AD3442CF5D58E4";

const ENV = Object.freeze({
  enabled: "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_BOOTSTRAP_ENABLED",
  killSwitch:
    "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_BOOTSTRAP_KILL_SWITCH",
  bundleJson:
    "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_BOOTSTRAP_READINESS_JSON",
  bundleSha256:
    "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_BOOTSTRAP_READINESS_SHA256",
});

function normalizeSha256(value) {
  const normalized = String(value || "").trim().toUpperCase();
  return /^[A-F0-9]{64}$/.test(normalized) && !/^0{64}$/.test(normalized)
    ? normalized
    : "";
}

function sha256File(file) {
  return crypto
    .createHash("sha256")
    .update(fs.readFileSync(path.resolve(file)))
    .digest("hex")
    .toUpperCase();
}

function parseStrictBoolean(value) {
  const raw = String(value == null ? "" : value).trim().toLowerCase();
  if (raw === "1" || raw === "true") return { valid: true, value: true };
  if (raw === "0" || raw === "false") return { valid: true, value: false };
  return { valid: false, value: false };
}

function safeSubject(subject = {}) {
  const subjectSha256 = normalizeSha256(subject.subjectSha256);
  const subjectTagSha256 =
    normalizeSha256(subject.subjectTagSha256) ||
    (subjectSha256
      ? crypto
          .createHash("sha256")
          .update(`bootstrap-tag:${subjectSha256}`)
          .digest("hex")
          .toUpperCase()
      : "");
  return Object.freeze({
    complete: subject.complete === true && Boolean(subjectSha256),
    subjectSha256,
    subjectTagSha256,
    rawIdentityIncluded: false,
  });
}

function blocked(reason, { subject = {}, legacyPreflight = null } = {}) {
  const safe = safeSubject(subject);
  return Object.freeze({
    version: GATE_VERSION,
    allowed: false,
    decision: "BLOCK",
    reason: String(reason || "BOOTSTRAP_BLOCKED"),
    failClosed: true,
    subject: safe,
    legacyEvidence: Object.freeze({
      valid: legacyPreflight?.evidence?.valid === true,
      reason: String(
        legacyPreflight?.evidence?.reason || legacyPreflight?.reason || "",
      ),
      substituted: false,
    }),
    bootstrapReadiness: Object.freeze({
      valid: false,
      gBundlePayloadSha256: "",
      legacyEvidenceSubstitutionForbidden: true,
      actualTrafficEvidenceRequiredFor15_3_4: true,
    }),
    runtimeBootstrapExecutionEligible: false,
    actualInternalUserExposureExecuted: false,
    providerCallsExecutedByGate: 0,
    actualOperationalTelemetry: false,
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
  });
}

function verifyProtectedDependencies() {
  const f16Gate = path.join(
    __dirname,
    "queryCandidatePlannerInternalCanaryApprovalBindingGate.js",
  );
  const gModule = path.join(
    __dirname,
    "queryCandidatePlannerFinalEvaluationEvidenceBundle.js",
  );
  const readinessBridge = path.join(
    __dirname,
    "queryCandidatePlannerInternalCanaryBootstrapReadinessBridge.js",
  );

  if (
    !fs.existsSync(f16Gate) ||
    !fs.existsSync(gModule) ||
    !fs.existsSync(readinessBridge)
  ) {
    return { valid: false, reason: "BOOTSTRAP_PROTECTED_DEPENDENCY_MISSING" };
  }

  if (sha256File(f16Gate) !== EXPECTED_F16_GATE_SHA256) {
    return { valid: false, reason: "BOOTSTRAP_F16_GATE_SHA_DRIFT" };
  }
  if (sha256File(gModule) !== EXPECTED_G_MODULE_SHA256) {
    return { valid: false, reason: "BOOTSTRAP_G_MODULE_SHA_DRIFT" };
  }
  if (sha256File(readinessBridge) !== EXPECTED_READINESS_BRIDGE_SHA256) {
    return { valid: false, reason: "BOOTSTRAP_READINESS_BRIDGE_SHA_DRIFT" };
  }
  return { valid: true, reason: "OK" };
}

function parseAndVerifyGBundle(env = {}) {
  const expectedSha = normalizeSha256(env[ENV.bundleSha256]);
  if (expectedSha !== EXPECTED_G_BUNDLE_PAYLOAD_SHA256) {
    return { valid: false, reason: "BOOTSTRAP_G_BUNDLE_SHA_MISMATCH" };
  }

  const raw = String(env[ENV.bundleJson] || "").trim();
  if (!raw) {
    return { valid: false, reason: "BOOTSTRAP_G_BUNDLE_JSON_REQUIRED" };
  }

  let bundle;
  try {
    bundle = JSON.parse(raw);
  } catch {
    return { valid: false, reason: "BOOTSTRAP_G_BUNDLE_JSON_INVALID" };
  }

  try {
    verifyFinalEvaluationEvidenceBundle(bundle);
  } catch (error) {
    return {
      valid: false,
      reason: String(error?.code || error?.message || "BOOTSTRAP_G_BUNDLE_INVALID"),
    };
  }

  if (
    normalizeSha256(bundle.bundlePayloadSha256) !==
    EXPECTED_G_BUNDLE_PAYLOAD_SHA256
  ) {
    return { valid: false, reason: "BOOTSTRAP_G_BUNDLE_PAYLOAD_SHA_MISMATCH" };
  }
  if (
    normalizeSha256(bundle.immutableBindings?.allowlistSha256) !==
    EXPECTED_ALLOWLIST_SHA256
  ) {
    return { valid: false, reason: "BOOTSTRAP_G_ALLOWLIST_BINDING_MISMATCH" };
  }
  if (
    normalizeSha256(bundle.immutableBindings?.approvalReceiptPayloadSha256) !==
    EXPECTED_APPROVAL_RECEIPT_PAYLOAD_SHA256
  ) {
    return { valid: false, reason: "BOOTSTRAP_G_APPROVAL_BINDING_MISMATCH" };
  }

  if (
    bundle.readiness?.bootstrapOnly !== true ||
    bundle.readiness?.internalAllowlistOnly !== true ||
    Number(bundle.readiness?.rolloutPercent) !== 0 ||
    bundle.legacy15_3EvidenceContract?.satisfiedByThisBundle !== false ||
    bundle.legacy15_3EvidenceContract?.substitutionForbidden !== true ||
    bundle.readiness?.actualTrafficEvidenceRequiredFor15_3_4 !== true ||
    bundle.authorizationBoundary?.runtimeAutoActivationAuthorized !== false ||
    bundle.authorizationBoundary?.actualInternalUserExposureAuthorized !== false ||
    bundle.authorizationBoundary?.percentageRolloutAuthorized !== false ||
    bundle.authorizationBoundary?.productionPromotionAuthorized !== false
  ) {
    return { valid: false, reason: "BOOTSTRAP_G_BOUNDARY_INVALID" };
  }

  return { valid: true, reason: "OK", bundle };
}

function evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapGate({
  env = process.env,
  featureControl = null,
  subject = {},
  legacyPreflight = null,
} = {}) {
  const safe = safeSubject(subject);

  // Bootstrap kill switch has highest precedence and must be explicitly OFF.
  const kill = parseStrictBoolean(env[ENV.killSwitch]);
  if (!kill.valid || kill.value) {
    return blocked(
      kill.valid
        ? "BOOTSTRAP_KILL_SWITCH_ACTIVE"
        : "BOOTSTRAP_KILL_SWITCH_EXPLICIT_VALUE_REQUIRED",
      { subject, legacyPreflight },
    );
  }

  const enabled = parseStrictBoolean(env[ENV.enabled]);
  if (!enabled.valid || !enabled.value) {
    return blocked(
      enabled.valid
        ? "BOOTSTRAP_DISABLED"
        : "BOOTSTRAP_ENABLED_EXPLICIT_VALUE_REQUIRED",
      { subject, legacyPreflight },
    );
  }

  if (!safe.complete) {
    return blocked("BOOTSTRAP_IMMUTABLE_SUBJECT_REQUIRED", {
      subject,
      legacyPreflight,
    });
  }

  // Never replace valid legacy evidence. Bootstrap is only the bridge for the
  // exact pre-canary readiness gap established at the end of Patch F/G.
  if (legacyPreflight?.allowed === true) {
    return blocked("BOOTSTRAP_NOT_REQUIRED_LEGACY_PREFLIGHT_ALREADY_ALLOWED", {
      subject,
      legacyPreflight,
    });
  }
  if (
    String(legacyPreflight?.reason || "") !== "READINESS_EVIDENCE_INVALID" ||
    legacyPreflight?.evidence?.valid === true
  ) {
    return blocked("BOOTSTRAP_ONLY_FOR_READINESS_EVIDENCE_INVALID", {
      subject,
      legacyPreflight,
    });
  }

  const deps = verifyProtectedDependencies();
  if (!deps.valid) {
    return blocked(deps.reason, { subject, legacyPreflight });
  }

  const g = parseAndVerifyGBundle(env);
  if (!g.valid) {
    return blocked(g.reason, { subject, legacyPreflight });
  }

  // Patch 15.3.3-B-4-F-A.2: F.1.6 intentionally remains byte-identical.
  // Bridge only its PRODUCTION_CANDIDATE_MERGE Feature Control evaluation
  // to the exact sanitized Patch 13.3 readiness snapshot. No legacy evidence
  // is substituted and the live bootstrap runtime remains observe-only/no-merge.
  const readinessBridge =
    resolveReadinessAwareApprovalFeatureControl({ featureControl });
  if (!readinessBridge.valid) {
    return blocked(
      `BOOTSTRAP_APPROVAL_READINESS_BRIDGE_BLOCKED_${String(
        readinessBridge.reason || "UNKNOWN",
      )}`,
      { subject, legacyPreflight },
    );
  }

  const approval =
    evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate({
      env,
      featureControl: readinessBridge.featureControl,
      subject: safe,
    });

  if (approval.allowed !== true || approval.decision !== "ALLOW") {
    return blocked(
      `BOOTSTRAP_APPROVAL_BINDING_BLOCKED_${String(approval.reason || "UNKNOWN")}`,
      { subject, legacyPreflight },
    );
  }

  if (safe.subjectSha256 !== EXPECTED_ALLOWLIST_SHA256) {
    return blocked("BOOTSTRAP_SUBJECT_NOT_EXACT_APPROVED_ALLOWLIST", {
      subject,
      legacyPreflight,
    });
  }

  if (
    approval.preflight?.promotionDecision?.allowed !== true ||
    approval.preflight?.promotionDecision?.decision !== "ALLOW" ||
    approval.preflight?.promotionDecision?.operation !==
      "PRODUCTION_CANDIDATE_MERGE" ||
    approval.preflight?.promotionDecision?.failClosed !== true ||
    approval.preflight?.promotionDecision?.adapterVersion !==
      "query_candidate_planner_controlled_production_merge_adapter_v1"
  ) {
    return blocked("BOOTSTRAP_APPROVAL_PROMOTION_CONTRACT_INVALID", {
      subject,
      legacyPreflight,
    });
  }

  return Object.freeze({
    version: GATE_VERSION,
    allowed: true,
    decision: "ALLOW",
    reason: "READY_FOR_SINGLE_SUBJECT_INTERNAL_CANARY_LIVE_BOOTSTRAP",
    failClosed: true,
    subject: safe,
    legacyEvidence: Object.freeze({
      valid: false,
      reason: "READINESS_EVIDENCE_INVALID",
      substituted: false,
    }),
    bootstrapReadiness: Object.freeze({
      valid: true,
      source: "PATCH_15_3_2_G_BOOTSTRAP_READINESS",
      gBundlePayloadSha256: EXPECTED_G_BUNDLE_PAYLOAD_SHA256,
      allowlistSha256: EXPECTED_ALLOWLIST_SHA256,
      approvalReceiptPayloadSha256:
        EXPECTED_APPROVAL_RECEIPT_PAYLOAD_SHA256,
      legacyEvidenceSubstitutionForbidden: true,
      actualTrafficEvidenceRequiredFor15_3_4: true,
      productionReadinessBridgeVersion: READINESS_BRIDGE_VERSION,
      productionReadinessFileSha256: EXPECTED_READINESS_FILE_SHA256,
    }),
    approvalBinding: approval.preflight?.approvalBinding || null,
    promotionDecision: approval.preflight?.promotionDecision || null,
    runtimeBootstrapExecutionEligible: true,
    actualInternalUserExposureExecuted: false,
    providerCallsExecutedByGate: 0,
    actualOperationalTelemetry: false,
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
    guardrails: Object.freeze({
      internalAllowlistOnly: true,
      singleApprovedSubjectOnly: true,
      rolloutPercent: 0,
      killSwitchRequired: true,
      primaryFallbackRequired: true,
      semanticProfilerOnly: true,
      legacyEvidenceUntouched: true,
      generalUsersBlocked: true,
      routeModifiedByThisGate: false,
      environmentModifiedByThisGate: false,
      providerCalledByThisGate: false,
      failClosed: true,
    }),
  });
}

module.exports = Object.freeze({
  GATE_VERSION,
  EXPECTED_G_BUNDLE_PAYLOAD_SHA256,
  EXPECTED_ALLOWLIST_SHA256,
  EXPECTED_APPROVAL_RECEIPT_PAYLOAD_SHA256,
  EXPECTED_F16_GATE_SHA256,
  EXPECTED_G_MODULE_SHA256,
  EXPECTED_READINESS_BRIDGE_SHA256,
  ENV,
  normalizeSha256,
  parseStrictBoolean,
  verifyProtectedDependencies,
  parseAndVerifyGBundle,
  evaluateQueryCandidatePlannerInternalCanaryLiveBootstrapGate,
});
