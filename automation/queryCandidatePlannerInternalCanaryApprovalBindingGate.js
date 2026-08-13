"use strict";

const crypto = require("crypto");

const GATE_VERSION =
  "query_candidate_planner_internal_canary_approval_binding_gate_v1";

const PREFLIGHT_VERSION =
  "query_candidate_planner_internal_allowlist_canary_result_v1";

const RECEIPT_VERSION =
  "query_candidate_planner_internal_canary_manual_approval_receipt_v1";

const RECEIPT_SCOPE = "INTERNAL_ALLOWLIST_CANARY_ONLY";

const APPROVAL_DECISION =
  "INTERNAL_ALLOWLIST_CANARY_MANUAL_APPROVAL_GRANTED";

const EXPECTED_CANDIDATE_PAYLOAD_SHA256 =
  "928F6A6E0AA8683D63A5A2CB62199FA460EB84494B119EB7E171000843D484EA";

const ENV = Object.freeze({
  receiptJson:
    "QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_RECEIPT_JSON",
  approvalBundleSha256:
    "QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_BUNDLE_SHA256",
  allowlistSha256:
    "QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256",

  internalCanaryEnabled:
    "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED",
  internalCanaryKillSwitch:
    "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH",
  internalCanaryLlmMode:
    "QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_LLM_MODE",

  globalKillSwitch:
    "QUERY_CANDIDATE_PLANNER_KILL_SWITCH",
  featureEnabled:
    "QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED",
  shadowEnabled:
    "QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED",
  providerEnabled:
    "QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED",
  providerKillSwitch:
    "QUERY_CANDIDATE_PLANNER_PROVIDER_KILL_SWITCH",

  productionEnabled:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED",
  productionCandidateMergeEnabled:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED",
  productionReadyAssignmentEnabled:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED",
  productionRouteEnabled:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED",
  productionKillSwitch:
    "QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH",

  promotionGateEnabled:
    "QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED",
  promotionAudienceMode:
    "QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE",
  promotionRolloutPercent:
    "QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT",
});

const REQUIRED_TRUE = Object.freeze([
  ENV.internalCanaryEnabled,
  ENV.featureEnabled,
  ENV.shadowEnabled,
  ENV.providerEnabled,
  ENV.productionEnabled,
  ENV.productionCandidateMergeEnabled,
  ENV.promotionGateEnabled,
]);

const REQUIRED_FALSE = Object.freeze([
  ENV.internalCanaryKillSwitch,
  ENV.globalKillSwitch,
  ENV.providerKillSwitch,
  ENV.productionKillSwitch,
  ENV.productionReadyAssignmentEnabled,
  ENV.productionRouteEnabled,
]);

const REQUIRED_FEATURE_OPERATIONS = Object.freeze([
  "SHADOW_EXECUTION",
  "PROVIDER_CALL",
  "PRODUCTION_CANDIDATE_MERGE",
]);

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (!isObject(value)) return value;
  const out = {};
  for (const key of Object.keys(value).sort()) out[key] = stable(value[key]);
  return out;
}

function canonicalJson(value) {
  return JSON.stringify(stable(value));
}

function sha256Json(value) {
  return crypto
    .createHash("sha256")
    .update(canonicalJson(value))
    .digest("hex")
    .toUpperCase();
}

function sha256Text(value) {
  return crypto
    .createHash("sha256")
    .update(String(value || ""))
    .digest("hex")
    .toUpperCase();
}

function normalizeSha256(value) {
  const normalized = String(value || "").trim().toUpperCase();
  return /^[A-F0-9]{64}$/.test(normalized) && !/^0{64}$/.test(normalized)
    ? normalized
    : "";
}

function parseStrictBoolean(value) {
  const raw = String(value == null ? "" : value).trim().toLowerCase();
  if (raw === "1" || raw === "true") return { valid: true, value: true };
  if (raw === "0" || raw === "false") return { valid: true, value: false };
  return { valid: false, value: false };
}

function parseAllowlist(value) {
  const entries = String(value || "")
    .split(/[,;\s]+/)
    .map((item) => normalizeSha256(item))
    .filter(Boolean);
  return [...new Set(entries)];
}

function safeSubject(subject = {}) {
  const subjectSha256 = normalizeSha256(subject.subjectSha256);
  const subjectTagSha256 =
    normalizeSha256(subject.subjectTagSha256) ||
    (subjectSha256
      ? sha256Text(`safe-tag:${subjectSha256}`)
      : "");

  return Object.freeze({
    complete: subject.complete === true && Boolean(subjectSha256),
    subjectSha256,
    subjectTagSha256,
    rawIdentityIncluded: false,
  });
}

function blocked({
  reason,
  subject,
  receipt = null,
  approvalBundleSha256 = "",
  allowlistCount = 0,
  featureControl = null,
} = {}) {
  const safe = safeSubject(subject);
  const candidateSha =
    normalizeSha256(receipt?.immutableBindings?.candidatePayloadSha256) || "";
  const receiptPayloadSha =
    normalizeSha256(receipt?.approvalReceiptPayloadSha256) || "";

  const promotionDecision = Object.freeze({
    version:
      "query_candidate_planner_controlled_production_promotion_gate_decision_v1",
    allowed: false,
    decision: "BLOCK",
    operation: "PRODUCTION_CANDIDATE_MERGE",
    failClosed: true,
    adapterVersion:
      "query_candidate_planner_controlled_production_merge_adapter_v1",
    reason: String(reason || "F_1_6_APPROVAL_BINDING_BLOCKED"),
    audience: Object.freeze({
      path: "ALLOWLIST",
      allowlistMatched: false,
      allowlistCount: Number(allowlistCount || 0),
      rolloutPercent: 0,
    }),
  });

  const evidence = Object.freeze({
    valid: false,
    reason: String(reason || "F_1_6_APPROVAL_BINDING_BLOCKED"),
    evidenceSha256: receiptPayloadSha || approvalBundleSha256 || "",
    summary: Object.freeze({
      source:
        "F_1_6_MANUAL_APPROVAL_BINDING_GATE",
      candidatePayloadSha256: candidateSha,
      approvalReceiptPayloadSha256: receiptPayloadSha,
      actualOperationalTelemetry: false,
      historicalLiveProviderParityEvidence: true,
      internalCanaryOnly: true,
      productionPromotionAuthorized: false,
      canaryEvidenceCollectionRequired: true,
    }),
    rawEvidenceIncluded: false,
  });

  const preflight = Object.freeze({
    version: PREFLIGHT_VERSION,
    status: "BLOCKED",
    allowed: false,
    reason: String(reason || "F_1_6_APPROVAL_BINDING_BLOCKED"),
    subject: Object.freeze({
      complete: safe.complete,
      subjectSha256: safe.subjectSha256,
      subjectTagSha256: safe.subjectTagSha256,
      rawIdentityIncluded: false,
    }),
    evidence,
    promotionDecision,
    approvalBinding: Object.freeze({
      version: GATE_VERSION,
      valid: false,
      candidatePayloadSha256: candidateSha,
      approvalReceiptPayloadSha256: receiptPayloadSha,
      approvalBundleSha256: approvalBundleSha256 || "",
      allowlistMatched: false,
      internalCanaryApprovalGranted: false,
      runtimeCanaryAuthorized: false,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    }),
    guardrails: Object.freeze({
      allowlistOnly: true,
      generalUsersBlocked: true,
      deterministicRolloutEnabled: false,
      rolloutPercent: 0,
      primaryFallbackAvailable: true,
      controlledProductionMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      plannerEscalationAllowed: false,
      semanticProfilerOnly: true,
      failClosed: true,
    }),
  });

  return Object.freeze({
    version: GATE_VERSION,
    allowed: false,
    decision: "BLOCK",
    reason: String(reason || "F_1_6_APPROVAL_BINDING_BLOCKED"),
    failClosed: true,
    preflight,
    featureControlPresent: Boolean(featureControl),
    providerCallsExecutedByGate: 0,
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
  });
}

function parseReceiptFromEnv(env = {}) {
  const raw = String(env[ENV.receiptJson] || "").trim();
  if (!raw) {
    return {
      valid: false,
      reason: "F_1_6_APPROVAL_RECEIPT_JSON_REQUIRED",
      receipt: null,
    };
  }

  try {
    const receipt = JSON.parse(raw);
    return { valid: true, reason: "OK", receipt };
  } catch {
    return {
      valid: false,
      reason: "F_1_6_APPROVAL_RECEIPT_JSON_INVALID",
      receipt: null,
    };
  }
}

function validateReceipt(receipt = {}) {
  if (receipt.version !== RECEIPT_VERSION) {
    return { valid: false, reason: "F_1_6_RECEIPT_VERSION_INVALID" };
  }
  if (receipt.scope !== RECEIPT_SCOPE) {
    return { valid: false, reason: "F_1_6_RECEIPT_SCOPE_INVALID" };
  }
  if (receipt.decision !== APPROVAL_DECISION) {
    return { valid: false, reason: "F_1_6_RECEIPT_DECISION_INVALID" };
  }

  const candidatePayloadSha256 = normalizeSha256(
    receipt.immutableBindings?.candidatePayloadSha256,
  );
  if (
    candidatePayloadSha256 !== EXPECTED_CANDIDATE_PAYLOAD_SHA256
  ) {
    return {
      valid: false,
      reason: "F_1_6_CANDIDATE_PAYLOAD_SHA_MISMATCH",
    };
  }

  const allowlistSha256 = normalizeSha256(
    receipt.immutableBindings?.allowlistSha256,
  );
  if (!allowlistSha256) {
    return {
      valid: false,
      reason: "F_1_6_RECEIPT_ALLOWLIST_SHA_INVALID",
    };
  }

  const observedReceiptSha256 = normalizeSha256(
    receipt.approvalReceiptPayloadSha256,
  );
  if (!observedReceiptSha256) {
    return {
      valid: false,
      reason: "F_1_6_RECEIPT_PAYLOAD_SHA_INVALID",
    };
  }

  const copy = JSON.parse(JSON.stringify(receipt));
  delete copy.approvalReceiptPayloadSha256;
  const recalculated = sha256Json(copy);
  if (recalculated !== observedReceiptSha256) {
    return {
      valid: false,
      reason: "F_1_6_RECEIPT_PAYLOAD_INTEGRITY_INVALID",
    };
  }

  if (
    receipt.manualApproval?.approvedByOperator !== true ||
    receipt.manualApproval?.evidenceBundleReviewed !== true ||
    receipt.manualApproval?.allowlistHashReviewed !== true ||
    receipt.manualApproval?.approvalIsRuntimeActivation !== false
  ) {
    return {
      valid: false,
      reason: "F_1_6_MANUAL_APPROVAL_CONTRACT_INVALID",
    };
  }

  if (
    receipt.evidenceSnapshot?.operationalDecision !== "EVALUATION_PASS" ||
    receipt.evidenceSnapshot?.assessmentDecision !==
      "ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS" ||
    Number(receipt.evidenceSnapshot?.failedCheckCount) !== 0 ||
    Number(receipt.evidenceSnapshot?.absoluteCostFailureCount) !== 0 ||
    receipt.evidenceSnapshot?.cacheCostAvoidancePassed !== true ||
    receipt.evidenceSnapshot?.evaluatorWorktreeEqualsHead !== true
  ) {
    return {
      valid: false,
      reason: "F_1_6_RECEIPT_EVIDENCE_SNAPSHOT_INVALID",
    };
  }

  if (receipt.evidenceSnapshot?.actualOperationalTelemetry !== false) {
    return {
      valid: false,
      reason: "F_1_6_UNEXPECTED_OPERATIONAL_TELEMETRY_CLAIM",
    };
  }

  if (
    receipt.authorizationBoundary?.internalCanaryApprovalGranted !== true ||
    receipt.authorizationBoundary?.runtimeGateBindingApplied !== false ||
    receipt.authorizationBoundary?.runtimeCanaryAuthorized !== false ||
    receipt.authorizationBoundary?.percentageRolloutAuthorized !== false ||
    receipt.authorizationBoundary?.productionPromotionAuthorized !== false ||
    receipt.authorizationBoundary?.productionMergeAuthorized !== false
  ) {
    return {
      valid: false,
      reason: "F_1_6_RECEIPT_AUTHORIZATION_BOUNDARY_INVALID",
    };
  }

  if (
    receipt.guardrails?.noGateMutation !== true ||
    receipt.guardrails?.noEnvironmentMutation !== true ||
    receipt.guardrails?.noRouteMutation !== true ||
    receipt.guardrails?.noFeatureFlagMutation !== true ||
    receipt.guardrails?.noKillSwitchMutation !== true ||
    receipt.guardrails?.noAllowlistMutation !== true ||
    Number(receipt.guardrails?.providerCallsExecutedByReceiptBuilder) !== 0
  ) {
    return {
      valid: false,
      reason: "F_1_6_RECEIPT_GUARDRAIL_INVALID",
    };
  }

  const serialized = JSON.stringify(receipt);
  for (const forbidden of [
    '"immutableAccountId"',
    '"allowlistSubjects"',
    '"responseId"',
    '"inputTokens"',
    '"outputTokens"',
  ]) {
    if (serialized.includes(forbidden)) {
      return {
        valid: false,
        reason: "F_1_6_RECEIPT_PRIVACY_BOUNDARY_INVALID",
      };
    }
  }

  return {
    valid: true,
    reason: "OK",
    candidatePayloadSha256,
    allowlistSha256,
    approvalReceiptPayloadSha256: observedReceiptSha256,
  };
}

function resolveFeatureControl(featureControl) {
  if (featureControl && typeof featureControl.evaluate === "function") {
    return featureControl;
  }

  try {
    const runtime = require("./queryCandidatePlannerFeatureControlRuntime");
    if (typeof runtime.getQueryCandidatePlannerFeatureControl === "function") {
      return runtime.getQueryCandidatePlannerFeatureControl();
    }
  } catch {
    // fail closed below
  }

  return null;
}

function evaluateFeatureControl(control) {
  if (!control || typeof control.evaluate !== "function") {
    return {
      valid: false,
      reason: "F_1_6_FEATURE_CONTROL_UNAVAILABLE",
      decisions: [],
    };
  }

  const decisions = [];

  for (const operation of REQUIRED_FEATURE_OPERATIONS) {
    let decision;
    try {
      decision = control.evaluate(operation);
    } catch {
      return {
        valid: false,
        reason: "F_1_6_FEATURE_CONTROL_EVALUATION_FAILED",
        decisions,
      };
    }

    decisions.push({
      operation,
      allowed: decision?.allowed === true,
      reason: String(decision?.reason || ""),
    });

    if (decision?.allowed !== true) {
      return {
        valid: false,
        reason:
          String(decision?.reason || "") ||
          `F_1_6_FEATURE_CONTROL_BLOCKED_${operation}`,
        decisions,
      };
    }
  }

  return { valid: true, reason: "OK", decisions };
}

function evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate({
  env = process.env,
  featureControl = null,
  subject = {},
} = {}) {
  const safe = safeSubject(subject);

  if (!safe.complete) {
    return blocked({
      reason: "F_1_6_IMMUTABLE_CANARY_SUBJECT_REQUIRED",
      subject,
    });
  }

  const receiptParse = parseReceiptFromEnv(env);
  if (!receiptParse.valid) {
    return blocked({
      reason: receiptParse.reason,
      subject,
    });
  }

  const receipt = receiptParse.receipt;
  const receiptValidation = validateReceipt(receipt);
  if (!receiptValidation.valid) {
    return blocked({
      reason: receiptValidation.reason,
      subject,
      receipt,
    });
  }

  const approvalBundleSha256 = normalizeSha256(
    env[ENV.approvalBundleSha256],
  );
  if (!approvalBundleSha256) {
    return blocked({
      reason: "F_1_6_APPROVAL_BUNDLE_SHA_REQUIRED",
      subject,
      receipt,
    });
  }

  if (
    approvalBundleSha256 !==
    receiptValidation.approvalReceiptPayloadSha256
  ) {
    return blocked({
      reason: "F_1_6_APPROVAL_BUNDLE_SHA_MISMATCH",
      subject,
      receipt,
      approvalBundleSha256,
    });
  }

  const allowlist = parseAllowlist(env[ENV.allowlistSha256]);
  if (!allowlist.length) {
    return blocked({
      reason: "F_1_6_ALLOWLIST_REQUIRED",
      subject,
      receipt,
      approvalBundleSha256,
    });
  }

  if (!allowlist.includes(receiptValidation.allowlistSha256)) {
    return blocked({
      reason: "F_1_6_RECEIPT_ALLOWLIST_NOT_IN_RUNTIME_ALLOWLIST",
      subject,
      receipt,
      approvalBundleSha256,
      allowlistCount: allowlist.length,
    });
  }

  if (safe.subjectSha256 !== receiptValidation.allowlistSha256) {
    return blocked({
      reason: "F_1_6_REQUEST_SUBJECT_NOT_APPROVED",
      subject,
      receipt,
      approvalBundleSha256,
      allowlistCount: allowlist.length,
    });
  }

  if (!allowlist.includes(safe.subjectSha256)) {
    return blocked({
      reason: "F_1_6_REQUEST_SUBJECT_NOT_IN_ALLOWLIST",
      subject,
      receipt,
      approvalBundleSha256,
      allowlistCount: allowlist.length,
    });
  }

  for (const name of REQUIRED_TRUE) {
    const parsed = parseStrictBoolean(env[name]);
    if (!parsed.valid) {
      return blocked({
        reason: `F_1_6_INVALID_BOOLEAN_${name}`,
        subject,
        receipt,
        approvalBundleSha256,
        allowlistCount: allowlist.length,
      });
    }
    if (!parsed.value) {
      return blocked({
        reason: `F_1_6_REQUIRED_FLAG_DISABLED_${name}`,
        subject,
        receipt,
        approvalBundleSha256,
        allowlistCount: allowlist.length,
      });
    }
  }

  for (const name of REQUIRED_FALSE) {
    const parsed = parseStrictBoolean(env[name]);
    if (!parsed.valid) {
      return blocked({
        reason: `F_1_6_INVALID_BOOLEAN_${name}`,
        subject,
        receipt,
        approvalBundleSha256,
        allowlistCount: allowlist.length,
      });
    }
    if (parsed.value) {
      return blocked({
        reason: `F_1_6_KILL_OR_FORBIDDEN_FLAG_ACTIVE_${name}`,
        subject,
        receipt,
        approvalBundleSha256,
        allowlistCount: allowlist.length,
      });
    }
  }

  const audienceMode = String(
    env[ENV.promotionAudienceMode] || "",
  ).trim().toUpperCase();

  if (audienceMode !== "ALLOWLIST") {
    return blocked({
      reason: "F_1_6_ALLOWLIST_AUDIENCE_REQUIRED",
      subject,
      receipt,
      approvalBundleSha256,
      allowlistCount: allowlist.length,
    });
  }

  if (String(env[ENV.promotionRolloutPercent] || "").trim() !== "0") {
    return blocked({
      reason: "F_1_6_ROLLOUT_PERCENT_MUST_BE_ZERO",
      subject,
      receipt,
      approvalBundleSha256,
      allowlistCount: allowlist.length,
    });
  }

  const llmMode = String(env[ENV.internalCanaryLlmMode] || "")
    .trim()
    .toUpperCase();
  if (llmMode !== "SEMANTIC_PROFILER_ONLY") {
    return blocked({
      reason: "F_1_6_SEMANTIC_PROFILER_ONLY_REQUIRED",
      subject,
      receipt,
      approvalBundleSha256,
      allowlistCount: allowlist.length,
    });
  }

  const control = resolveFeatureControl(featureControl);
  const controlResult = evaluateFeatureControl(control);
  if (!controlResult.valid) {
    return blocked({
      reason: controlResult.reason,
      subject,
      receipt,
      approvalBundleSha256,
      allowlistCount: allowlist.length,
      featureControl: control,
    });
  }

  const promotionDecision = Object.freeze({
    version:
      "query_candidate_planner_controlled_production_promotion_gate_decision_v1",
    allowed: true,
    decision: "ALLOW",
    operation: "PRODUCTION_CANDIDATE_MERGE",
    failClosed: true,
    adapterVersion:
      "query_candidate_planner_controlled_production_merge_adapter_v1",
    reason: "F_1_6_INTERNAL_ALLOWLIST_MANUAL_APPROVAL_BINDING_ALLOWED",
    audience: Object.freeze({
      path: "ALLOWLIST",
      allowlistMatched: true,
      allowlistCount: allowlist.length,
      rolloutPercent: 0,
    }),
  });

  const evidence = Object.freeze({
    valid: true,
    reason:
      "F_1_6_MANUAL_APPROVED_CANONICAL_BENCHMARK_WITH_HISTORICAL_LIVE_PARITY",
    evidenceSha256:
      receiptValidation.approvalReceiptPayloadSha256,
    summary: Object.freeze({
      source:
        "CANONICAL_BENCHMARK_WITH_APPROVED_ACTUAL_PRICING_AND_HISTORICAL_LIVE_PARITY",
      candidatePayloadSha256:
        receiptValidation.candidatePayloadSha256,
      approvalReceiptPayloadSha256:
        receiptValidation.approvalReceiptPayloadSha256,
      actualOperationalTelemetry: false,
      historicalLiveProviderParityEvidence: true,
      internalCanaryOnly: true,
      internalCanaryManualApprovalGranted: true,
      runtimeGateBindingApplied: true,
      runtimeCanaryAuthorized: true,
      canaryEvidenceCollectionRequired: true,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    }),
    rawEvidenceIncluded: false,
  });

  const preflight = Object.freeze({
    version: PREFLIGHT_VERSION,
    status: "ALLOWLIST_PREFLIGHT_ALLOWED",
    allowed: true,
    reason:
      "F_1_6_INTERNAL_ALLOWLIST_MANUAL_APPROVAL_BINDING_ALLOWED",
    subject: Object.freeze({
      complete: true,
      subjectSha256: safe.subjectSha256,
      subjectTagSha256: safe.subjectTagSha256,
      rawIdentityIncluded: false,
    }),
    evidence,
    promotionDecision,
    approvalBinding: Object.freeze({
      version: GATE_VERSION,
      valid: true,
      candidatePayloadSha256:
        receiptValidation.candidatePayloadSha256,
      approvalReceiptPayloadSha256:
        receiptValidation.approvalReceiptPayloadSha256,
      approvalBundleSha256,
      allowlistMatched: true,
      internalCanaryApprovalGranted: true,
      runtimeGateBindingApplied: true,
      runtimeCanaryAuthorized: true,
      percentageRolloutAuthorized: false,
      productionPromotionAuthorized: false,
    }),
    guardrails: Object.freeze({
      allowlistOnly: true,
      generalUsersBlocked: true,
      deterministicRolloutEnabled: false,
      rolloutPercent: 0,
      primaryFallbackAvailable: true,
      controlledProductionMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      plannerEscalationAllowed: false,
      semanticProfilerOnly: true,
      failClosed: true,
    }),
  });

  return Object.freeze({
    version: GATE_VERSION,
    allowed: true,
    decision: "ALLOW",
    reason:
      "F_1_6_INTERNAL_ALLOWLIST_MANUAL_APPROVAL_BINDING_ALLOWED",
    failClosed: true,
    preflight,
    featureControlPresent: true,
    featureControlDecisions: Object.freeze(controlResult.decisions),
    providerCallsExecutedByGate: 0,
    actualOperationalTelemetry: false,
    canaryEvidenceCollectionRequired: true,
    runtimeCanaryAuthorized: true,
    percentageRolloutAuthorized: false,
    productionPromotionAuthorized: false,
  });
}

module.exports = Object.freeze({
  GATE_VERSION,
  PREFLIGHT_VERSION,
  RECEIPT_VERSION,
  RECEIPT_SCOPE,
  APPROVAL_DECISION,
  EXPECTED_CANDIDATE_PAYLOAD_SHA256,
  ENV,
  REQUIRED_TRUE,
  REQUIRED_FALSE,
  REQUIRED_FEATURE_OPERATIONS,
  canonicalJson,
  sha256Json,
  normalizeSha256,
  parseStrictBoolean,
  parseAllowlist,
  validateReceipt,
  evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate,
});
