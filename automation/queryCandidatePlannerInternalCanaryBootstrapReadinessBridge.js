"use strict";

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const {
  OPERATIONS,
  evaluateReadinessGate,
} = require("./queryCandidatePlannerFeatureControl");

const BRIDGE_VERSION =
  "query_candidate_planner_internal_canary_bootstrap_readiness_bridge_v1";

const READINESS_SOURCE =
  "PATCH_13_3_HISTORICAL_LIVE_PROVIDER_CACHE_PARITY_READINESS";

const EXPECTED_HISTORICAL_READINESS_SOURCE_SHA256 =
  "33B70E7B4278CBC7E6F66D10CC6AA0F8FA7219A46E553EAD70612494E654F7D5";

const EXPECTED_READINESS_FILE_SHA256 =
  "46D1211AF4F318DAB91D137F0728C3AE6F246CD8B85A2582802CCB6DB1475AC4";

const EXPECTED_READINESS_GATE_SHA256 =
  "12FE722248FF2403A334FFBE735F97EEC7CC52DE7099A118C4144FD16D3E7823";

const DEFAULT_READINESS_FILE = path.join(
  __dirname,
  "..",
  "evaluation",
  "queryCandidatePlannerInternalCanaryBootstrapProductionReadiness.v1.json",
);

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

function invalid(reason, extra = {}) {
  return Object.freeze({
    version: BRIDGE_VERSION,
    valid: false,
    reason: String(reason || "BOOTSTRAP_READINESS_BRIDGE_INVALID"),
    failClosed: true,
    productionCandidateMergeOnly: true,
    historicalReadinessSourceSha256:
      EXPECTED_HISTORICAL_READINESS_SOURCE_SHA256,
    ...extra,
  });
}

function validateBootstrapProductionReadinessGate(readinessGate = null) {
  if (!readinessGate || typeof readinessGate !== "object") {
    return invalid("BOOTSTRAP_PRODUCTION_READINESS_OBJECT_REQUIRED");
  }

  const evaluated = evaluateReadinessGate(readinessGate);
  if (evaluated?.valid !== true) {
    return invalid(
      String(evaluated?.reason || "READINESS_EVIDENCE_INVALID"),
      { readinessEvaluation: evaluated || null },
    );
  }

  if (
    normalizeSha256(readinessGate.gateSha256) !==
    EXPECTED_READINESS_GATE_SHA256
  ) {
    return invalid("BOOTSTRAP_PRODUCTION_READINESS_GATE_SHA_MISMATCH", {
      readinessEvaluation: evaluated,
    });
  }

  const guardrails = readinessGate.guardrails || {};
  if (
    guardrails.shadowOnlyEvidence !== true ||
    guardrails.productionPromotionAllowed !== false ||
    guardrails.productionRouteAutoWired !== false ||
    guardrails.productionCandidateMergeAllowed !== false ||
    guardrails.productionReadyAssignmentAllowed !== false ||
    guardrails.manualPromotionReviewRequired !== true ||
    guardrails.failClosed !== true
  ) {
    return invalid("BOOTSTRAP_PRODUCTION_READINESS_GUARDRAIL_INVALID", {
      readinessEvaluation: evaluated,
    });
  }

  return Object.freeze({
    version: BRIDGE_VERSION,
    valid: true,
    reason: "BOOTSTRAP_PRODUCTION_READINESS_VALID",
    failClosed: true,
    productionCandidateMergeOnly: true,
    readinessGate,
    readinessEvaluation: evaluated,
    historicalReadinessSourceSha256:
      EXPECTED_HISTORICAL_READINESS_SOURCE_SHA256,
  });
}

function loadBootstrapProductionReadinessGate({
  file = DEFAULT_READINESS_FILE,
} = {}) {
  const resolved = path.resolve(file);
  if (!fs.existsSync(resolved)) {
    return invalid("BOOTSTRAP_PRODUCTION_READINESS_FILE_MISSING", {
      readinessFileSha256: "",
    });
  }

  const fileSha256 = sha256File(resolved);
  if (fileSha256 !== EXPECTED_READINESS_FILE_SHA256) {
    return invalid("BOOTSTRAP_PRODUCTION_READINESS_FILE_SHA_MISMATCH", {
      readinessFileSha256: fileSha256,
    });
  }

  let readinessGate;
  try {
    const text = fs.readFileSync(resolved, "utf8").replace(/^\uFEFF/, "");
    readinessGate = JSON.parse(text);
  } catch {
    return invalid("BOOTSTRAP_PRODUCTION_READINESS_JSON_INVALID", {
      readinessFileSha256: fileSha256,
    });
  }

  const validation = validateBootstrapProductionReadinessGate(readinessGate);
  if (!validation.valid) {
    return invalid(validation.reason, {
      readinessFileSha256: fileSha256,
      readinessEvaluation: validation.readinessEvaluation || null,
    });
  }

  return Object.freeze({
    ...validation,
    source: READINESS_SOURCE,
    readinessFileSha256: fileSha256,
  });
}

function createReadinessAwareApprovalFeatureControl({
  featureControl = null,
  readinessGate = null,
} = {}) {
  if (!featureControl || typeof featureControl.evaluate !== "function") {
    return invalid("BOOTSTRAP_APPROVAL_FEATURE_CONTROL_REQUIRED");
  }

  const validation = validateBootstrapProductionReadinessGate(readinessGate);
  if (!validation.valid) return validation;

  const bridgedFeatureControl = Object.freeze({
    evaluate(operation, options) {
      if (operation === OPERATIONS.PRODUCTION_CANDIDATE_MERGE) {
        return featureControl.evaluate(operation, {
          ...(options && typeof options === "object" ? options : {}),
          readinessGate: validation.readinessGate,
        });
      }
      return featureControl.evaluate(operation, options);
    },
  });

  return Object.freeze({
    version: BRIDGE_VERSION,
    valid: true,
    reason: "BOOTSTRAP_APPROVAL_FEATURE_CONTROL_READINESS_BRIDGED",
    failClosed: true,
    productionCandidateMergeOnly: true,
    featureControl: bridgedFeatureControl,
    readinessGate: validation.readinessGate,
    readinessEvaluation: validation.readinessEvaluation,
    historicalReadinessSourceSha256:
      EXPECTED_HISTORICAL_READINESS_SOURCE_SHA256,
  });
}

function resolveReadinessAwareApprovalFeatureControl({
  featureControl = null,
  readinessFile = DEFAULT_READINESS_FILE,
} = {}) {
  const loaded = loadBootstrapProductionReadinessGate({ file: readinessFile });
  if (!loaded.valid) return loaded;

  const bridged = createReadinessAwareApprovalFeatureControl({
    featureControl,
    readinessGate: loaded.readinessGate,
  });
  if (!bridged.valid) return bridged;

  return Object.freeze({
    ...bridged,
    source: loaded.source,
    readinessFileSha256: loaded.readinessFileSha256,
  });
}

module.exports = Object.freeze({
  BRIDGE_VERSION,
  READINESS_SOURCE,
  EXPECTED_HISTORICAL_READINESS_SOURCE_SHA256,
  EXPECTED_READINESS_FILE_SHA256,
  EXPECTED_READINESS_GATE_SHA256,
  DEFAULT_READINESS_FILE,
  normalizeSha256,
  sha256File,
  validateBootstrapProductionReadinessGate,
  loadBootstrapProductionReadinessGate,
  createReadinessAwareApprovalFeatureControl,
  resolveReadinessAwareApprovalFeatureControl,
});
