"use strict";

const crypto = require("crypto");
const {
  parseAllowlist,
  parseRegistry,
} = require("./queryCandidatePlannerRealShadowEvidenceConfig");
const {
  encryptEvidencePayload,
  decryptEvidencePayload,
  ENCRYPTION_VERSION,
} = require("./queryCandidatePlannerRealShadowEvidenceCrypto");
const {
  verifyProductionSafety,
} = require("./queryCandidatePlannerRealShadowPreparation");

const SECURE_DEPLOYMENT_VERSION =
  "query_candidate_planner_real_shadow_secure_deployment_v1";
const FOUNDATION_SUMMARY_VERSION =
  "query_candidate_planner_real_shadow_evidence_foundation_summary_v1";
const SHA256_RE = /^[a-f0-9]{64}$/i;
const SECRET_FORMAT_RE = /^[A-Za-z0-9_-]{64}$/;

function text(value, maxLength = 500000) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function sha256(value) {
  return crypto
    .createHash("sha256")
    .update(typeof value === "string" ? value : JSON.stringify(value))
    .digest("hex");
}

function booleanValue(value, fallback = false) {
  const normalized = text(value, 20).toLowerCase();
  if (!normalized) return fallback;
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off"].includes(normalized)) return false;
  return fallback;
}

function foundationErrors(summary = {}) {
  const errors = [];
  if (!summary || typeof summary !== "object" || Array.isArray(summary)) {
    return ["REAL_SHADOW_FOUNDATION_SUMMARY_OBJECT_REQUIRED"];
  }
  if (summary.version !== FOUNDATION_SUMMARY_VERSION) {
    errors.push("REAL_SHADOW_FOUNDATION_SUMMARY_VERSION_INVALID");
  }
  if (summary.phase !== "15.3-A") {
    errors.push("REAL_SHADOW_FOUNDATION_PHASE_INVALID");
  }
  if (summary.decision !== "REAL_SHADOW_EVIDENCE_FOUNDATION_PASS") {
    errors.push("REAL_SHADOW_FOUNDATION_DECISION_NOT_PASS");
  }
  if (summary.readyForPatch15_3_2_C !== true) {
    errors.push("REAL_SHADOW_FOUNDATION_NOT_READY_FOR_PATCH_15_3_2_C");
  }
  if (summary.sourceCatalogComplete !== true) {
    errors.push("REAL_SHADOW_FOUNDATION_SOURCE_CATALOG_INCOMPLETE");
  }
  if (summary.ledgerComplete !== true) {
    errors.push("REAL_SHADOW_FOUNDATION_LEDGER_INCOMPLETE");
  }
  if (summary.registryFinalized !== true) {
    errors.push("REAL_SHADOW_FOUNDATION_REGISTRY_NOT_FINALIZED");
  }
  if (
    !Number.isInteger(summary.expectedCaseCount) ||
    summary.expectedCaseCount <= 0 ||
    summary.completedCaseCount !== summary.expectedCaseCount
  ) {
    errors.push("REAL_SHADOW_FOUNDATION_CASE_COVERAGE_INCOMPLETE");
  }
  if (
    !Number.isInteger(summary.expectedRejectionCaseCount) ||
    summary.expectedRejectionEvidenceCount !== summary.expectedRejectionCaseCount
  ) {
    errors.push("REAL_SHADOW_FOUNDATION_EXPECTED_REJECTION_INCOMPLETE");
  }
  for (const field of ["sourceCatalogSha256", "ledgerSha256", "registrySha256"]) {
    if (!SHA256_RE.test(text(summary[field], 256))) {
      errors.push(`REAL_SHADOW_FOUNDATION_${field.toUpperCase()}_INVALID`);
    }
  }
  if (summary.actualTrafficOnly !== true) {
    errors.push("REAL_SHADOW_FOUNDATION_ACTUAL_TRAFFIC_REQUIRED");
  }
  if (summary.syntheticFingerprintForbidden !== true) {
    errors.push("REAL_SHADOW_FOUNDATION_SYNTHETIC_FINGERPRINT_MUST_BE_FORBIDDEN");
  }
  if (summary.legacyLedgerAccepted !== false) {
    errors.push("REAL_SHADOW_FOUNDATION_LEGACY_LEDGER_MUST_BE_REJECTED");
  }
  if (summary.rawIdentityIncluded !== false) {
    errors.push("REAL_SHADOW_FOUNDATION_RAW_IDENTITY_FORBIDDEN");
  }
  if (summary.collectorEnabledByThisPhase !== false) {
    errors.push("REAL_SHADOW_FOUNDATION_COLLECTOR_STATE_INVALID");
  }
  if (summary.internalCanaryEnabledByThisPhase !== false) {
    errors.push("REAL_SHADOW_FOUNDATION_INTERNAL_CANARY_STATE_INVALID");
  }
  if (summary.productionPromotionAuthorized !== false) {
    errors.push("REAL_SHADOW_FOUNDATION_PRODUCTION_PROMOTION_MUST_BE_BLOCKED");
  }
  return errors;
}


function strictRegistryFingerprintErrors(registry = {}) {
  const errors = [];
  const cases = Array.isArray(registry?.cases) ? registry.cases : [];
  for (const [index, item] of cases.entries()) {
    const requestRaw = String(item?.requestFingerprintSha256 == null ? "" : item.requestFingerprintSha256).trim();
    const uploadRaw = String(item?.uploadFingerprintSha256 == null ? "" : item.uploadFingerprintSha256).trim();
    if (requestRaw && !SHA256_RE.test(requestRaw)) {
      errors.push(`REAL_SHADOW_DEPLOYMENT_REGISTRY_REQUEST_FINGERPRINT_INVALID:${index}`);
    }
    if (uploadRaw && !SHA256_RE.test(uploadRaw)) {
      errors.push(`REAL_SHADOW_DEPLOYMENT_REGISTRY_UPLOAD_FINGERPRINT_INVALID:${index}`);
    }
    if (!requestRaw && !uploadRaw) {
      errors.push(`REAL_SHADOW_DEPLOYMENT_REGISTRY_FINGERPRINT_REQUIRED:${index}`);
    }
  }
  return errors;
}

function runEncryptionSelfTest(secret) {
  if (!SECRET_FORMAT_RE.test(String(secret == null ? "" : secret).trim())) {
    return Object.freeze({ passed: false, reason: "REAL_SHADOW_EVIDENCE_SECRET_FORMAT_INVALID" });
  }
  const payload = Object.freeze({
    version: "secure_deployment_self_test_v1",
    kind: "SELF_TEST",
    marker: crypto.randomBytes(12).toString("hex"),
  });
  const encrypted = encryptEvidencePayload(payload, secret);
  const decrypted = decryptEvidencePayload(encrypted, secret);
  if (JSON.stringify(decrypted) !== JSON.stringify(payload)) {
    return Object.freeze({ passed: false, reason: "REAL_SHADOW_ENCRYPTION_ROUND_TRIP_FAILED" });
  }
  if (encrypted.ciphertext.includes(payload.marker)) {
    return Object.freeze({ passed: false, reason: "REAL_SHADOW_ENCRYPTION_PLAINTEXT_LEAK" });
  }
  const wrongSecret = secret[0] === "A"
    ? `B${secret.slice(1)}`
    : `A${secret.slice(1)}`;
  let wrongSecretRejected = false;
  try {
    decryptEvidencePayload(encrypted, wrongSecret);
  } catch (_error) {
    wrongSecretRejected = true;
  }
  if (!wrongSecretRejected) {
    return Object.freeze({ passed: false, reason: "REAL_SHADOW_WRONG_SECRET_NOT_REJECTED" });
  }
  return Object.freeze({
    passed: true,
    reason: "REAL_SHADOW_ENCRYPTION_SELF_TEST_PASS",
    encryptionVersion: ENCRYPTION_VERSION,
    ciphertextPlaintextFree: true,
    wrongSecretRejected: true,
  });
}

function evaluateRealShadowSecureDeployment({
  foundationSummary,
  registry,
  env = process.env,
} = {}) {
  const errors = [...foundationErrors(foundationSummary)];
  const rawSecret = text(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET,
    2000,
  );
  const secretFormatValid = SECRET_FORMAT_RE.test(rawSecret);
  if (!rawSecret) errors.push("REAL_SHADOW_EVIDENCE_SECRET_REQUIRED");
  if (rawSecret && !secretFormatValid) {
    errors.push("REAL_SHADOW_EVIDENCE_SECRET_FORMAT_INVALID");
  }

  const collectorEnabled = booleanValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED,
    false,
  );
  if (collectorEnabled) {
    errors.push("REAL_SHADOW_COLLECTOR_MUST_REMAIN_DISABLED_DURING_PATCH_15_3_2_C");
  }

  const allowlist = parseAllowlist(
    env.QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256,
  );
  if (allowlist.length === 0) {
    errors.push("REAL_SHADOW_DEPLOYMENT_ALLOWLIST_REQUIRED");
  }

  const strictRegistryErrors = strictRegistryFingerprintErrors(registry || {});
  errors.push(...strictRegistryErrors);
  const parsedRegistry = parseRegistry(JSON.stringify(registry || {}));
  if (!parsedRegistry.valid) {
    errors.push(`REAL_SHADOW_DEPLOYMENT_REGISTRY_INVALID:${parsedRegistry.reason}`);
  }

  const runtimeRegistryRaw = text(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON,
    500000,
  );
  let runtimeRegistryObject = null;
  try { runtimeRegistryObject = runtimeRegistryRaw ? JSON.parse(runtimeRegistryRaw) : null; } catch (_error) {}
  errors.push(...strictRegistryFingerprintErrors(runtimeRegistryObject || {}));
  const runtimeRegistry = parseRegistry(runtimeRegistryRaw);
  if (!runtimeRegistry.valid) {
    errors.push(`REAL_SHADOW_DEPLOYMENT_RUNTIME_REGISTRY_INVALID:${runtimeRegistry.reason}`);
  }

  const registrySha256 = parsedRegistry.valid
    ? sha256(JSON.stringify(registry))
    : "";
  const runtimeRegistrySha256 = runtimeRegistry.valid && runtimeRegistryObject
    ? sha256(JSON.stringify(runtimeRegistryObject))
    : "";
  const expectedRegistrySha256 = text(foundationSummary?.registrySha256, 256).toLowerCase();
  const configuredRegistrySha256 = text(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_SHA256,
    256,
  ).toLowerCase();
  if (!SHA256_RE.test(configuredRegistrySha256)) {
    errors.push("REAL_SHADOW_DEPLOYMENT_REGISTRY_SHA256_REQUIRED");
  }

  if (
    parsedRegistry.valid &&
    expectedRegistrySha256 &&
    registrySha256 !== expectedRegistrySha256
  ) {
    errors.push("REAL_SHADOW_DEPLOYMENT_REGISTRY_FOUNDATION_HASH_MISMATCH");
  }
  if (
    parsedRegistry.valid &&
    runtimeRegistry.valid &&
    runtimeRegistrySha256 !== registrySha256
  ) {
    errors.push("REAL_SHADOW_DEPLOYMENT_RUNTIME_REGISTRY_MISMATCH");
  }
  if (
    SHA256_RE.test(configuredRegistrySha256) &&
    configuredRegistrySha256 !== expectedRegistrySha256
  ) {
    errors.push("REAL_SHADOW_DEPLOYMENT_CONFIGURED_REGISTRY_HASH_FOUNDATION_MISMATCH");
  }
  if (
    SHA256_RE.test(configuredRegistrySha256) &&
    runtimeRegistrySha256 &&
    configuredRegistrySha256 !== runtimeRegistrySha256
  ) {
    errors.push("REAL_SHADOW_DEPLOYMENT_CONFIGURED_REGISTRY_HASH_RUNTIME_MISMATCH");
  }

  const productionSafety = verifyProductionSafety(env);
  if (!productionSafety.safe) {
    for (const failed of productionSafety.failed) {
      errors.push(`REAL_SHADOW_DEPLOYMENT_UNSAFE_PRODUCTION_FLAG:${failed}`);
    }
  }

  let encryptionSelfTest = Object.freeze({
    passed: false,
    reason: "REAL_SHADOW_ENCRYPTION_SELF_TEST_NOT_RUN",
  });
  if (secretFormatValid) {
    try {
      encryptionSelfTest = runEncryptionSelfTest(rawSecret);
      if (!encryptionSelfTest.passed) errors.push(encryptionSelfTest.reason);
    } catch (error) {
      errors.push(String(error?.code || "REAL_SHADOW_ENCRYPTION_SELF_TEST_FAILED"));
    }
  }

  const uniqueErrors = Object.freeze([...new Set(errors)]);
  const ready = uniqueErrors.length === 0;
  return Object.freeze({
    version: SECURE_DEPLOYMENT_VERSION,
    phase: "15.3-B",
    patch: "15.3.2-C",
    ready,
    reason: ready
      ? "REAL_SHADOW_SECURE_DEPLOYMENT_READY_FOR_LIMITED_COLLECTOR_ACTIVATION"
      : uniqueErrors[0] || "REAL_SHADOW_SECURE_DEPLOYMENT_BLOCKED",
    errors: uniqueErrors,
    foundationReady: foundationErrors(foundationSummary).length === 0,
    registrySha256,
    runtimeRegistrySha256,
    foundationRegistrySha256: expectedRegistrySha256,
    configuredRegistrySha256,
    registryCaseCount: parsedRegistry.valid && Array.isArray(registry?.cases) ? registry.cases.length : 0,
    allowlistEntryCount: allowlist.length,
    secretConfigured: Boolean(rawSecret),
    secretFormatValid,
    secretSha256: secretFormatValid ? sha256(rawSecret) : "",
    encryptionSelfTest,
    productionSafety,
    collectorEnabled: collectorEnabled,
    collectorEnabledByThisOperation: false,
    internalCanaryEnabledByThisOperation: false,
    productionPromotionAuthorized: false,
    rawSecretIncluded: false,
    rawIdentityIncluded: false,
    readyForPatch15_3_2_D: ready,
  });
}

function evaluateRealShadowSecureRuntime({ env = process.env } = {}) {
  const errors = [];
  const rawSecret = text(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET,
    2000,
  );
  const secretFormatValid = SECRET_FORMAT_RE.test(rawSecret);
  if (!rawSecret) errors.push("REAL_SHADOW_EVIDENCE_SECRET_REQUIRED");
  if (rawSecret && !secretFormatValid) errors.push("REAL_SHADOW_EVIDENCE_SECRET_FORMAT_INVALID");

  const collectorEnabled = booleanValue(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED,
    false,
  );
  if (collectorEnabled) {
    errors.push("REAL_SHADOW_COLLECTOR_MUST_REMAIN_DISABLED_DURING_PATCH_15_3_2_C");
  }

  const allowlist = parseAllowlist(env.QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256);
  if (allowlist.length === 0) errors.push("REAL_SHADOW_DEPLOYMENT_ALLOWLIST_REQUIRED");

  const runtimeRegistryRaw = text(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON,
    500000,
  );
  let runtimeRegistryObject = null;
  try { runtimeRegistryObject = runtimeRegistryRaw ? JSON.parse(runtimeRegistryRaw) : null; } catch (_error) {}
  errors.push(...strictRegistryFingerprintErrors(runtimeRegistryObject || {}));
  const runtimeRegistry = parseRegistry(runtimeRegistryRaw);
  if (!runtimeRegistry.valid) {
    errors.push(`REAL_SHADOW_DEPLOYMENT_RUNTIME_REGISTRY_INVALID:${runtimeRegistry.reason}`);
  }
  const runtimeRegistrySha256 = runtimeRegistry.valid && runtimeRegistryObject
    ? sha256(JSON.stringify(runtimeRegistryObject))
    : "";
  const configuredRegistrySha256 = text(
    env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_SHA256,
    256,
  ).toLowerCase();
  if (!SHA256_RE.test(configuredRegistrySha256)) {
    errors.push("REAL_SHADOW_DEPLOYMENT_REGISTRY_SHA256_REQUIRED");
  } else if (runtimeRegistrySha256 && configuredRegistrySha256 !== runtimeRegistrySha256) {
    errors.push("REAL_SHADOW_DEPLOYMENT_CONFIGURED_REGISTRY_HASH_RUNTIME_MISMATCH");
  }

  const productionSafety = verifyProductionSafety(env);
  if (!productionSafety.safe) {
    for (const failed of productionSafety.failed) {
      errors.push(`REAL_SHADOW_DEPLOYMENT_UNSAFE_PRODUCTION_FLAG:${failed}`);
    }
  }

  let encryptionSelfTest = Object.freeze({ passed: false, reason: "REAL_SHADOW_ENCRYPTION_SELF_TEST_NOT_RUN" });
  if (secretFormatValid) {
    encryptionSelfTest = runEncryptionSelfTest(rawSecret);
    if (!encryptionSelfTest.passed) errors.push(encryptionSelfTest.reason);
  }

  const uniqueErrors = Object.freeze([...new Set(errors)]);
  const ready = uniqueErrors.length === 0;
  return Object.freeze({
    version: SECURE_DEPLOYMENT_VERSION,
    phase: "15.3-B",
    patch: "15.3.2-C",
    runtimeOnly: true,
    ready,
    reason: ready
      ? "REAL_SHADOW_SECURE_RUNTIME_READY_FOR_LIMITED_COLLECTOR_ACTIVATION"
      : uniqueErrors[0] || "REAL_SHADOW_SECURE_RUNTIME_BLOCKED",
    errors: uniqueErrors,
    runtimeRegistrySha256,
    configuredRegistrySha256,
    registryCaseCount: runtimeRegistryObject && Array.isArray(runtimeRegistryObject.cases)
      ? runtimeRegistryObject.cases.length
      : 0,
    allowlistEntryCount: allowlist.length,
    secretConfigured: Boolean(rawSecret),
    secretFormatValid,
    secretSha256: secretFormatValid ? sha256(rawSecret) : "",
    encryptionSelfTest,
    productionSafety,
    collectorEnabled,
    collectorEnabledByThisOperation: false,
    internalCanaryEnabledByThisOperation: false,
    productionPromotionAuthorized: false,
    rawSecretIncluded: false,
    readyForPatch15_3_2_D: ready,
  });
}

module.exports = Object.freeze({
  SECURE_DEPLOYMENT_VERSION,
  FOUNDATION_SUMMARY_VERSION,
  sha256,
  foundationErrors,
  strictRegistryFingerprintErrors,
  runEncryptionSelfTest,
  evaluateRealShadowSecureDeployment,
  evaluateRealShadowSecureRuntime,
});
