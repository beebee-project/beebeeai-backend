"use strict";

const crypto = require("crypto");
const {
  exactSha256,
  validateRealShadowFingerprintLedger,
  finalizeRealShadowCaseRegistry,
} = require("./queryCandidatePlannerRealShadowRegistryFinalization");
const {
  validateUploadableSourceCatalog,
  sourceCatalogSha256,
} = require("./queryCandidatePlannerRealShadowUploadableSourceCatalog");

const FOUNDATION_VERSION =
  "query_candidate_planner_real_shadow_evidence_foundation_v1";
const ATTESTATION_VERSION =
  "query_candidate_planner_real_shadow_expected_rejection_attestation_v1";
const FOUNDATION_SUMMARY_VERSION =
  "query_candidate_planner_real_shadow_evidence_foundation_summary_v1";
const COMPLETED_STATUSES = Object.freeze(["COMPLETED", "COMPLETED_SAFE"]);
const CAPTURE_SOURCES = Object.freeze([
  "API_SHADOW_OBSERVATION",
  "INTERNAL_PREVIEW",
]);
const FORBIDDEN_KEYS = new Set([
  "email",
  "name",
  "userid",
  "user_id",
  "accountid",
  "account_id",
  "googleid",
  "google_id",
  "tenantid",
  "tenant_id",
  "organizationid",
  "organization_id",
  "filename",
  "file_name",
  "originalfilename",
  "original_file_name",
  "querytableskey",
  "query_tables_key",
  "storagekey",
  "storage_key",
  "rawrows",
  "raw_rows",
  "samplevalues",
  "sample_values",
  "jwt",
  "bearer",
  "accesstoken",
  "access_token",
  "secret",
  "rawpayload",
  "raw_payload",
]);

function text(value, maxLength = 240) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function sha256(value) {
  return crypto
    .createHash("sha256")
    .update(typeof value === "string" ? value : JSON.stringify(value))
    .digest("hex");
}

function validIsoTimestamp(value, now = Date.now) {
  const raw = text(value, 80);
  const parsed = Date.parse(raw);
  const nowMs = typeof now === "function" ? Number(now()) : Number(now);
  return (
    Boolean(raw) &&
    Number.isFinite(parsed) &&
    Number.isFinite(nowMs) &&
    parsed <= nowMs + 5 * 60 * 1000
  );
}

function expectedRejectedCases(accuracyDataset = {}) {
  const cases = Array.isArray(accuracyDataset?.cases) ? accuracyDataset.cases : [];
  return cases.filter(
    (item) => item?.labels?.unsupported?.expectedRejected === true,
  );
}

function findForbiddenPaths(value, path = "$", output = []) {
  if (Array.isArray(value)) {
    value.forEach((item, index) =>
      findForbiddenPaths(item, `${path}[${index}]`, output),
    );
    return output;
  }
  if (!value || typeof value !== "object") return output;
  for (const [key, item] of Object.entries(value)) {
    const normalized = String(key).replace(/[^a-zA-Z0-9_]/g, "").toLowerCase();
    if (FORBIDDEN_KEYS.has(normalized)) output.push(`${path}.${key}`);
    findForbiddenPaths(item, `${path}.${key}`, output);
  }
  return output;
}

function buildExpectedRejectionAttestation({
  accuracyDataset,
  sourceCatalog,
  ledger,
  caseId,
  requestFingerprintSha256,
  uploadFingerprintSha256,
  observationStatus,
  observationReason,
  shadowAccepted,
  captureSource = "INTERNAL_PREVIEW",
  observedAt = new Date().toISOString(),
  now = Date.now,
} = {}) {
  const normalizedCaseId = text(caseId, 160);
  const expectedCases = expectedRejectedCases(accuracyDataset);
  const expectedCase = expectedCases.find((item) => item.caseId === normalizedCaseId);
  if (!expectedCase) {
    const error = new Error("case is not an expected-rejection accuracy case");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_CASE_REQUIRED";
    throw error;
  }

  const catalogValidation = validateUploadableSourceCatalog({
    accuracyDataset,
    catalog: sourceCatalog,
    requireComplete: true,
    verifyFiles: true,
    now,
  });
  if (!catalogValidation.valid || !catalogValidation.complete) {
    const error = new Error(catalogValidation.errors.join("; "));
    error.code = "REAL_SHADOW_FOUNDATION_SOURCE_CATALOG_INVALID";
    throw error;
  }

  const ledgerValidation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    sourceCatalog,
    ledger,
    requireComplete: false,
    now,
  });
  if (!ledgerValidation.valid) {
    const error = new Error(ledgerValidation.errors.join("; "));
    error.code = "REAL_SHADOW_FOUNDATION_LEDGER_INVALID";
    throw error;
  }

  const ledgerCase = (ledger?.cases || []).find(
    (item) => item.caseId === normalizedCaseId,
  );
  if (!ledgerCase) {
    const error = new Error("ledger case not found");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_LEDGER_CASE_REQUIRED";
    throw error;
  }

  const requestHash = exactSha256(requestFingerprintSha256);
  const uploadHash = exactSha256(uploadFingerprintSha256);
  if (!requestHash) {
    const error = new Error("request fingerprint must be exactly 64 hex characters");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_REQUEST_FINGERPRINT_INVALID";
    throw error;
  }
  if (!uploadHash) {
    const error = new Error("upload fingerprint must be exactly 64 hex characters");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_UPLOAD_FINGERPRINT_INVALID";
    throw error;
  }
  if (requestHash === uploadHash) {
    const error = new Error("request and upload fingerprints must differ");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_FINGERPRINTS_IDENTICAL";
    throw error;
  }

  const ledgerRequest = exactSha256(ledgerCase.requestFingerprintSha256);
  const ledgerUpload = exactSha256(ledgerCase.uploadFingerprintSha256);
  if (ledgerRequest && ledgerRequest !== requestHash) {
    const error = new Error("request fingerprint does not match the ledger case");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_REQUEST_MISMATCH";
    throw error;
  }
  if (ledgerUpload && ledgerUpload !== uploadHash) {
    const error = new Error("upload fingerprint does not match the ledger case");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_UPLOAD_MISMATCH";
    throw error;
  }

  const normalizedStatus = text(observationStatus, 60).toUpperCase();
  const normalizedReason = text(observationReason, 120).toUpperCase();
  const accepted = Number(shadowAccepted);
  const normalizedCaptureSource = text(captureSource, 80).toUpperCase();

  if (!COMPLETED_STATUSES.includes(normalizedStatus)) {
    const error = new Error("expected rejection must finish in a completed safe state");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_STATUS_INVALID";
    throw error;
  }
  if (!normalizedReason) {
    const error = new Error("expected rejection observation reason is required");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_REASON_REQUIRED";
    throw error;
  }
  if (!Number.isInteger(accepted) || accepted !== 0) {
    const error = new Error("expected rejection must accept zero shadow candidates");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_ACCEPTED_NONZERO";
    throw error;
  }
  if (!CAPTURE_SOURCES.includes(normalizedCaptureSource)) {
    const error = new Error("capture source invalid");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_CAPTURE_SOURCE_INVALID";
    throw error;
  }
  if (!validIsoTimestamp(observedAt, now)) {
    const error = new Error("observedAt invalid");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_TIMESTAMP_INVALID";
    throw error;
  }

  const catalogCase = (sourceCatalog?.cases || []).find(
    (item) => item.caseId === normalizedCaseId,
  );
  const sourceArtifactSha256 = exactSha256(catalogCase?.sourceArtifactSha256);
  if (!sourceArtifactSha256) {
    const error = new Error("source artifact binding required");
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_SOURCE_BINDING_REQUIRED";
    throw error;
  }

  const attestation = Object.freeze({
    version: ATTESTATION_VERSION,
    caseId: normalizedCaseId,
    expectedRejected: true,
    sourceArtifactSha256,
    requestFingerprintSha256: requestHash,
    uploadFingerprintSha256: uploadHash,
    captureSource: normalizedCaptureSource,
    observedAt: new Date(Date.parse(observedAt)).toISOString(),
    observationStatus: normalizedStatus,
    observationReason: normalizedReason,
    shadowAccepted: 0,
    candidatePayloadIncluded: false,
    rawIdentityIncluded: false,
    rawFileContentIncluded: false,
  });

  const forbidden = findForbiddenPaths(attestation);
  if (forbidden.length > 0) {
    const error = new Error(`forbidden attestation field: ${forbidden[0]}`);
    error.code = "REAL_SHADOW_EXPECTED_REJECTION_PRIVACY_VIOLATION";
    throw error;
  }

  return Object.freeze({
    version: FOUNDATION_VERSION,
    attestation,
    attestationSha256: sha256(JSON.stringify(attestation)),
  });
}

function validateExpectedRejectionAttestation({
  accuracyDataset,
  sourceCatalog,
  ledger,
  attestation,
  now = Date.now,
} = {}) {
  const errors = [];
  if (!attestation || typeof attestation !== "object" || Array.isArray(attestation)) {
    errors.push("EXPECTED_REJECTION_ATTESTATION_OBJECT_REQUIRED");
  }
  if (attestation?.version !== ATTESTATION_VERSION) {
    errors.push("EXPECTED_REJECTION_ATTESTATION_VERSION_INVALID");
  }
  const caseId = text(attestation?.caseId, 160);
  const expectedCase = expectedRejectedCases(accuracyDataset).find(
    (item) => item.caseId === caseId,
  );
  if (!expectedCase) errors.push("EXPECTED_REJECTION_CASE_INVALID");

  const ledgerCase = (ledger?.cases || []).find((item) => item.caseId === caseId);
  const catalogCase = (sourceCatalog?.cases || []).find((item) => item.caseId === caseId);
  const requestHash = exactSha256(attestation?.requestFingerprintSha256);
  const uploadHash = exactSha256(attestation?.uploadFingerprintSha256);
  const sourceHash = exactSha256(attestation?.sourceArtifactSha256);

  if (!requestHash) errors.push("EXPECTED_REJECTION_REQUEST_FINGERPRINT_INVALID");
  if (!uploadHash) errors.push("EXPECTED_REJECTION_UPLOAD_FINGERPRINT_INVALID");
  if (requestHash && uploadHash && requestHash === uploadHash) {
    errors.push("EXPECTED_REJECTION_FINGERPRINTS_IDENTICAL");
  }
  if (!sourceHash) errors.push("EXPECTED_REJECTION_SOURCE_HASH_INVALID");
  if (ledgerCase && requestHash !== exactSha256(ledgerCase.requestFingerprintSha256)) {
    errors.push("EXPECTED_REJECTION_REQUEST_LEDGER_MISMATCH");
  }
  if (ledgerCase && uploadHash !== exactSha256(ledgerCase.uploadFingerprintSha256)) {
    errors.push("EXPECTED_REJECTION_UPLOAD_LEDGER_MISMATCH");
  }
  if (catalogCase && sourceHash !== exactSha256(catalogCase.sourceArtifactSha256)) {
    errors.push("EXPECTED_REJECTION_SOURCE_CATALOG_MISMATCH");
  }
  if (attestation?.expectedRejected !== true) {
    errors.push("EXPECTED_REJECTION_FLAG_REQUIRED");
  }
  if (!COMPLETED_STATUSES.includes(text(attestation?.observationStatus, 60).toUpperCase())) {
    errors.push("EXPECTED_REJECTION_STATUS_INVALID");
  }
  if (!text(attestation?.observationReason, 120)) {
    errors.push("EXPECTED_REJECTION_REASON_REQUIRED");
  }
  if (Number(attestation?.shadowAccepted) !== 0) {
    errors.push("EXPECTED_REJECTION_ACCEPTED_MUST_BE_ZERO");
  }
  if (!CAPTURE_SOURCES.includes(text(attestation?.captureSource, 80).toUpperCase())) {
    errors.push("EXPECTED_REJECTION_CAPTURE_SOURCE_INVALID");
  }
  if (!validIsoTimestamp(attestation?.observedAt, now)) {
    errors.push("EXPECTED_REJECTION_OBSERVED_AT_INVALID");
  }
  if (attestation?.candidatePayloadIncluded !== false) {
    errors.push("EXPECTED_REJECTION_CANDIDATE_PAYLOAD_MUST_BE_EXCLUDED");
  }
  if (attestation?.rawIdentityIncluded !== false) {
    errors.push("EXPECTED_REJECTION_RAW_IDENTITY_MUST_BE_EXCLUDED");
  }
  if (attestation?.rawFileContentIncluded !== false) {
    errors.push("EXPECTED_REJECTION_RAW_FILE_CONTENT_MUST_BE_EXCLUDED");
  }
  findForbiddenPaths(attestation).forEach((item) =>
    errors.push(`forbidden field: ${item}`),
  );

  return Object.freeze({
    valid: errors.length === 0,
    errors: Object.freeze([...new Set(errors)]),
    caseId,
    attestationSha256:
      errors.length === 0 ? sha256(JSON.stringify(attestation)) : "",
  });
}

function evaluateRealShadowEvidenceFoundation({
  accuracyDataset,
  sourceCatalog,
  ledger,
  expectedRejectionAttestations = [],
  now = Date.now,
} = {}) {
  const errors = [];

  const catalogValidation = validateUploadableSourceCatalog({
    accuracyDataset,
    catalog: sourceCatalog,
    requireComplete: true,
    verifyFiles: true,
    now,
  });
  if (!catalogValidation.valid || !catalogValidation.complete) {
    errors.push("REAL_SHADOW_FOUNDATION_SOURCE_CATALOG_INCOMPLETE");
    errors.push(...catalogValidation.errors);
  }

  const ledgerValidation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    sourceCatalog,
    ledger,
    requireComplete: true,
    now,
  });
  if (!ledgerValidation.valid || !ledgerValidation.complete) {
    errors.push("REAL_SHADOW_FOUNDATION_LEDGER_INCOMPLETE");
    errors.push(...ledgerValidation.errors);
  }

  const finalization = finalizeRealShadowCaseRegistry({
    accuracyDataset,
    sourceCatalog,
    ledger,
    now,
  });
  if (!finalization.valid) {
    errors.push("REAL_SHADOW_FOUNDATION_REGISTRY_FINALIZATION_BLOCKED");
    errors.push(...(finalization.errors || []));
  }

  const expectedCases = expectedRejectedCases(accuracyDataset);
  const attestations = Array.isArray(expectedRejectionAttestations)
    ? expectedRejectionAttestations
    : [];
  const byCaseId = new Map();
  for (const attestation of attestations) {
    const caseId = text(attestation?.caseId, 160);
    if (!caseId) continue;
    if (byCaseId.has(caseId)) errors.push(`duplicate expected rejection attestation: ${caseId}`);
    byCaseId.set(caseId, attestation);
  }
  for (const expectedCase of expectedCases) {
    const attestation = byCaseId.get(expectedCase.caseId);
    if (!attestation) {
      errors.push(`missing expected rejection attestation: ${expectedCase.caseId}`);
      continue;
    }
    const validation = validateExpectedRejectionAttestation({
      accuracyDataset,
      sourceCatalog,
      ledger,
      attestation,
      now,
    });
    if (!validation.valid) errors.push(...validation.errors);
  }
  for (const caseId of byCaseId.keys()) {
    if (!expectedCases.some((item) => item.caseId === caseId)) {
      errors.push(`unexpected expected rejection attestation: ${caseId}`);
    }
  }

  const dedupedErrors = Object.freeze([...new Set(errors)]);
  const ready = dedupedErrors.length === 0;
  const summary = Object.freeze({
    version: FOUNDATION_SUMMARY_VERSION,
    phase: "15.3-A",
    decision: ready
      ? "REAL_SHADOW_EVIDENCE_FOUNDATION_PASS"
      : "REAL_SHADOW_EVIDENCE_FOUNDATION_BLOCKED",
    readyForPatch15_3_2_C: ready,
    sourceCatalogComplete: Boolean(catalogValidation.complete),
    ledgerComplete: Boolean(ledgerValidation.complete),
    registryFinalized: Boolean(finalization.valid),
    expectedCaseCount: Number(ledgerValidation.expectedCaseCount) || 0,
    completedCaseCount: Number(ledgerValidation.completedCount) || 0,
    expectedRejectionCaseCount: expectedCases.length,
    expectedRejectionEvidenceCount: expectedCases.filter((item) =>
      byCaseId.has(item.caseId),
    ).length,
    sourceCatalogSha256:
      catalogValidation.valid ? sourceCatalogSha256(sourceCatalog) : "",
    ledgerSha256: ledgerValidation.valid ? sha256(JSON.stringify(ledger)) : "",
    registrySha256: finalization.valid ? finalization.registrySha256 : "",
    actualTrafficOnly: ledger?.actualTrafficOnly === true,
    syntheticFingerprintForbidden: ledger?.syntheticFingerprintForbidden === true,
    legacyLedgerAccepted: false,
    rawIdentityIncluded: false,
    collectorEnabledByThisPhase: false,
    internalCanaryEnabledByThisPhase: false,
    productionPromotionAuthorized: false,
    errors: dedupedErrors,
  });

  return Object.freeze({
    version: FOUNDATION_VERSION,
    valid: ready,
    reason: ready
      ? "REAL_SHADOW_EVIDENCE_FOUNDATION_READY_FOR_ENCRYPTION_SECRET"
      : dedupedErrors[0] || "REAL_SHADOW_EVIDENCE_FOUNDATION_BLOCKED",
    errors: dedupedErrors,
    summary,
    registry: finalization.valid ? finalization.registry : null,
    registrySha256: finalization.valid ? finalization.registrySha256 : "",
  });
}

module.exports = Object.freeze({
  FOUNDATION_VERSION,
  ATTESTATION_VERSION,
  FOUNDATION_SUMMARY_VERSION,
  COMPLETED_STATUSES,
  CAPTURE_SOURCES,
  findForbiddenPaths,
  expectedRejectedCases,
  buildExpectedRejectionAttestation,
  validateExpectedRejectionAttestation,
  evaluateRealShadowEvidenceFoundation,
});
