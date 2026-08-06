"use strict";

const crypto = require("crypto");
const {
  DRAFT_VERSION,
  buildRealShadowCaseRegistry,
} = require("./queryCandidatePlannerRealShadowPreparation");

const FINALIZATION_VERSION =
  "query_candidate_planner_real_shadow_registry_finalization_v1";
const LEDGER_VERSION =
  "query_candidate_planner_real_shadow_fingerprint_ledger_v1";
const ACCURACY_DATASET_VERSION =
  "query_candidate_planner_accuracy_evaluation_dataset_v1";
const SHA256_RE = /^[a-f0-9]{64}$/i;
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
  "_id",
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

function unique(values = []) {
  return [...new Set(values)];
}

function accuracyCases(accuracyDataset = {}) {
  if (
    accuracyDataset?.version !== ACCURACY_DATASET_VERSION ||
    !Array.isArray(accuracyDataset?.cases) ||
    accuracyDataset.cases.length === 0
  ) {
    const error = new Error("valid Patch 15.0 accuracy dataset is required");
    error.code = "REAL_SHADOW_ACCURACY_DATASET_REQUIRED";
    throw error;
  }
  return accuracyDataset.cases;
}

function canonicalCaseIds(accuracyDataset = {}) {
  return accuracyCases(accuracyDataset).map((item) => text(item.caseId, 160));
}

function validIsoTimestamp(value, now = Date.now) {
  const raw = text(value, 80);
  const parsed = Date.parse(raw);
  if (!raw || !Number.isFinite(parsed)) return false;
  const nowMs = typeof now === "function" ? Number(now()) : Number(now);
  if (!Number.isFinite(nowMs)) return false;
  return parsed <= nowMs + 5 * 60 * 1000;
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

function buildRealShadowFingerprintLedgerScaffold(
  accuracyDataset,
  {
    registryId = "internal_real_shadow_2026_08_v1",
  } = {},
) {
  const cases = accuracyCases(accuracyDataset).map((item) =>
    Object.freeze({
      caseId: text(item.caseId, 160),
      scenarioId: `${text(item.caseId, 140)}_internal_01`,
      requestFingerprintSha256: "",
      uploadFingerprintSha256: "",
      captureSource: "",
      capturedAt: "",
      actualTraffic: true,
      synthetic: false,
      expectedColdCostMicrousd: 0,
      modelId: "semantic_profiler_default",
      operatorNote: "record only from an actual internal request",
    }),
  );
  return Object.freeze({
    version: LEDGER_VERSION,
    registryId: text(registryId, 160),
    sourceDatasetId: text(accuracyDataset.datasetId, 160),
    actualTrafficOnly: true,
    syntheticFingerprintForbidden: true,
    rawIdentityIncluded: false,
    cases: Object.freeze(cases),
  });
}

function validateRealShadowFingerprintLedger({
  accuracyDataset,
  ledger,
  requireComplete = true,
  now = Date.now,
} = {}) {
  const errors = [];
  let expectedCaseIds = [];
  try {
    expectedCaseIds = canonicalCaseIds(accuracyDataset);
  } catch (error) {
    errors.push(error.code || "REAL_SHADOW_ACCURACY_DATASET_REQUIRED");
  }

  if (!ledger || typeof ledger !== "object" || Array.isArray(ledger)) {
    errors.push("FINGERPRINT_LEDGER_OBJECT_REQUIRED");
  }
  if (ledger?.version !== LEDGER_VERSION) {
    errors.push("FINGERPRINT_LEDGER_VERSION_INVALID");
  }
  if (!text(ledger?.registryId, 160)) errors.push("registryId required");
  if (ledger?.actualTrafficOnly !== true) {
    errors.push("ACTUAL_TRAFFIC_ONLY_REQUIRED");
  }
  if (ledger?.syntheticFingerprintForbidden !== true) {
    errors.push("SYNTHETIC_FINGERPRINT_FORBIDDEN_REQUIRED");
  }
  if (ledger?.rawIdentityIncluded !== false) {
    errors.push("RAW_IDENTITY_MUST_BE_EXCLUDED");
  }

  const forbiddenPaths = findForbiddenPaths(ledger);
  forbiddenPaths.forEach((path) => errors.push(`forbidden field: ${path}`));

  const cases = Array.isArray(ledger?.cases) ? ledger.cases : [];
  if (cases.length === 0) errors.push("FINGERPRINT_LEDGER_CASES_REQUIRED");

  const byCaseId = new Map();
  for (const [index, item] of cases.entries()) {
    const caseId = text(item?.caseId, 160);
    if (!caseId) {
      errors.push(`cases[${index}].caseId required`);
      continue;
    }
    if (byCaseId.has(caseId)) errors.push(`duplicate caseId: ${caseId}`);
    byCaseId.set(caseId, item);
  }

  for (const caseId of expectedCaseIds) {
    if (!byCaseId.has(caseId)) errors.push(`missing caseId: ${caseId}`);
  }
  for (const caseId of byCaseId.keys()) {
    if (!expectedCaseIds.includes(caseId)) errors.push(`unexpected caseId: ${caseId}`);
  }

  const requestFingerprints = new Set();
  const uploadFingerprints = new Set();
  let completedCount = 0;

  for (const caseId of expectedCaseIds) {
    const item = byCaseId.get(caseId);
    if (!item) continue;
    const requestFingerprintSha256 = text(
      item.requestFingerprintSha256,
      64,
    ).toLowerCase();
    const uploadFingerprintSha256 = text(
      item.uploadFingerprintSha256,
      64,
    ).toLowerCase();
    const captureSource = text(item.captureSource, 80).toUpperCase();
    const complete =
      SHA256_RE.test(requestFingerprintSha256) &&
      SHA256_RE.test(uploadFingerprintSha256) &&
      CAPTURE_SOURCES.includes(captureSource) &&
      validIsoTimestamp(item.capturedAt, now) &&
      item.actualTraffic === true &&
      item.synthetic === false;

    if (requireComplete || requestFingerprintSha256) {
      if (!SHA256_RE.test(requestFingerprintSha256)) {
        errors.push(`${caseId}: actual request fingerprint required`);
      }
    }
    if (requireComplete || uploadFingerprintSha256) {
      if (!SHA256_RE.test(uploadFingerprintSha256)) {
        errors.push(`${caseId}: actual upload fingerprint required`);
      }
    }
    if (requireComplete || captureSource) {
      if (!CAPTURE_SOURCES.includes(captureSource)) {
        errors.push(`${caseId}: captureSource invalid`);
      }
    }
    if (requireComplete || text(item.capturedAt, 80)) {
      if (!validIsoTimestamp(item.capturedAt, now)) {
        errors.push(`${caseId}: capturedAt invalid`);
      }
    }
    if (item.actualTraffic !== true) {
      errors.push(`${caseId}: actualTraffic must be true`);
    }
    if (item.synthetic !== false) {
      errors.push(`${caseId}: synthetic must be false`);
    }

    if (SHA256_RE.test(requestFingerprintSha256)) {
      if (requestFingerprints.has(requestFingerprintSha256)) {
        errors.push(`${caseId}: duplicate request fingerprint`);
      }
      requestFingerprints.add(requestFingerprintSha256);
    }
    if (SHA256_RE.test(uploadFingerprintSha256)) {
      if (uploadFingerprints.has(uploadFingerprintSha256)) {
        errors.push(`${caseId}: duplicate upload fingerprint`);
      }
      uploadFingerprints.add(uploadFingerprintSha256);
    }
    if (complete) completedCount += 1;
  }

  const dedupedErrors = Object.freeze(unique(errors));
  return Object.freeze({
    version: FINALIZATION_VERSION,
    valid: dedupedErrors.length === 0,
    complete: completedCount === expectedCaseIds.length,
    reason:
      dedupedErrors.length > 0
        ? dedupedErrors[0]
        : completedCount === expectedCaseIds.length
          ? "REAL_SHADOW_FINGERPRINT_LEDGER_COMPLETE"
          : "REAL_SHADOW_FINGERPRINT_LEDGER_INCOMPLETE",
    errors: dedupedErrors,
    expectedCaseCount: expectedCaseIds.length,
    completedCount,
    remainingCount: Math.max(0, expectedCaseIds.length - completedCount),
    requestFingerprintCount: requestFingerprints.size,
    uploadFingerprintCount: uploadFingerprints.size,
    rawIdentityIncluded: false,
    syntheticFingerprintAllowed: false,
  });
}

function upsertRealShadowFingerprintCapture({
  accuracyDataset,
  ledger,
  caseId,
  requestFingerprintSha256,
  uploadFingerprintSha256,
  captureSource,
  capturedAt = new Date().toISOString(),
  expectedColdCostMicrousd = 0,
  modelId = "semantic_profiler_default",
  operatorNote = "",
  now = Date.now,
} = {}) {
  const expectedCaseIds = canonicalCaseIds(accuracyDataset);
  const normalizedCaseId = text(caseId, 160);
  if (!expectedCaseIds.includes(normalizedCaseId)) {
    const error = new Error(`unknown caseId: ${normalizedCaseId}`);
    error.code = "REAL_SHADOW_CASE_ID_NOT_IN_ACCURACY_DATASET";
    throw error;
  }

  const requestHash = text(requestFingerprintSha256, 64).toLowerCase();
  const uploadHash = text(uploadFingerprintSha256, 64).toLowerCase();
  const normalizedSource = text(captureSource, 80).toUpperCase();
  if (!SHA256_RE.test(requestHash)) {
    const error = new Error("actual request fingerprint must be 64 hex characters");
    error.code = "REAL_SHADOW_REQUEST_FINGERPRINT_INVALID";
    throw error;
  }
  if (!SHA256_RE.test(uploadHash)) {
    const error = new Error("actual upload fingerprint must be 64 hex characters");
    error.code = "REAL_SHADOW_UPLOAD_FINGERPRINT_INVALID";
    throw error;
  }
  if (!CAPTURE_SOURCES.includes(normalizedSource)) {
    const error = new Error("capture source invalid");
    error.code = "REAL_SHADOW_CAPTURE_SOURCE_INVALID";
    throw error;
  }
  if (!validIsoTimestamp(capturedAt, now)) {
    const error = new Error("capturedAt must be a valid non-future ISO timestamp");
    error.code = "REAL_SHADOW_CAPTURE_TIMESTAMP_INVALID";
    throw error;
  }
  const cost = Number(expectedColdCostMicrousd);
  if (!Number.isInteger(cost) || cost < 0) {
    const error = new Error("expectedColdCostMicrousd must be a non-negative integer");
    error.code = "REAL_SHADOW_EXPECTED_COLD_COST_INVALID";
    throw error;
  }

  const baseValidation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    ledger,
    requireComplete: false,
    now,
  });
  if (!baseValidation.valid) {
    const error = new Error(baseValidation.errors.join("; "));
    error.code = "REAL_SHADOW_FINGERPRINT_LEDGER_INVALID";
    throw error;
  }

  for (const item of ledger.cases) {
    if (item.caseId === normalizedCaseId) continue;
    if (text(item.requestFingerprintSha256, 64).toLowerCase() === requestHash) {
      const error = new Error("request fingerprint already assigned to another case");
      error.code = "REAL_SHADOW_REQUEST_FINGERPRINT_DUPLICATE";
      throw error;
    }
    if (text(item.uploadFingerprintSha256, 64).toLowerCase() === uploadHash) {
      const error = new Error("upload fingerprint already assigned to another case");
      error.code = "REAL_SHADOW_UPLOAD_FINGERPRINT_DUPLICATE";
      throw error;
    }
  }

  const updatedCases = ledger.cases.map((item) => {
    if (item.caseId !== normalizedCaseId) return Object.freeze({ ...item });
    return Object.freeze({
      caseId: normalizedCaseId,
      scenarioId: text(
        item.scenarioId || `${normalizedCaseId}_internal_01`,
        160,
      ),
      requestFingerprintSha256: requestHash,
      uploadFingerprintSha256: uploadHash,
      captureSource: normalizedSource,
      capturedAt: new Date(Date.parse(capturedAt)).toISOString(),
      actualTraffic: true,
      synthetic: false,
      expectedColdCostMicrousd: cost,
      modelId: text(modelId || item.modelId || "semantic_profiler_default", 120),
      operatorNote: text(operatorNote || item.operatorNote, 240),
    });
  });

  const updated = Object.freeze({
    version: LEDGER_VERSION,
    registryId: text(ledger.registryId, 160),
    sourceDatasetId: text(ledger.sourceDatasetId, 160),
    actualTrafficOnly: true,
    syntheticFingerprintForbidden: true,
    rawIdentityIncluded: false,
    cases: Object.freeze(updatedCases),
  });
  const validation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    ledger: updated,
    requireComplete: false,
    now,
  });
  if (!validation.valid) {
    const error = new Error(validation.errors.join("; "));
    error.code = "REAL_SHADOW_FINGERPRINT_CAPTURE_REJECTED";
    throw error;
  }
  return Object.freeze({
    version: FINALIZATION_VERSION,
    ledger: updated,
    ledgerSha256: sha256(JSON.stringify(updated)),
    recordedCaseId: normalizedCaseId,
    completedCount: validation.completedCount,
    remainingCount: validation.remainingCount,
    complete: validation.complete,
    rawIdentityIncluded: false,
  });
}

function finalizeRealShadowCaseRegistry({
  accuracyDataset,
  ledger,
  now = Date.now,
} = {}) {
  const validation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    ledger,
    requireComplete: true,
    now,
  });
  if (!validation.valid || !validation.complete) {
    return Object.freeze({
      version: FINALIZATION_VERSION,
      valid: false,
      reason: validation.reason,
      errors: validation.errors,
      registry: null,
      registrySha256: "",
      caseCount: validation.completedCount,
      expectedCaseCount: validation.expectedCaseCount,
      rawIdentityIncluded: false,
    });
  }

  const draft = Object.freeze({
    version: DRAFT_VERSION,
    registryId: text(ledger.registryId, 160),
    sourceDatasetId: text(ledger.sourceDatasetId, 160),
    actualTrafficOnly: true,
    syntheticFingerprintForbidden: true,
    cases: Object.freeze(
      ledger.cases.map((item) =>
        Object.freeze({
          caseId: text(item.caseId, 160),
          scenarioId: text(item.scenarioId, 160),
          requestFingerprintSha256: text(
            item.requestFingerprintSha256,
            64,
          ).toLowerCase(),
          uploadFingerprintSha256: text(
            item.uploadFingerprintSha256,
            64,
          ).toLowerCase(),
          expectedColdCostMicrousd: Number(item.expectedColdCostMicrousd) || 0,
          modelId: text(item.modelId || "semantic_profiler_default", 120),
          operatorNote: "actual internal request fingerprint verified",
        }),
      ),
    ),
  });
  const registryResult = buildRealShadowCaseRegistry({
    accuracyDataset,
    draft,
    requireUploadFingerprint: true,
  });
  if (!registryResult.valid) {
    return Object.freeze({
      version: FINALIZATION_VERSION,
      valid: false,
      reason: registryResult.reason,
      errors: registryResult.errors,
      registry: null,
      registrySha256: "",
      caseCount: registryResult.caseCount,
      expectedCaseCount: registryResult.expectedCaseCount,
      rawIdentityIncluded: false,
    });
  }
  return Object.freeze({
    version: FINALIZATION_VERSION,
    valid: true,
    reason: "REAL_SHADOW_CASE_REGISTRY_FINALIZED",
    errors: Object.freeze([]),
    registry: registryResult.registry,
    registrySha256: registryResult.registrySha256,
    ledgerSha256: sha256(JSON.stringify(ledger)),
    caseCount: registryResult.caseCount,
    expectedCaseCount: registryResult.expectedCaseCount,
    requestFingerprintCount: registryResult.requestFingerprintCount,
    uploadFingerprintCount: registryResult.uploadFingerprintCount,
    source: "REAL_INTERNAL_REQUEST",
    actualTraffic: true,
    synthetic: false,
    rawIdentityIncluded: false,
  });
}

module.exports = Object.freeze({
  FINALIZATION_VERSION,
  LEDGER_VERSION,
  ACCURACY_DATASET_VERSION,
  SHA256_RE,
  CAPTURE_SOURCES,
  sha256,
  findForbiddenPaths,
  buildRealShadowFingerprintLedgerScaffold,
  validateRealShadowFingerprintLedger,
  upsertRealShadowFingerprintCapture,
  finalizeRealShadowCaseRegistry,
});
