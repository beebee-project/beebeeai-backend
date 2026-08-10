const crypto = require("crypto");
const {
  DRAFT_VERSION,
  buildRealShadowCaseRegistry,
} = require("./queryCandidatePlannerRealShadowPreparation");
const {
  SOURCE_CATALOG_VERSION,
  SHA256_RE,
  exactText,
  sourceCatalogSha256,
  inspectSourceArtifact,
  validateUploadableSourceCatalog,
} = require("./queryCandidatePlannerRealShadowUploadableSourceCatalog");

const FINALIZATION_VERSION =
  "query_candidate_planner_real_shadow_registry_finalization_v2";
const LEDGER_VERSION =
  "query_candidate_planner_real_shadow_fingerprint_ledger_v2";
const LEGACY_LEDGER_VERSION =
  "query_candidate_planner_real_shadow_fingerprint_ledger_v1";
const ACCURACY_DATASET_VERSION =
  "query_candidate_planner_accuracy_evaluation_dataset_v1";
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

function boundedText(value, maxLength = 240) {
  return exactText(value).slice(0, maxLength);
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
    accuracyDataset.cases.length !== 10
  ) {
    const error = new Error(
      "valid ten-case Patch 15.0 accuracy dataset required",
    );
    error.code = "REAL_SHADOW_ACCURACY_DATASET_REQUIRED";
    throw error;
  }
  return accuracyDataset.cases;
}

function canonicalCaseIds(accuracyDataset = {}) {
  return accuracyCases(accuracyDataset).map((item) =>
    boundedText(item.caseId, 160),
  );
}

function validIsoTimestamp(value, now = Date.now) {
  const raw = exactText(value);
  const parsed = Date.parse(raw);
  if (!raw || !Number.isFinite(parsed)) return false;
  const nowMs = typeof now === "function" ? Number(now()) : Number(now);
  return Number.isFinite(nowMs) && parsed <= nowMs + 5 * 60 * 1000;
}

function exactSha256(value) {
  const raw = exactText(value).toLowerCase();
  return SHA256_RE.test(raw) && raw.length === 64 ? raw : "";
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
    const normalized = String(key)
      .replace(/[^a-zA-Z0-9_]/g, "")
      .toLowerCase();
    if (FORBIDDEN_KEYS.has(normalized)) output.push(`${path}.${key}`);
    findForbiddenPaths(item, `${path}.${key}`, output);
  }
  return output;
}

function validatedSourceCatalog(
  accuracyDataset,
  sourceCatalog,
  now = Date.now,
) {
  const validation = validateUploadableSourceCatalog({
    accuracyDataset,
    catalog: sourceCatalog,
    requireComplete: true,
    verifyFiles: true,
    now,
  });
  if (!validation.valid || !validation.complete) {
    const error = new Error(validation.errors.join("; ") || validation.reason);
    error.code = "REAL_SHADOW_UPLOADABLE_SOURCE_CATALOG_REQUIRED";
    throw error;
  }
  if (sourceCatalog.version !== SOURCE_CATALOG_VERSION) {
    const error = new Error("finalized source catalog version required");
    error.code = "REAL_SHADOW_FINALIZED_SOURCE_CATALOG_REQUIRED";
    throw error;
  }
  return sourceCatalog;
}

function buildRealShadowFingerprintLedgerScaffold(
  accuracyDataset,
  sourceCatalog,
  { registryId = "internal_real_shadow_2026_08_v2", now = Date.now } = {},
) {
  const verifiedCatalog = validatedSourceCatalog(
    accuracyDataset,
    sourceCatalog,
    now,
  );
  const byCaseId = new Map(
    verifiedCatalog.cases.map((item) => [item.caseId, item]),
  );
  const cases = accuracyCases(accuracyDataset).map((item) => {
    const source = byCaseId.get(item.caseId);
    return Object.freeze({
      caseId: boundedText(item.caseId, 160),
      scenarioId: `${boundedText(item.caseId, 140)}_internal_01`,
      sourceArtifactSha256: exactSha256(source.sourceArtifactSha256),
      sourceKind: boundedText(source.sourceKind, 40).toUpperCase(),
      requestFingerprintSha256: "",
      uploadFingerprintSha256: "",
      captureSource: "",
      capturedAt: "",
      actualTraffic: true,
      synthetic: false,
      expectedColdCostMicrousd: 0,
      modelId: "semantic_profiler_default",
      operatorNote: "rerun actual source after B.1 catalog finalization",
    });
  });
  return Object.freeze({
    version: LEDGER_VERSION,
    registryId: boundedText(registryId, 160),
    sourceDatasetId: boundedText(accuracyDataset.datasetId, 160),
    sourceCatalogId: boundedText(verifiedCatalog.catalogId, 160),
    sourceCatalogSha256: sourceCatalogSha256(verifiedCatalog),
    legacyCapturesPreserved: false,
    actualTrafficOnly: true,
    syntheticFingerprintForbidden: true,
    sourceArtifactBindingRequired: true,
    rawIdentityIncluded: false,
    cases: Object.freeze(cases),
  });
}

function validateRealShadowFingerprintLedger({
  accuracyDataset,
  sourceCatalog,
  ledger,
  requireComplete = true,
  now = Date.now,
} = {}) {
  const errors = [];
  let expectedCaseIds = [];
  let verifiedCatalog = null;
  try {
    expectedCaseIds = canonicalCaseIds(accuracyDataset);
    verifiedCatalog = validatedSourceCatalog(
      accuracyDataset,
      sourceCatalog,
      now,
    );
  } catch (error) {
    errors.push(error.code || "REAL_SHADOW_UPLOADABLE_SOURCE_CATALOG_REQUIRED");
  }

  if (!ledger || typeof ledger !== "object" || Array.isArray(ledger)) {
    errors.push("FINGERPRINT_LEDGER_OBJECT_REQUIRED");
  }
  if (ledger?.version === LEGACY_LEDGER_VERSION) {
    errors.push("REAL_SHADOW_LEGACY_LEDGER_REJECTED");
  } else if (ledger?.version !== LEDGER_VERSION) {
    errors.push("FINGERPRINT_LEDGER_VERSION_INVALID");
  }
  if (!boundedText(ledger?.registryId, 160)) errors.push("registryId required");
  if (!exactSha256(ledger?.sourceCatalogSha256)) {
    errors.push("sourceCatalogSha256 required");
  }
  if (
    verifiedCatalog &&
    exactSha256(ledger?.sourceCatalogSha256) !==
      sourceCatalogSha256(verifiedCatalog)
  ) {
    errors.push("REAL_SHADOW_SOURCE_CATALOG_BINDING_MISMATCH");
  }
  if (ledger?.legacyCapturesPreserved !== false) {
    errors.push("LEGACY_CAPTURE_PRESERVATION_FORBIDDEN");
  }
  if (ledger?.actualTrafficOnly !== true)
    errors.push("ACTUAL_TRAFFIC_ONLY_REQUIRED");
  if (ledger?.syntheticFingerprintForbidden !== true) {
    errors.push("SYNTHETIC_FINGERPRINT_FORBIDDEN_REQUIRED");
  }
  if (ledger?.sourceArtifactBindingRequired !== true) {
    errors.push("SOURCE_ARTIFACT_BINDING_REQUIRED");
  }
  if (ledger?.rawIdentityIncluded !== false) {
    errors.push("RAW_IDENTITY_MUST_BE_EXCLUDED");
  }
  findForbiddenPaths(ledger).forEach((item) =>
    errors.push(`forbidden field: ${item}`),
  );

  const cases = Array.isArray(ledger?.cases) ? ledger.cases : [];
  const byCaseId = new Map();
  cases.forEach((item, index) => {
    const caseId = boundedText(item?.caseId, 160);
    if (!caseId) errors.push(`cases[${index}].caseId required`);
    else if (byCaseId.has(caseId)) errors.push(`duplicate caseId: ${caseId}`);
    else byCaseId.set(caseId, item);
  });
  expectedCaseIds.forEach((caseId) => {
    if (!byCaseId.has(caseId)) errors.push(`missing caseId: ${caseId}`);
  });
  for (const caseId of byCaseId.keys()) {
    if (!expectedCaseIds.includes(caseId))
      errors.push(`unexpected caseId: ${caseId}`);
  }

  const catalogByCaseId = new Map(
    (verifiedCatalog?.cases || []).map((item) => [item.caseId, item]),
  );
  const requestFingerprints = new Set();
  const uploadFingerprints = new Set();
  let completedCount = 0;

  for (const caseId of expectedCaseIds) {
    const item = byCaseId.get(caseId);
    if (!item) continue;
    const source = catalogByCaseId.get(caseId);
    const requestRaw = exactText(item.requestFingerprintSha256);
    const uploadRaw = exactText(item.uploadFingerprintSha256);
    const requestFingerprintSha256 = exactSha256(requestRaw);
    const uploadFingerprintSha256 = exactSha256(uploadRaw);
    const sourceArtifactSha256 = exactSha256(item.sourceArtifactSha256);
    const captureSource = boundedText(item.captureSource, 80).toUpperCase();
    const hasCapture = Boolean(
      requestRaw || uploadRaw || captureSource || exactText(item.capturedAt),
    );

    if (!sourceArtifactSha256) {
      errors.push(`${caseId}: sourceArtifactSha256 required`);
    } else if (
      source &&
      sourceArtifactSha256 !== exactSha256(source.sourceArtifactSha256)
    ) {
      errors.push(`${caseId}: source artifact binding mismatch`);
    }
    if (requestRaw && requestRaw.length !== 64) {
      errors.push(
        `${caseId}: request fingerprint must be exactly 64 characters`,
      );
    }
    if (uploadRaw && uploadRaw.length !== 64) {
      errors.push(
        `${caseId}: upload fingerprint must be exactly 64 characters`,
      );
    }
    if (requireComplete || hasCapture) {
      if (!requestFingerprintSha256) {
        errors.push(`${caseId}: actual request fingerprint required`);
      }
      if (!uploadFingerprintSha256) {
        errors.push(`${caseId}: actual upload fingerprint required`);
      }
      if (
        requestFingerprintSha256 === uploadFingerprintSha256 &&
        requestFingerprintSha256
      ) {
        errors.push(`${caseId}: request and upload fingerprints must differ`);
      }
      if (!CAPTURE_SOURCES.includes(captureSource)) {
        errors.push(`${caseId}: captureSource invalid`);
      }
      if (!validIsoTimestamp(item.capturedAt, now)) {
        errors.push(`${caseId}: capturedAt invalid`);
      }
    }
    if (item.actualTraffic !== true)
      errors.push(`${caseId}: actualTraffic must be true`);
    if (item.synthetic !== false)
      errors.push(`${caseId}: synthetic must be false`);

    if (requestFingerprintSha256) {
      if (requestFingerprints.has(requestFingerprintSha256)) {
        errors.push(`${caseId}: duplicate request fingerprint`);
      }
      requestFingerprints.add(requestFingerprintSha256);
    }
    if (uploadFingerprintSha256) {
      if (uploadFingerprints.has(uploadFingerprintSha256)) {
        errors.push(`${caseId}: duplicate upload fingerprint`);
      }
      uploadFingerprints.add(uploadFingerprintSha256);
    }
    const complete =
      Boolean(sourceArtifactSha256) &&
      Boolean(requestFingerprintSha256) &&
      Boolean(uploadFingerprintSha256) &&
      requestFingerprintSha256 !== uploadFingerprintSha256 &&
      CAPTURE_SOURCES.includes(captureSource) &&
      validIsoTimestamp(item.capturedAt, now) &&
      item.actualTraffic === true &&
      item.synthetic === false;
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
    legacyLedgerAccepted: false,
    rawIdentityIncluded: false,
  });
}

function upsertRealShadowFingerprintCapture({
  accuracyDataset,
  sourceCatalog,
  ledger,
  sourceFilePath,
  caseId,
  requestFingerprintSha256,
  uploadFingerprintSha256,
  captureSource,
  capturedAt = new Date().toISOString(),
  expectedColdCostMicrousd = 0,
  modelId = "semantic_profiler_default",
  operatorNote = "actual internal request captured",
  now = Date.now,
} = {}) {
  const normalizedCaseId = boundedText(caseId, 160);
  if (!canonicalCaseIds(accuracyDataset).includes(normalizedCaseId)) {
    const error = new Error(`unknown caseId: ${normalizedCaseId}`);
    error.code = "REAL_SHADOW_CASE_ID_NOT_IN_ACCURACY_DATASET";
    throw error;
  }
  const requestRaw = exactText(requestFingerprintSha256);
  const uploadRaw = exactText(uploadFingerprintSha256);
  if (requestRaw.length !== 64 || !SHA256_RE.test(requestRaw)) {
    const error = new Error(
      "request fingerprint must be exactly 64 hex characters",
    );
    error.code = "REAL_SHADOW_REQUEST_FINGERPRINT_INVALID";
    throw error;
  }
  if (uploadRaw.length !== 64 || !SHA256_RE.test(uploadRaw)) {
    const error = new Error(
      "upload fingerprint must be exactly 64 hex characters",
    );
    error.code = "REAL_SHADOW_UPLOAD_FINGERPRINT_INVALID";
    throw error;
  }
  const requestHash = requestRaw.toLowerCase();
  const uploadHash = uploadRaw.toLowerCase();
  if (requestHash === uploadHash) {
    const error = new Error("request and upload fingerprints must differ");
    error.code = "REAL_SHADOW_REQUEST_UPLOAD_FINGERPRINT_IDENTICAL";
    throw error;
  }
  const normalizedSource = boundedText(captureSource, 80).toUpperCase();
  if (!CAPTURE_SOURCES.includes(normalizedSource)) {
    const error = new Error("capture source invalid");
    error.code = "REAL_SHADOW_CAPTURE_SOURCE_INVALID";
    throw error;
  }
  if (!validIsoTimestamp(capturedAt, now)) {
    const error = new Error(
      "capturedAt must be a valid non-future ISO timestamp",
    );
    error.code = "REAL_SHADOW_CAPTURE_TIMESTAMP_INVALID";
    throw error;
  }
  const cost = Number(expectedColdCostMicrousd);
  if (!Number.isInteger(cost) || cost < 0) {
    const error = new Error(
      "expectedColdCostMicrousd must be a non-negative integer",
    );
    error.code = "REAL_SHADOW_EXPECTED_COLD_COST_INVALID";
    throw error;
  }

  const baseValidation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    sourceCatalog,
    ledger,
    requireComplete: false,
    now,
  });
  if (!baseValidation.valid) {
    const error = new Error(baseValidation.errors.join("; "));
    error.code = baseValidation.errors.includes(
      "REAL_SHADOW_LEGACY_LEDGER_REJECTED",
    )
      ? "REAL_SHADOW_LEGACY_LEDGER_REJECTED"
      : "REAL_SHADOW_FINGERPRINT_LEDGER_INVALID";
    throw error;
  }

  const catalogItem = sourceCatalog.cases.find(
    (item) => item.caseId === normalizedCaseId,
  );
  const inspected = inspectSourceArtifact(sourceFilePath);
  if (
    !catalogItem ||
    inspected.sourceArtifactSha256 !==
      exactSha256(catalogItem.sourceArtifactSha256)
  ) {
    const error = new Error("source file does not match the catalog binding");
    error.code = "REAL_SHADOW_CAPTURE_SOURCE_ARTIFACT_MISMATCH";
    throw error;
  }

  for (const item of ledger.cases) {
    if (item.caseId === normalizedCaseId) continue;
    if (exactSha256(item.requestFingerprintSha256) === requestHash) {
      const error = new Error(
        "request fingerprint already assigned to another case",
      );
      error.code = "REAL_SHADOW_REQUEST_FINGERPRINT_DUPLICATE";
      throw error;
    }
    if (exactSha256(item.uploadFingerprintSha256) === uploadHash) {
      const error = new Error(
        "upload fingerprint already assigned to another case",
      );
      error.code = "REAL_SHADOW_UPLOAD_FINGERPRINT_DUPLICATE";
      throw error;
    }
  }

  const updatedCases = ledger.cases.map((item) => {
    if (item.caseId !== normalizedCaseId) return Object.freeze({ ...item });
    return Object.freeze({
      ...item,
      sourceArtifactSha256: inspected.sourceArtifactSha256,
      requestFingerprintSha256: requestHash,
      uploadFingerprintSha256: uploadHash,
      captureSource: normalizedSource,
      capturedAt: new Date(Date.parse(capturedAt)).toISOString(),
      actualTraffic: true,
      synthetic: false,
      expectedColdCostMicrousd: cost,
      modelId: boundedText(
        modelId || item.modelId || "semantic_profiler_default",
        120,
      ),
      operatorNote: boundedText(operatorNote || item.operatorNote, 240),
    });
  });
  const updated = Object.freeze({
    ...ledger,
    version: LEDGER_VERSION,
    sourceCatalogSha256: sourceCatalogSha256(sourceCatalog),
    legacyCapturesPreserved: false,
    cases: Object.freeze(updatedCases),
  });
  const validation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    sourceCatalog,
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
  sourceCatalog,
  ledger,
  now = Date.now,
} = {}) {
  const validation = validateRealShadowFingerprintLedger({
    accuracyDataset,
    sourceCatalog,
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
    registryId: boundedText(ledger.registryId, 160),
    sourceDatasetId: boundedText(ledger.sourceDatasetId, 160),
    actualTrafficOnly: true,
    syntheticFingerprintForbidden: true,
    cases: Object.freeze(
      ledger.cases.map((item) =>
        Object.freeze({
          caseId: boundedText(item.caseId, 160),
          scenarioId: boundedText(item.scenarioId, 160),
          requestFingerprintSha256: exactSha256(item.requestFingerprintSha256),
          uploadFingerprintSha256: exactSha256(item.uploadFingerprintSha256),
          expectedColdCostMicrousd: Number(item.expectedColdCostMicrousd) || 0,
          modelId: boundedText(
            item.modelId || "semantic_profiler_default",
            120,
          ),
          operatorNote: "actual source-bound internal request verified",
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
    sourceCatalogSha256: sourceCatalogSha256(sourceCatalog),
    caseCount: registryResult.caseCount,
    expectedCaseCount: registryResult.expectedCaseCount,
    requestFingerprintCount: registryResult.requestFingerprintCount,
    uploadFingerprintCount: registryResult.uploadFingerprintCount,
    source: "REAL_UPLOADABLE_SOURCE_BOUND_INTERNAL_REQUEST",
    actualTraffic: true,
    synthetic: false,
    rawIdentityIncluded: false,
  });
}

module.exports = Object.freeze({
  FINALIZATION_VERSION,
  LEDGER_VERSION,
  LEGACY_LEDGER_VERSION,
  ACCURACY_DATASET_VERSION,
  SHA256_RE,
  CAPTURE_SOURCES,
  sha256,
  exactSha256,
  findForbiddenPaths,
  buildRealShadowFingerprintLedgerScaffold,
  validateRealShadowFingerprintLedger,
  upsertRealShadowFingerprintCapture,
  finalizeRealShadowCaseRegistry,
});
