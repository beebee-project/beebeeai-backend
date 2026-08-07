"use strict";

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const SOURCE_CATALOG_DRAFT_VERSION =
  "query_candidate_planner_real_shadow_uploadable_source_catalog_draft_v1";
const SOURCE_CATALOG_VERSION =
  "query_candidate_planner_real_shadow_uploadable_source_catalog_v1";
const ACCURACY_DATASET_VERSION =
  "query_candidate_planner_accuracy_evaluation_dataset_v1";
const SHA256_RE = /^[a-f0-9]{64}$/i;
const ALLOWED_SOURCE_KINDS = Object.freeze([
  "REAL_ANONYMIZED",
  "PUBLIC_DATASET",
]);
const ALLOWED_EXTENSIONS = Object.freeze([".csv", ".xlsx", ".xls"]);

function exactText(value) {
  return String(value == null ? "" : value).trim();
}

function boundedText(value, maxLength = 240) {
  return exactText(value).slice(0, maxLength);
}

function sha256Buffer(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

function sha256Json(value) {
  return sha256Buffer(Buffer.from(JSON.stringify(value), "utf8"));
}

function accuracyCases(accuracyDataset = {}) {
  if (
    accuracyDataset?.version !== ACCURACY_DATASET_VERSION ||
    !Array.isArray(accuracyDataset?.cases) ||
    accuracyDataset.cases.length !== 10
  ) {
    const error = new Error("Patch 15.0 ten-case accuracy dataset required");
    error.code = "REAL_SHADOW_TEN_CASE_ACCURACY_DATASET_REQUIRED";
    throw error;
  }
  return accuracyDataset.cases;
}

function validIsoTimestamp(value, now = Date.now) {
  const raw = exactText(value);
  const parsed = Date.parse(raw);
  const nowMs = typeof now === "function" ? Number(now()) : Number(now);
  return (
    raw.length > 0 &&
    Number.isFinite(parsed) &&
    Number.isFinite(nowMs) &&
    parsed <= nowMs + 5 * 60 * 1000
  );
}

function canonicalBindingProjection(catalog = {}) {
  return Object.freeze({
    version: SOURCE_CATALOG_VERSION,
    catalogId: boundedText(catalog.catalogId, 160),
    sourceDatasetId: boundedText(catalog.sourceDatasetId, 160),
    cases: Object.freeze(
      (Array.isArray(catalog.cases) ? catalog.cases : []).map((item) =>
        Object.freeze({
          caseId: boundedText(item.caseId, 160),
          sourceKind: boundedText(item.sourceKind, 40).toUpperCase(),
          sourceArtifactSha256: exactText(
            item.sourceArtifactSha256,
          ).toLowerCase(),
          sourceSizeBytes: Number(item.sourceSizeBytes) || 0,
          sourceExtension: boundedText(item.sourceExtension, 16).toLowerCase(),
          expectedDomain: boundedText(item.expectedDomain, 120),
          expectedIntent: boundedText(item.expectedIntent, 120),
          semanticCompatibilityConfirmed:
            item.semanticCompatibilityConfirmed === true,
          uploadable: item.uploadable === true,
        }),
      ),
    ),
  });
}

function sourceCatalogSha256(catalog = {}) {
  return sha256Json(canonicalBindingProjection(catalog));
}

function inspectSourceArtifact(sourcePath) {
  const resolvedPath = path.resolve(exactText(sourcePath));
  if (!resolvedPath || !fs.existsSync(resolvedPath)) {
    const error = new Error("source artifact does not exist");
    error.code = "REAL_SHADOW_SOURCE_ARTIFACT_NOT_FOUND";
    throw error;
  }
  const stat = fs.statSync(resolvedPath);
  if (!stat.isFile() || stat.size <= 0) {
    const error = new Error("source artifact must be a non-empty file");
    error.code = "REAL_SHADOW_SOURCE_ARTIFACT_FILE_REQUIRED";
    throw error;
  }
  const extension = path.extname(resolvedPath).toLowerCase();
  if (!ALLOWED_EXTENSIONS.includes(extension)) {
    const error = new Error(`source extension not allowed: ${extension}`);
    error.code = "REAL_SHADOW_SOURCE_EXTENSION_NOT_ALLOWED";
    throw error;
  }
  const content = fs.readFileSync(resolvedPath);
  return Object.freeze({
    sourcePath: resolvedPath,
    sourceArtifactSha256: sha256Buffer(content),
    sourceSizeBytes: stat.size,
    sourceExtension: extension,
  });
}

function buildUploadableSourceCatalogScaffold(
  accuracyDataset,
  { catalogId = "internal_real_shadow_uploadable_sources_2026_08_v1" } = {},
) {
  const cases = accuracyCases(accuracyDataset).map((item) =>
    Object.freeze({
      caseId: boundedText(item.caseId, 160),
      scenario: boundedText(item.scenario, 500),
      expectedDomain: boundedText(item.labels?.domain?.expected, 120),
      expectedIntent: boundedText(item.labels?.intent?.expected, 120),
      sourceKind: "",
      sourcePath: "",
      sourceArtifactSha256: "",
      sourceSizeBytes: 0,
      sourceExtension: "",
      uploadable: false,
      semanticCompatibilityConfirmed: false,
      verifiedAt: "",
      operatorNote: "bind only a real anonymized or public uploadable file",
    }),
  );
  return Object.freeze({
    version: SOURCE_CATALOG_DRAFT_VERSION,
    catalogId: boundedText(catalogId, 160),
    sourceDatasetId: boundedText(accuracyDataset.datasetId, 160),
    sourcePolicy: Object.freeze({
      allowedKinds: ALLOWED_SOURCE_KINDS,
      syntheticForbidden: true,
      generatedFixtureForbidden: true,
      uploadableFileRequired: true,
      semanticCompatibilityRequired: true,
    }),
    privatePathsIncluded: true,
    rawWorkbookDataIncluded: false,
    cases: Object.freeze(cases),
  });
}

function validateUploadableSourceCatalog({
  accuracyDataset,
  catalog,
  requireComplete = true,
  verifyFiles = true,
  now = Date.now,
} = {}) {
  const errors = [];
  let expectedCases = [];
  try {
    expectedCases = accuracyCases(accuracyDataset);
  } catch (error) {
    errors.push(error.code || "REAL_SHADOW_TEN_CASE_ACCURACY_DATASET_REQUIRED");
  }
  if (!catalog || typeof catalog !== "object" || Array.isArray(catalog)) {
    errors.push("SOURCE_CATALOG_OBJECT_REQUIRED");
  }
  if (
    catalog?.version !== SOURCE_CATALOG_DRAFT_VERSION &&
    catalog?.version !== SOURCE_CATALOG_VERSION
  ) {
    errors.push("SOURCE_CATALOG_VERSION_INVALID");
  }
  if (!boundedText(catalog?.catalogId, 160)) errors.push("catalogId required");
  if (catalog?.sourcePolicy?.syntheticForbidden !== true) {
    errors.push("SYNTHETIC_SOURCE_MUST_BE_FORBIDDEN");
  }
  if (catalog?.sourcePolicy?.generatedFixtureForbidden !== true) {
    errors.push("GENERATED_FIXTURE_MUST_BE_FORBIDDEN");
  }
  const items = Array.isArray(catalog?.cases) ? catalog.cases : [];
  const byCaseId = new Map();
  for (const [index, item] of items.entries()) {
    const caseId = boundedText(item?.caseId, 160);
    if (!caseId) {
      errors.push(`cases[${index}].caseId required`);
      continue;
    }
    if (byCaseId.has(caseId)) errors.push(`duplicate caseId: ${caseId}`);
    byCaseId.set(caseId, item);
  }
  const expectedIds = expectedCases.map((item) => boundedText(item.caseId, 160));
  expectedIds.forEach((caseId) => {
    if (!byCaseId.has(caseId)) errors.push(`missing caseId: ${caseId}`);
  });
  for (const caseId of byCaseId.keys()) {
    if (!expectedIds.includes(caseId)) errors.push(`unexpected caseId: ${caseId}`);
  }

  const artifactHashes = new Set();
  const sourcePaths = new Set();
  let completedCount = 0;
  for (const expected of expectedCases) {
    const caseId = boundedText(expected.caseId, 160);
    const item = byCaseId.get(caseId);
    if (!item) continue;
    const sourceKind = boundedText(item.sourceKind, 40).toUpperCase();
    const sourcePath = exactText(item.sourcePath);
    const sourceArtifactSha256 = exactText(
      item.sourceArtifactSha256,
    ).toLowerCase();
    const sourceSizeBytes = Number(item.sourceSizeBytes);
    const sourceExtension = boundedText(item.sourceExtension, 16).toLowerCase();
    const hasAnyBinding = Boolean(
      sourceKind || sourcePath || sourceArtifactSha256 || sourceSizeBytes,
    );
    if (requireComplete || hasAnyBinding) {
      if (!ALLOWED_SOURCE_KINDS.includes(sourceKind)) {
        errors.push(`${caseId}: sourceKind must be REAL_ANONYMIZED or PUBLIC_DATASET`);
      }
      if (!sourcePath) errors.push(`${caseId}: sourcePath required`);
      if (!SHA256_RE.test(sourceArtifactSha256)) {
        errors.push(`${caseId}: sourceArtifactSha256 invalid`);
      }
      if (!Number.isInteger(sourceSizeBytes) || sourceSizeBytes <= 0) {
        errors.push(`${caseId}: sourceSizeBytes invalid`);
      }
      if (!ALLOWED_EXTENSIONS.includes(sourceExtension)) {
        errors.push(`${caseId}: sourceExtension invalid`);
      }
      if (item.uploadable !== true) errors.push(`${caseId}: uploadable must be true`);
      if (item.semanticCompatibilityConfirmed !== true) {
        errors.push(`${caseId}: semantic compatibility confirmation required`);
      }
      if (!validIsoTimestamp(item.verifiedAt, now)) {
        errors.push(`${caseId}: verifiedAt invalid`);
      }
    }
    if (SHA256_RE.test(sourceArtifactSha256)) {
      if (artifactHashes.has(sourceArtifactSha256)) {
        errors.push(`${caseId}: duplicate source artifact`);
      }
      artifactHashes.add(sourceArtifactSha256);
    }
    if (sourcePath) {
      const normalizedPath = path.resolve(sourcePath).toLowerCase();
      if (sourcePaths.has(normalizedPath)) errors.push(`${caseId}: duplicate source path`);
      sourcePaths.add(normalizedPath);
    }
    let fileVerified = !verifyFiles;
    if (verifyFiles && sourcePath) {
      try {
        const inspected = inspectSourceArtifact(sourcePath);
        fileVerified =
          inspected.sourceArtifactSha256 === sourceArtifactSha256 &&
          inspected.sourceSizeBytes === sourceSizeBytes &&
          inspected.sourceExtension === sourceExtension;
        if (!fileVerified) errors.push(`${caseId}: source artifact metadata mismatch`);
      } catch (error) {
        errors.push(`${caseId}: ${error.code || error.message}`);
      }
    }
    const complete =
      ALLOWED_SOURCE_KINDS.includes(sourceKind) &&
      Boolean(sourcePath) &&
      SHA256_RE.test(sourceArtifactSha256) &&
      Number.isInteger(sourceSizeBytes) &&
      sourceSizeBytes > 0 &&
      ALLOWED_EXTENSIONS.includes(sourceExtension) &&
      item.uploadable === true &&
      item.semanticCompatibilityConfirmed === true &&
      validIsoTimestamp(item.verifiedAt, now) &&
      fileVerified;
    if (complete) completedCount += 1;
  }

  const uniqueErrors = Object.freeze([...new Set(errors)]);
  return Object.freeze({
    version: SOURCE_CATALOG_VERSION,
    valid: uniqueErrors.length === 0,
    complete: completedCount === expectedCases.length,
    reason:
      uniqueErrors.length > 0
        ? uniqueErrors[0]
        : completedCount === expectedCases.length
          ? "REAL_SHADOW_UPLOADABLE_SOURCE_CATALOG_COMPLETE"
          : "REAL_SHADOW_UPLOADABLE_SOURCE_CATALOG_INCOMPLETE",
    errors: uniqueErrors,
    expectedCaseCount: expectedCases.length,
    completedCount,
    remainingCount: Math.max(0, expectedCases.length - completedCount),
    distinctArtifactCount: artifactHashes.size,
    syntheticSourceAllowed: false,
  });
}

function bindUploadableSource({
  accuracyDataset,
  catalog,
  caseId,
  sourcePath,
  sourceKind,
  semanticCompatibilityConfirmed,
  verifiedAt = new Date().toISOString(),
  operatorNote = "real uploadable source verified",
  now = Date.now,
} = {}) {
  const expectedCases = accuracyCases(accuracyDataset);
  const normalizedCaseId = boundedText(caseId, 160);
  const expected = expectedCases.find((item) => item.caseId === normalizedCaseId);
  if (!expected) {
    const error = new Error(`unknown caseId: ${normalizedCaseId}`);
    error.code = "REAL_SHADOW_SOURCE_CASE_ID_UNKNOWN";
    throw error;
  }
  const normalizedKind = boundedText(sourceKind, 40).toUpperCase();
  if (!ALLOWED_SOURCE_KINDS.includes(normalizedKind)) {
    const error = new Error("only REAL_ANONYMIZED or PUBLIC_DATASET is allowed");
    error.code = "REAL_SHADOW_SOURCE_KIND_INVALID";
    throw error;
  }
  if (semanticCompatibilityConfirmed !== true) {
    const error = new Error("semantic compatibility confirmation required");
    error.code = "REAL_SHADOW_SOURCE_SEMANTIC_CONFIRMATION_REQUIRED";
    throw error;
  }
  if (!validIsoTimestamp(verifiedAt, now)) {
    const error = new Error("verifiedAt invalid");
    error.code = "REAL_SHADOW_SOURCE_VERIFIED_AT_INVALID";
    throw error;
  }
  const inspected = inspectSourceArtifact(sourcePath);
  const baseValidation = validateUploadableSourceCatalog({
    accuracyDataset,
    catalog,
    requireComplete: false,
    verifyFiles: false,
    now,
  });
  if (!baseValidation.valid) {
    const error = new Error(baseValidation.errors.join("; "));
    error.code = "REAL_SHADOW_SOURCE_CATALOG_INVALID";
    throw error;
  }
  for (const item of catalog.cases) {
    if (item.caseId === normalizedCaseId) continue;
    if (
      exactText(item.sourceArtifactSha256).toLowerCase() ===
      inspected.sourceArtifactSha256
    ) {
      const error = new Error("source artifact already assigned to another case");
      error.code = "REAL_SHADOW_SOURCE_ARTIFACT_DUPLICATE";
      throw error;
    }
    if (
      item.sourcePath &&
      path.resolve(item.sourcePath).toLowerCase() ===
        inspected.sourcePath.toLowerCase()
    ) {
      const error = new Error("source path already assigned to another case");
      error.code = "REAL_SHADOW_SOURCE_PATH_DUPLICATE";
      throw error;
    }
  }
  const updatedCases = catalog.cases.map((item) => {
    if (item.caseId !== normalizedCaseId) return Object.freeze({ ...item });
    return Object.freeze({
      caseId: normalizedCaseId,
      scenario: boundedText(expected.scenario, 500),
      expectedDomain: boundedText(expected.labels?.domain?.expected, 120),
      expectedIntent: boundedText(expected.labels?.intent?.expected, 120),
      sourceKind: normalizedKind,
      sourcePath: inspected.sourcePath,
      sourceArtifactSha256: inspected.sourceArtifactSha256,
      sourceSizeBytes: inspected.sourceSizeBytes,
      sourceExtension: inspected.sourceExtension,
      uploadable: true,
      semanticCompatibilityConfirmed: true,
      verifiedAt: new Date(Date.parse(verifiedAt)).toISOString(),
      operatorNote: boundedText(operatorNote, 240),
    });
  });
  const updated = Object.freeze({
    ...catalog,
    version: SOURCE_CATALOG_DRAFT_VERSION,
    cases: Object.freeze(updatedCases),
  });
  const validation = validateUploadableSourceCatalog({
    accuracyDataset,
    catalog: updated,
    requireComplete: false,
    verifyFiles: true,
    now,
  });
  if (!validation.valid) {
    const error = new Error(validation.errors.join("; "));
    error.code = "REAL_SHADOW_SOURCE_BINDING_REJECTED";
    throw error;
  }
  return Object.freeze({
    catalog: updated,
    recordedCaseId: normalizedCaseId,
    completedCount: validation.completedCount,
    remainingCount: validation.remainingCount,
    complete: validation.complete,
    sourceArtifactSha256: inspected.sourceArtifactSha256,
    sourceSizeBytes: inspected.sourceSizeBytes,
  });
}

function finalizeUploadableSourceCatalog({
  accuracyDataset,
  catalog,
  now = Date.now,
} = {}) {
  const validation = validateUploadableSourceCatalog({
    accuracyDataset,
    catalog,
    requireComplete: true,
    verifyFiles: true,
    now,
  });
  if (!validation.valid || !validation.complete) {
    return Object.freeze({
      valid: false,
      reason: validation.reason,
      errors: validation.errors,
      privateCatalog: null,
      publicCatalog: null,
      sourceCatalogSha256: "",
      completedCount: validation.completedCount,
      expectedCaseCount: validation.expectedCaseCount,
    });
  }
  const privateCatalog = Object.freeze({
    ...catalog,
    version: SOURCE_CATALOG_VERSION,
  });
  const publicCatalog = canonicalBindingProjection(privateCatalog);
  const hash = sourceCatalogSha256(privateCatalog);
  return Object.freeze({
    valid: true,
    reason: "REAL_SHADOW_UPLOADABLE_SOURCE_CATALOG_FINALIZED",
    errors: Object.freeze([]),
    privateCatalog,
    publicCatalog,
    sourceCatalogSha256: hash,
    completedCount: validation.completedCount,
    expectedCaseCount: validation.expectedCaseCount,
    synthetic: false,
    actualUploadableSources: true,
  });
}

module.exports = Object.freeze({
  SOURCE_CATALOG_DRAFT_VERSION,
  SOURCE_CATALOG_VERSION,
  ACCURACY_DATASET_VERSION,
  SHA256_RE,
  ALLOWED_SOURCE_KINDS,
  ALLOWED_EXTENSIONS,
  exactText,
  sha256Buffer,
  sourceCatalogSha256,
  inspectSourceArtifact,
  buildUploadableSourceCatalogScaffold,
  validateUploadableSourceCatalog,
  bindUploadableSource,
  finalizeUploadableSourceCatalog,
});
