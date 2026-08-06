"use strict";

const crypto = require("crypto");
const {
  parseRegistry,
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
} = require("./queryCandidatePlannerRealShadowEvidenceConfig");

const PREPARATION_VERSION =
  "query_candidate_planner_real_shadow_preparation_v1";
const DRAFT_VERSION =
  "query_candidate_planner_real_shadow_case_registry_draft_v1";
const REGISTRY_VERSION =
  "query_candidate_planner_real_shadow_case_registry_v1";
const SECRET_BYTES = 48;
const SHA256_RE = /^[a-f0-9]{64}$/i;
const BASE64URL_64_RE = /^[A-Za-z0-9_-]{64}$/;

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
    accuracyDataset?.version !==
      "query_candidate_planner_accuracy_evaluation_dataset_v1" ||
    !Array.isArray(accuracyDataset?.cases) ||
    accuracyDataset.cases.length === 0
  ) {
    const error = new Error("valid Patch 15.0 accuracy dataset is required");
    error.code = "REAL_SHADOW_ACCURACY_DATASET_REQUIRED";
    throw error;
  }
  return accuracyDataset.cases;
}

function generateRealShadowEvidenceSecret({ randomBytes = crypto.randomBytes } = {}) {
  const raw = randomBytes(SECRET_BYTES);
  if (!Buffer.isBuffer(raw) || raw.length !== SECRET_BYTES) {
    const error = new Error("secret generator must return exactly 48 random bytes");
    error.code = "REAL_SHADOW_SECRET_RANDOM_SOURCE_INVALID";
    throw error;
  }
  const secret = raw.toString("base64url");
  if (!BASE64URL_64_RE.test(secret)) {
    const error = new Error("generated secret format invalid");
    error.code = "REAL_SHADOW_SECRET_FORMAT_INVALID";
    throw error;
  }
  return Object.freeze({
    version: PREPARATION_VERSION,
    secret,
    secretSha256: sha256(secret),
    entropyBytes: SECRET_BYTES,
    format: "BASE64URL_64",
    reusableWithOtherSecrets: false,
  });
}

function buildRealShadowCaseRegistryScaffold(
  accuracyDataset,
  {
    registryId = "internal_real_shadow_2026_08_v1",
    modelId = "semantic_profiler_default",
  } = {},
) {
  const cases = accuracyCases(accuracyDataset).map((item) =>
    Object.freeze({
      caseId: text(item.caseId, 160),
      scenarioId: `${text(item.caseId, 140)}_internal_01`,
      requestFingerprintSha256: "",
      uploadFingerprintSha256: "",
      expectedColdCostMicrousd: 0,
      modelId: text(modelId, 120),
      operatorNote: "fill fingerprints from actual internal preview only",
    }),
  );
  return Object.freeze({
    version: DRAFT_VERSION,
    registryId: text(registryId, 160),
    sourceDatasetId: text(accuracyDataset.datasetId, 160),
    actualTrafficOnly: true,
    syntheticFingerprintForbidden: true,
    cases: Object.freeze(cases),
  });
}

function buildRealShadowCaseRegistry({
  accuracyDataset,
  draft,
  requireUploadFingerprint = true,
} = {}) {
  const expectedCases = accuracyCases(accuracyDataset);
  const errors = [];
  if (!draft || typeof draft !== "object" || Array.isArray(draft)) {
    errors.push("REGISTRY_DRAFT_OBJECT_REQUIRED");
  }
  if (draft?.version !== DRAFT_VERSION) {
    errors.push("REGISTRY_DRAFT_VERSION_INVALID");
  }
  const draftCases = Array.isArray(draft?.cases) ? draft.cases : [];
  if (draftCases.length === 0) errors.push("REGISTRY_DRAFT_CASES_REQUIRED");

  const byCaseId = new Map();
  for (const [index, item] of draftCases.entries()) {
    const caseId = text(item?.caseId, 160);
    if (!caseId) {
      errors.push(`cases[${index}].caseId required`);
      continue;
    }
    if (byCaseId.has(caseId)) errors.push(`duplicate caseId: ${caseId}`);
    byCaseId.set(caseId, item);
  }

  const expectedCaseIds = expectedCases.map((item) => text(item.caseId, 160));
  const extraCaseIds = [...byCaseId.keys()].filter(
    (caseId) => !expectedCaseIds.includes(caseId),
  );
  for (const caseId of extraCaseIds) errors.push(`unexpected caseId: ${caseId}`);

  const requestFingerprints = new Set();
  const uploadFingerprints = new Set();
  const registryCases = [];

  for (const expected of expectedCases) {
    const caseId = text(expected.caseId, 160);
    const item = byCaseId.get(caseId);
    if (!item) {
      errors.push(`missing caseId: ${caseId}`);
      continue;
    }
    const scenarioId = text(item.scenarioId || `${caseId}_internal_01`, 160);
    const requestFingerprintSha256 = text(
      item.requestFingerprintSha256,
      64,
    ).toLowerCase();
    const uploadFingerprintSha256 = text(
      item.uploadFingerprintSha256,
      64,
    ).toLowerCase();
    const modelId = text(item.modelId || "semantic_profiler_default", 120);
    const expectedColdCostMicrousd = Number(item.expectedColdCostMicrousd);

    if (!scenarioId) errors.push(`${caseId}: scenarioId required`);
    if (!SHA256_RE.test(requestFingerprintSha256)) {
      errors.push(`${caseId}: actual request fingerprint required`);
    }
    if (requireUploadFingerprint && !SHA256_RE.test(uploadFingerprintSha256)) {
      errors.push(`${caseId}: actual upload fingerprint required`);
    }
    if (
      uploadFingerprintSha256 &&
      !SHA256_RE.test(uploadFingerprintSha256)
    ) {
      errors.push(`${caseId}: upload fingerprint invalid`);
    }
    if (
      !Number.isInteger(expectedColdCostMicrousd) ||
      expectedColdCostMicrousd < 0
    ) {
      errors.push(`${caseId}: expectedColdCostMicrousd invalid`);
    }
    if (!modelId) errors.push(`${caseId}: modelId required`);

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

    registryCases.push(
      Object.freeze({
        caseId,
        scenarioId,
        requestFingerprintSha256:
          SHA256_RE.test(requestFingerprintSha256)
            ? requestFingerprintSha256
            : "",
        uploadFingerprintSha256:
          SHA256_RE.test(uploadFingerprintSha256)
            ? uploadFingerprintSha256
            : "",
        expectedColdCostMicrousd:
          Number.isInteger(expectedColdCostMicrousd) &&
          expectedColdCostMicrousd >= 0
            ? expectedColdCostMicrousd
            : 0,
        modelId,
      }),
    );
  }

  const registry = Object.freeze({
    version: REGISTRY_VERSION,
    registryId: text(draft?.registryId, 160),
    cases: Object.freeze(registryCases),
  });
  if (!registry.registryId) errors.push("registryId required");

  const parsed = parseRegistry(JSON.stringify(registry));
  if (!parsed.valid) {
    for (const error of parsed.errors || [parsed.reason]) {
      errors.push(`runtime contract: ${error}`);
    }
  }

  const dedupedErrors = Object.freeze(unique(errors));
  return Object.freeze({
    version: PREPARATION_VERSION,
    valid: dedupedErrors.length === 0,
    reason:
      dedupedErrors.length === 0
        ? "REAL_SHADOW_CASE_REGISTRY_READY"
        : dedupedErrors[0],
    errors: dedupedErrors,
    registry: dedupedErrors.length === 0 ? registry : null,
    registrySha256:
      dedupedErrors.length === 0 ? sha256(JSON.stringify(registry)) : "",
    caseCount: registryCases.length,
    expectedCaseCount: expectedCases.length,
    requestFingerprintCount: requestFingerprints.size,
    uploadFingerprintCount: uploadFingerprints.size,
    actualTrafficOnly: true,
    syntheticFingerprintAllowed: false,
  });
}

function booleanValue(value, fallback) {
  const normalized = text(value, 20).toLowerCase();
  if (!normalized) return fallback;
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off"].includes(normalized)) return false;
  return fallback;
}

function verifyProductionSafety(env = {}) {
  const checks = Object.freeze({
    internalCanaryDisabled: !booleanValue(
      env.QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED,
      false,
    ),
    internalCanaryKillSwitchActive: booleanValue(
      env.QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH,
      true,
    ),
    promotionGateDisabled: !booleanValue(
      env.QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED,
      false,
    ),
    promotionAudienceBlocked:
      text(
        env.QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE || "BLOCKED",
        40,
      ).toUpperCase() === "BLOCKED",
    rolloutZero:
      Number(
        env.QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT || 0,
      ) === 0,
    productionDisabled: !booleanValue(
      env.QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED,
      false,
    ),
    productionMergeDisabled: !booleanValue(
      env.QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED,
      false,
    ),
    productionReadyAssignmentDisabled: !booleanValue(
      env.QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED,
      false,
    ),
    productionRouteDisabled: !booleanValue(
      env.QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED,
      false,
    ),
    productionKillSwitchActive: booleanValue(
      env.QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH,
      true,
    ),
  });
  const failed = Object.entries(checks)
    .filter(([, passed]) => !passed)
    .map(([name]) => name);
  return Object.freeze({
    safe: failed.length === 0,
    checks,
    failed: Object.freeze(failed),
  });
}

function verifyRealShadowPreparation({
  accuracyDataset,
  registry,
  secret,
  allowlistSha256,
  env = {},
} = {}) {
  const errors = [];
  const normalizedSecret = text(secret, 1000);
  if (!BASE64URL_64_RE.test(normalizedSecret)) {
    errors.push("REAL_SHADOW_EVIDENCE_SECRET_MUST_BE_64_CHARACTER_BASE64URL");
  }
  const allowlist = unique(
    text(allowlistSha256, 20000)
      .split(",")
      .map((entry) => entry.trim().toLowerCase())
      .filter(Boolean),
  );
  if (allowlist.length === 0 || allowlist.some((entry) => !SHA256_RE.test(entry))) {
    errors.push("PROMOTION_ALLOWLIST_SHA256_INVALID");
  }

  const draft = Object.freeze({
    version: DRAFT_VERSION,
    registryId: registry?.registryId,
    cases: registry?.cases,
  });
  const registryResult = buildRealShadowCaseRegistry({
    accuracyDataset,
    draft,
    requireUploadFingerprint: true,
  });
  if (!registryResult.valid) errors.push(...registryResult.errors);

  const safety = verifyProductionSafety(env);
  if (!safety.safe) {
    for (const failed of safety.failed) errors.push(`unsafe production flag: ${failed}`);
  }

  const runtimeEnv = {
    ...env,
    QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED: "1",
    QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET: normalizedSecret,
    QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256: allowlist.join(","),
    QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON:
      registryResult.valid ? JSON.stringify(registryResult.registry) : "",
  };
  const runtimeConfig = parseQueryCandidatePlannerRealShadowEvidenceConfig(
    runtimeEnv,
  );
  if (!runtimeConfig.configurationValid) {
    errors.push(`runtime config: ${runtimeConfig.reason}`);
  }

  const dedupedErrors = Object.freeze(unique(errors));
  return Object.freeze({
    version: PREPARATION_VERSION,
    ready: dedupedErrors.length === 0,
    reason:
      dedupedErrors.length === 0
        ? "REAL_SHADOW_PREPARATION_READY_COLLECTOR_STILL_DISABLED"
        : dedupedErrors[0],
    errors: dedupedErrors,
    registrySha256: registryResult.registrySha256,
    secretSha256: BASE64URL_64_RE.test(normalizedSecret)
      ? sha256(normalizedSecret)
      : "",
    allowlistEntryCount: allowlist.filter((entry) => SHA256_RE.test(entry)).length,
    caseCount: registryResult.caseCount,
    productionSafety: safety,
    collectorConfigurationWouldBeValid: runtimeConfig.configurationValid,
    collectorEnabledByThisOperation: false,
    internalCanaryEnabledByThisOperation: false,
    productionPromotionAuthorized: false,
    rawSecretIncluded: false,
    rawIdentityIncluded: false,
  });
}

module.exports = Object.freeze({
  PREPARATION_VERSION,
  DRAFT_VERSION,
  REGISTRY_VERSION,
  SECRET_BYTES,
  SHA256_RE,
  BASE64URL_64_RE,
  sha256,
  generateRealShadowEvidenceSecret,
  buildRealShadowCaseRegistryScaffold,
  buildRealShadowCaseRegistry,
  verifyProductionSafety,
  verifyRealShadowPreparation,
});
