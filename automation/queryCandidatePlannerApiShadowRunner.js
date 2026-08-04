"use strict";

const {
  sha256,
  candidateIdentity,
} = require("./queryCandidatePlannerShadowComparator");

const RUNNER_VERSION =
  "query_candidate_planner_api_shadow_runner_v1";
const INPUT_VERSION =
  "query_candidate_planner_api_shadow_input_v1";

const FORBIDDEN_KEYS = new Set([
  "rows",
  "rawRows",
  "rawData",
  "sampleValues",
  "samples",
  "fileName",
  "originalFileName",
  "originalName",
  "queryTablesKey",
  "tenantId",
  "email",
]);

function text(value) {
  return String(value == null ? "" : value).trim();
}

function number(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function sanitizeColumn(column = {}, index = 0) {
  return Object.freeze({
    columnId:
      text(column.columnId) ||
      text(column.id) ||
      `column_${index + 1}`,
    header:
      text(column.header) ||
      text(column.name) ||
      text(column.label),
    type: text(column.type),
    role: text(column.role || column.semanticRole),
    metricFamily: text(column.metricFamily),
    defaultAggregation: text(column.defaultAggregation),
    unitSemantic: text(column.unitSemantic || column.unit),
    uniqueCount: number(column.uniqueCount),
    uniqueRatio: number(column.uniqueRatio),
  });
}

function sanitizeTable(table = {}, index = 0) {
  const columns = Array.isArray(table.columns) ? table.columns : [];
  return Object.freeze({
    tableId:
      text(table.tableId) ||
      text(table.id) ||
      `table_${index + 1}`,
    sheetOrdinal: number(table.sheetIndex ?? table.sheetOrdinal, index),
    rowCount: number(table.rowCount),
    columnCount: number(table.columnCount, columns.length),
    isPrimary: table.isPrimary === true,
    isVirtual: table.isVirtual === true || table.virtual === true,
    queryable: table.tableUsage?.queryable !== false,
    analysisEligible: table.tableUsage?.analysisEligible !== false,
    templateEligible: table.tableUsage?.templateEligible !== false,
    columns: Object.freeze(columns.slice(0, 120).map(sanitizeColumn)),
  });
}

function candidateArrays(payload = {}) {
  return [
    payload.candidateUiPayload?.recommendedCandidates,
    payload.topCandidates,
    payload.businessTemplateCandidates,
    payload.analysisRecipeCandidates,
    payload.multiSourceCandidates,
    payload.categoryCandidates,
    payload.dashboardCandidates,
    payload.secondaryCandidates,
  ].filter(Array.isArray);
}

function sanitizeCandidate(candidate = {}, index = 0) {
  return Object.freeze({
    candidateId: candidateIdentity(candidate, index),
    candidateType: text(
      candidate.candidateType || candidate.type || candidate.recipeType,
    ),
    recipeId: text(candidate.recipeId || candidate.recipeType),
    operation: text(candidate.operation || candidate.recipeType),
    tableId: text(candidate.tableId || candidate.sourceTableId),
    sourceTableIds: Object.freeze(
      (Array.isArray(candidate.sourceTableIds)
        ? candidate.sourceTableIds
        : []
      ).map(text).filter(Boolean),
    ),
    status: text(
      candidate.status ||
        candidate.feasibilityStatus ||
        candidate.disposition,
    ),
    rank: number(
      candidate.rank ?? candidate.shadowRank ?? candidate.score?.rank,
      index + 1,
    ),
  });
}

function sanitizeCandidates(payload = {}) {
  const seen = new Set();
  const output = [];
  for (const list of candidateArrays(payload)) {
    for (const candidate of list) {
      if (!candidate || typeof candidate !== "object") continue;
      const sanitized = sanitizeCandidate(candidate, output.length);
      if (seen.has(sanitized.candidateId)) continue;
      seen.add(sanitized.candidateId);
      output.push(sanitized);
      if (output.length >= 200) return Object.freeze(output);
    }
  }
  return Object.freeze(output);
}

function buildSafeApiShadowContext({ request = {}, primaryPayload = {} } = {}) {
  const tables = Array.isArray(primaryPayload.normalizedQueryTables)
    ? primaryPayload.normalizedQueryTables
    : [];
  const safeTables = Object.freeze(tables.slice(0, 20).map(sanitizeTable));
  const candidates = sanitizeCandidates(primaryPayload);
  const requestFingerprintSha256 = sha256({
    fileHash: text(primaryPayload.fileHash),
    sheetStateSig: text(primaryPayload.sheetStateSig),
    source: text(primaryPayload.source),
    tableIds: safeTables.map((table) => table.tableId),
    candidateIds: candidates.map((candidate) => candidate.candidateId),
    requestMethod: text(request.method),
    requestPath: text(request.originalUrl || request.path),
  });

  const semanticProfile = Object.freeze({
    version: "api_shadow_semantic_profile_v1",
    source: Object.freeze({
      caseId: `api_shadow_${requestFingerprintSha256.slice(0, 16)}`,
      requestFingerprintSha256,
    }),
    privacy: Object.freeze({
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      originalFileIncluded: false,
      fileNameIncluded: false,
    }),
    tables: safeTables,
  });

  const resolutionItems = Object.freeze(
    candidates.map((candidate) =>
      Object.freeze({
        candidateId: candidate.candidateId,
        recipeId: candidate.recipeId,
        operation: candidate.operation,
        tableId: candidate.tableId,
        sourceTableIds: candidate.sourceTableIds,
        result: candidate.status || "OBSERVED_PRIMARY",
        status: candidate.status || "OBSERVED_PRIMARY",
        rank: candidate.rank,
      }),
    ),
  );

  const candidateResolution = Object.freeze({
    version: "api_shadow_primary_candidate_resolution_v1",
    source: Object.freeze({ requestFingerprintSha256 }),
    items: resolutionItems,
    resolutionSha256: sha256(resolutionItems),
  });

  return Object.freeze({
    version: INPUT_VERSION,
    requestFingerprintSha256,
    semanticProfile,
    candidateResolution,
    candidateFamilyResolution: Object.freeze({
      version: "api_shadow_primary_candidate_family_resolution_v1",
      items: resolutionItems,
    }),
    candidateFeasibilityResolution: Object.freeze({
      version: "api_shadow_primary_candidate_feasibility_resolution_v1",
      items: resolutionItems,
    }),
    rankingResolution: Object.freeze({
      version: "api_shadow_primary_candidate_ranking_resolution_v1",
      items: resolutionItems,
    }),
    primaryCandidateCount: candidates.length,
    tableCount: safeTables.length,
    privacy: Object.freeze({
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      originalFileIncluded: false,
      fileNameIncluded: false,
      queryTablesKeyIncluded: false,
      tenantIdIncluded: false,
      rawPrimaryResponseIncluded: false,
    }),
  });
}

function primaryResponseContractSha256(primaryPayload = {}) {
  const tableSchemas = (Array.isArray(primaryPayload.normalizedQueryTables)
    ? primaryPayload.normalizedQueryTables
    : []
  ).slice(0, 20).map((table, tableIndex) => ({
    tableId: text(table.tableId || table.id || `table_${tableIndex + 1}`),
    rowCount: number(table.rowCount),
    columnIds: (Array.isArray(table.columns) ? table.columns : [])
      .slice(0, 120)
      .map((column, columnIndex) =>
        text(column.columnId || column.id || `column_${columnIndex + 1}`),
      ),
  }));
  const candidateOrders = candidateArrays(primaryPayload).map((list) =>
    list.slice(0, 200).map((candidate, index) =>
      candidateIdentity(candidate, index),
    ),
  );
  return sha256({
    ok: primaryPayload.ok === true,
    source: text(primaryPayload.source),
    fileHash: text(primaryPayload.fileHash),
    sheetStateSig: text(primaryPayload.sheetStateSig),
    topLevelKeys: Object.keys(primaryPayload).sort(),
    tableSchemas,
    candidateOrders,
  });
}

function blockedProvider() {
  const error = new Error("Provider call blocked by Feature Control");
  error.code = "PROVIDER_CALL_BLOCKED_BY_FEATURE_CONTROL";
  const throwBlocked = async () => {
    throw error;
  };
  return new Proxy(throwBlocked, {
    get(_target, property) {
      if (property === "blocked") return true;
      if (property === "code") return error.code;
      return throwBlocked;
    },
    apply() {
      throw error;
    },
  });
}

function assertNoForbiddenKeys(value, path = "root") {
  if (Array.isArray(value)) {
    value.forEach((item, index) =>
      assertNoForbiddenKeys(item, `${path}[${index}]`),
    );
    return true;
  }
  if (!value || typeof value !== "object") return true;
  for (const [key, child] of Object.entries(value)) {
    if (FORBIDDEN_KEYS.has(key)) {
      throw new Error(`Forbidden API shadow input key: ${path}.${key}`);
    }
    assertNoForbiddenKeys(child, `${path}.${key}`);
  }
  return true;
}

async function runQueryCandidatePlannerApiShadow({
  safeContext,
  providerDecision,
  cacheReadDecision,
  cacheWriteDecision,
  signal,
} = {}) {
  assertNoForbiddenKeys(safeContext);
  const bridge = require("./queryCandidatePlannerShadowBridge");
  if (typeof bridge.runCandidatePlannerLiveShadow !== "function") {
    const error = new Error(
      "queryCandidatePlannerShadowBridge.runCandidatePlannerLiveShadow is required",
    );
    error.code = "SHADOW_BRIDGE_EXPORT_MISSING";
    throw error;
  }

  const provider = providerDecision?.allowed ? undefined : blockedProvider();
  return bridge.runCandidatePlannerLiveShadow({
    caseId: safeContext.semanticProfile.source.caseId,
    semanticProfile: safeContext.semanticProfile,
    candidateResolution: safeContext.candidateResolution,
    candidateFamilyResolution: safeContext.candidateFamilyResolution,
    candidateFeasibilityResolution:
      safeContext.candidateFeasibilityResolution,
    rankingResolution: safeContext.rankingResolution,
    sourceCandidateResolution: safeContext.candidateResolution,
    provider,
    providerCallAllowed: providerDecision?.allowed === true,
    cacheReadAllowed: cacheReadDecision?.allowed === true,
    cacheWriteAllowed: cacheWriteDecision?.allowed === true,
    apiShadow: true,
    requestFingerprintSha256: safeContext.requestFingerprintSha256,
    signal,
  });
}

module.exports = Object.freeze({
  RUNNER_VERSION,
  INPUT_VERSION,
  FORBIDDEN_KEYS,
  buildSafeApiShadowContext,
  primaryResponseContractSha256,
  assertNoForbiddenKeys,
  runQueryCandidatePlannerApiShadow,
});
