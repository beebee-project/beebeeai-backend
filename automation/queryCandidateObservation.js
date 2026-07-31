const crypto = require("crypto");

const QUERY_CANDIDATE_OBSERVATION_VERSION = "query_candidate_observation_v1";
const QUERY_CANDIDATE_BASELINE_VERSION = "query_candidate_baseline_v1";

const CANDIDATE_GROUPS = Object.freeze([
  "analysisRecipeCandidates",
  "dashboardCandidates",
  "categoryCandidates",
  "businessTemplateCandidates",
  "multiSourceCandidates",
  "topCandidates",
  "secondaryCandidates",
]);

const HIDDEN_LEVELS = new Set([
  "hidden",
  "disabled",
  "quarantined",
  "definition-only",
  "definition_only",
]);

function normalizeText(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .trim();
}

function uniqueSorted(values = []) {
  return Array.from(
    new Set(
      (Array.isArray(values) ? values : []).map(normalizeText).filter(Boolean),
    ),
  ).sort((left, right) => left.localeCompare(right, "ko"));
}

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (!value || typeof value !== "object") return value;

  return Object.keys(value)
    .sort()
    .reduce((result, key) => {
      const item = value[key];
      if (item !== undefined) result[key] = canonicalize(item);
      return result;
    }, {});
}

function stableStringify(value) {
  return JSON.stringify(canonicalize(value));
}

function sha256(value) {
  const input = Buffer.isBuffer(value)
    ? value
    : Buffer.from(typeof value === "string" ? value : stableStringify(value));
  return crypto.createHash("sha256").update(input).digest("hex");
}

function numericOrZero(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function getQueryTables(payload = {}) {
  if (Array.isArray(payload.normalizedQueryTables)) {
    return payload.normalizedQueryTables;
  }
  if (Array.isArray(payload.tables)) return payload.tables;
  if (Array.isArray(payload.queryTables)) return payload.queryTables;
  return [];
}

function tableHeaders(table = {}) {
  const direct = asArray(table.headers || table.columns)
    .map((column) => {
      if (typeof column === "string") return column;
      return (
        column?.header ||
        column?.name ||
        column?.label ||
        column?.key ||
        column?.sourceHeader ||
        ""
      );
    })
    .map(normalizeText)
    .filter(Boolean);

  if (direct.length) return uniqueSorted(direct);

  const rows = Array.isArray(table.rows) ? table.rows : [];
  const firstObjectRow = rows.find(
    (row) => row && typeof row === "object" && !Array.isArray(row),
  );
  return firstObjectRow ? uniqueSorted(Object.keys(firstObjectRow)) : [];
}

function tableColumnRoles(table = {}) {
  const columns = asArray(table.columns || table.headers);
  return columns
    .map((column) => {
      if (!column || typeof column !== "object") return null;
      const header = normalizeText(
        column.header || column.name || column.label || column.key || "",
      );
      const role = normalizeText(
        column.semanticRole ||
          column.role ||
          column.detectedRole ||
          column.metricRole ||
          "",
      );
      const type = normalizeText(
        column.dataType || column.type || column.inferredType || "",
      );
      if (!header && !role && !type) return null;
      return { header, role, type };
    })
    .filter(Boolean)
    .sort((left, right) =>
      `${left.header}|${left.role}|${left.type}`.localeCompare(
        `${right.header}|${right.role}|${right.type}`,
        "ko",
      ),
    );
}

function summarizeQueryTable(table = {}, index = 0) {
  const usage = table.usage || table.sourceTablePolicy || {};
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const headers = tableHeaders(table);

  return {
    index,
    tableId: normalizeText(table.tableId || table.id || `table_${index + 1}`),
    sourceTableId: normalizeText(
      table.sourceTableId || table.transformation?.sourceTableId || "",
    ),
    sourceSheetName: normalizeText(
      table.sourceSheetName || table.sheetName || table.name || "",
    ),
    virtual: Boolean(
      table.isVirtual === true ||
      table.virtual === true ||
      table.transformation?.virtual === true,
    ),
    primary: Boolean(table.isPrimary === true || usage.primary === true),
    analysisEligible: Boolean(
      table.analysisEligible === true ||
      usage.analysisEligible === true ||
      usage.analysis === true,
    ),
    templateEligible: Boolean(
      table.templateEligible === true ||
      usage.templateEligible === true ||
      usage.template === true,
    ),
    rowCount: numericOrZero(
      table.rowCount ||
        table.dataRowCount ||
        table.stats?.rowCount ||
        rows.length,
    ),
    columnCount: numericOrZero(
      table.columnCount || table.stats?.columnCount || headers.length,
    ),
    headers,
    columnRoles: tableColumnRoles(table),
    dataQualityStatus: normalizeText(
      table.dataQuality?.status ||
        table.diagnostics?.inheritedQuality?.dataQuality?.status ||
        "",
    ),
  };
}

function candidateIdentifier(candidate = {}, index = 0) {
  return normalizeText(
    candidate.candidateId ||
      candidate.id ||
      candidate.templateId ||
      candidate.recipeId ||
      `${candidate.candidateType || candidate.type || "candidate"}_${index + 1}`,
  );
}

function candidateExposureLevel(candidate = {}) {
  return normalizeText(
    candidate.exposureLevel ||
      candidate.implementationLevel ||
      candidate.frontExposure?.exposureLevel ||
      candidate.templateMatch?.implementationLevel ||
      "",
  ).toLowerCase();
}

function classifyObservedCandidate(candidate = {}) {
  const exposureLevel = candidateExposureLevel(candidate);
  const hidden =
    candidate.frontExposure?.hideFromFrontend === true ||
    HIDDEN_LEVELS.has(exposureLevel);

  if (hidden) return "hidden";
  if (
    candidate.executionEligible === false ||
    candidate.recommendedEligible === false ||
    candidate.conditional === true ||
    candidate.requiresConfirmation === true
  ) {
    return "conditional";
  }
  return "eligible";
}

function summarizeCandidate(candidate = {}, index = 0, groupName = "") {
  const score = candidate.score || {};
  const rankScore = Number(candidate.rankScore);
  const confidence = Number(candidate.confidence);
  const priority = Number(candidate.priority);

  return {
    candidateId: candidateIdentifier(candidate, index),
    candidateType: normalizeText(
      candidate.candidateType || candidate.recipeType || candidate.type || "",
    ),
    title: normalizeText(
      candidate.title || candidate.label || candidate.name || "",
    ),
    templateId: normalizeText(candidate.templateId || ""),
    recipeIds: uniqueSorted([
      ...asArray(candidate.recipeIds),
      candidate.recipeId,
    ]),
    sourceTableIds: uniqueSorted([
      ...asArray(candidate.sourceTableIds),
      candidate.sourceTableId,
    ]),
    sourceSheetNames: uniqueSorted([
      ...asArray(candidate.sourceSheetNames),
      candidate.sourceSheetName,
    ]),
    outputTypes: uniqueSorted(candidate.outputTypes),
    reasonCodes: uniqueSorted(candidate.reasonCodes),
    groupName,
    observedClass: classifyObservedCandidate(candidate),
    exposureLevel: candidateExposureLevel(candidate),
    executionEligible: candidate.executionEligible !== false,
    recommendedEligible: candidate.recommendedEligible !== false,
    rank: numericOrZero(candidate.rank || index + 1),
    rankScore: Number.isFinite(rankScore) ? Number(rankScore.toFixed(6)) : null,
    confidence: Number.isFinite(confidence)
      ? Number(confidence.toFixed(6))
      : null,
    priority: Number.isFinite(priority) ? Number(priority.toFixed(6)) : null,
    candidateScoreVersion: normalizeText(
      candidate.candidateScoreVersion || score.version || "",
    ),
    rankingTier: normalizeText(candidate.rankingTier || ""),
  };
}

function candidateGroupsFromPayload(payload = {}) {
  const groups = {};
  for (const groupName of CANDIDATE_GROUPS) {
    groups[groupName] = asArray(payload[groupName]).map((candidate, index) =>
      summarizeCandidate(candidate, index, groupName),
    );
  }

  groups.uiRecommendedCandidates = asArray(
    payload.candidateUiPayload?.recommendedCandidates,
  ).map((candidate, index) =>
    summarizeCandidate(candidate, index, "uiRecommendedCandidates"),
  );

  return groups;
}

function deduplicateCandidates(groups = {}) {
  const preferredOrder = [
    "uiRecommendedCandidates",
    "topCandidates",
    "businessTemplateCandidates",
    "analysisRecipeCandidates",
    "dashboardCandidates",
    "categoryCandidates",
    "multiSourceCandidates",
    "secondaryCandidates",
  ];
  const seen = new Set();
  const rows = [];

  for (const groupName of preferredOrder) {
    for (const candidate of asArray(groups[groupName])) {
      const key = candidate.candidateId;
      if (!key || seen.has(key)) continue;
      seen.add(key);
      rows.push(candidate);
    }
  }
  return rows;
}

function candidateCountMap(groups = {}) {
  return Object.keys(groups)
    .sort()
    .reduce((result, key) => {
      result[key] = asArray(groups[key]).length;
      return result;
    }, {});
}

function buildQueryCandidateObservation({
  caseId = "",
  fileName = "",
  queryJson = {},
  candidatePayload = {},
  sourceQuerySha256 = "",
  sourceCandidateSha256 = "",
} = {}) {
  const tables = getQueryTables(queryJson).map(summarizeQueryTable);
  const groups = candidateGroupsFromPayload(candidatePayload);
  const candidates = deduplicateCandidates(groups);
  const idsByClass = {
    eligible: uniqueSorted(
      candidates
        .filter((candidate) => candidate.observedClass === "eligible")
        .map((candidate) => candidate.candidateId),
    ),
    conditional: uniqueSorted(
      candidates
        .filter((candidate) => candidate.observedClass === "conditional")
        .map((candidate) => candidate.candidateId),
    ),
    hidden: uniqueSorted(
      candidates
        .filter((candidate) => candidate.observedClass === "hidden")
        .map((candidate) => candidate.candidateId),
    ),
  };

  const topOrder = asArray(groups.topCandidates)
    .map((candidate) => candidate.candidateId)
    .filter(Boolean);
  const uiRecommendedOrder = asArray(groups.uiRecommendedCandidates)
    .map((candidate) => candidate.candidateId)
    .filter(Boolean);

  const observation = {
    version: QUERY_CANDIDATE_OBSERVATION_VERSION,
    caseId: normalizeText(caseId),
    fileName: normalizeText(fileName),
    source: {
      queryJsonSha256: normalizeText(sourceQuerySha256) || sha256(queryJson),
      candidatePayloadSha256:
        normalizeText(sourceCandidateSha256) || sha256(candidatePayload),
    },
    queryShape: {
      tableCount: tables.length,
      normalizedTableCount: Array.isArray(queryJson.normalizedQueryTables)
        ? queryJson.normalizedQueryTables.length
        : 0,
      physicalTableCount: Array.isArray(queryJson.tables)
        ? queryJson.tables.length
        : 0,
      primaryTableCount: tables.filter((table) => table.primary).length,
      analysisEligibleCount: tables.filter((table) => table.analysisEligible)
        .length,
      templateEligibleCount: tables.filter((table) => table.templateEligible)
        .length,
      shapeSha256: sha256(tables),
      tables,
    },
    candidateObservation: {
      candidateContractVersion: normalizeText(
        candidatePayload.candidateContract?.version ||
          candidatePayload.candidateGeneration?.candidateContract?.version ||
          "",
      ),
      candidateScoringVersion: normalizeText(
        candidatePayload.candidateScoring?.version ||
          candidatePayload.candidateGeneration?.candidateScoring?.version ||
          "",
      ),
      candidateUiPayloadVersion: normalizeText(
        candidatePayload.candidateUiPayload?.version || "",
      ),
      counts: candidateCountMap(groups),
      idsByClass,
      topOrder,
      uiRecommendedOrder,
      candidates,
    },
  };

  observation.observationSha256 = sha256({
    ...observation,
    observationSha256: undefined,
  });
  return observation;
}

function buildQueryCandidateBaseline(observation = {}) {
  const candidateObservation = observation.candidateObservation || {};
  const baseline = {
    version: QUERY_CANDIDATE_BASELINE_VERSION,
    observationVersion: observation.version || "",
    caseId: observation.caseId || "",
    fileName: observation.fileName || "",
    source: observation.source || {},
    expected: {
      queryShapeSha256: observation.queryShape?.shapeSha256 || "",
      tableCount: numericOrZero(observation.queryShape?.tableCount),
      analysisEligibleCount: numericOrZero(
        observation.queryShape?.analysisEligibleCount,
      ),
      candidateContractVersion:
        candidateObservation.candidateContractVersion || "",
      candidateScoringVersion:
        candidateObservation.candidateScoringVersion || "",
      candidateUiPayloadVersion:
        candidateObservation.candidateUiPayloadVersion || "",
      eligibleCandidateIds: asArray(candidateObservation.idsByClass?.eligible),
      conditionalCandidateIds: asArray(
        candidateObservation.idsByClass?.conditional,
      ),
      hiddenCandidateIds: asArray(candidateObservation.idsByClass?.hidden),
      topCandidateOrder: asArray(candidateObservation.topOrder),
      uiRecommendedOrder: asArray(candidateObservation.uiRecommendedOrder),
    },
  };
  baseline.baselineSha256 = sha256({ ...baseline, baselineSha256: undefined });
  return baseline;
}

function compareArray(label, actual = [], expected = [], differences = []) {
  const actualValue = JSON.stringify(asArray(actual));
  const expectedValue = JSON.stringify(asArray(expected));
  if (actualValue !== expectedValue) {
    differences.push({
      label,
      expected: asArray(expected),
      actual: asArray(actual),
    });
  }
}

function compareQueryCandidateBaseline(observation = {}, baseline = {}) {
  const expected = baseline.expected || {};
  const actual = observation.candidateObservation || {};
  const differences = [];

  const scalarChecks = [
    ["queryShapeSha256", observation.queryShape?.shapeSha256 || ""],
    ["tableCount", numericOrZero(observation.queryShape?.tableCount)],
    [
      "analysisEligibleCount",
      numericOrZero(observation.queryShape?.analysisEligibleCount),
    ],
    ["candidateContractVersion", actual.candidateContractVersion || ""],
    ["candidateScoringVersion", actual.candidateScoringVersion || ""],
    ["candidateUiPayloadVersion", actual.candidateUiPayloadVersion || ""],
  ];

  for (const [label, actualValue] of scalarChecks) {
    if (actualValue !== expected[label]) {
      differences.push({
        label,
        expected: expected[label],
        actual: actualValue,
      });
    }
  }

  compareArray(
    "eligibleCandidateIds",
    actual.idsByClass?.eligible,
    expected.eligibleCandidateIds,
    differences,
  );
  compareArray(
    "conditionalCandidateIds",
    actual.idsByClass?.conditional,
    expected.conditionalCandidateIds,
    differences,
  );
  compareArray(
    "hiddenCandidateIds",
    actual.idsByClass?.hidden,
    expected.hiddenCandidateIds,
    differences,
  );
  compareArray(
    "topCandidateOrder",
    actual.topOrder,
    expected.topCandidateOrder,
    differences,
  );
  compareArray(
    "uiRecommendedOrder",
    actual.uiRecommendedOrder,
    expected.uiRecommendedOrder,
    differences,
  );

  return {
    version: "query_candidate_baseline_compare_v1",
    caseId: observation.caseId || baseline.caseId || "",
    pass: differences.length === 0,
    differenceCount: differences.length,
    differences,
  };
}

module.exports = {
  QUERY_CANDIDATE_OBSERVATION_VERSION,
  QUERY_CANDIDATE_BASELINE_VERSION,
  CANDIDATE_GROUPS,
  normalizeText,
  canonicalize,
  stableStringify,
  sha256,
  getQueryTables,
  summarizeQueryTable,
  summarizeCandidate,
  classifyObservedCandidate,
  buildQueryCandidateObservation,
  buildQueryCandidateBaseline,
  compareQueryCandidateBaseline,
};
