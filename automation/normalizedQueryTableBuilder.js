const {
  COLUMN_ROLE_PATTERNS,
  BOOLEAN_VALUES,
  COLUMN_INFERENCE_THRESHOLDS,
} = require("./config/columnRoleConfig");

const DIAGNOSTICS_VERSION = "normalized_query_diagnostics_v1";
const NUMERIC_MEASURE_ROLE_VERSION = "numeric_measure_role_v1";

const NUMERIC_MEASURE_HEADER_PATTERN =
  /(재고|잔고|잔량|잔여|보유|수량|입고|출고|사용량|생산량|판매량|금액|매출|매입|비용|원가|단가|가격|실적|목표|인원|건수|시간|거리|중량|무게|용량|stock|inventory|balance|quantity|qty|amount|value|price|cost|revenue|sales|count|total)/i;

const EXPLICIT_DATE_HEADER_PATTERN =
  /(일자|날짜|연월|년월|기준월|기준일|기간|시점|분기|년도|연도|date|month|period|quarter|year)/i;

const DEFAULT_TABLE_USAGE = Object.freeze({
  version: "table_usage_quality_v1",
  queryable: true,
  analysisEligible: true,
  templateEligible: true,
  reasons: ["TABLE_USAGE_NOT_PROVIDED"],
  metrics: {},
});

function normalizeTableUsage(table = {}) {
  const usage = table.tableUsage || table.usage || null;
  if (!usage || typeof usage !== "object") return { ...DEFAULT_TABLE_USAGE };

  return {
    version: usage.version || DEFAULT_TABLE_USAGE.version,
    queryable: usage.queryable !== false,
    analysisEligible: usage.analysisEligible !== false,
    templateEligible: usage.templateEligible !== false,
    reasons:
      Array.isArray(usage.reasons) && usage.reasons.length
        ? usage.reasons
        : DEFAULT_TABLE_USAGE.reasons,
    metrics: usage.metrics || {},
  };
}

function isAnalysisEligibleTable(table = {}) {
  return normalizeTableUsage(table).analysisEligible !== false;
}

function inheritVirtualTableUsage(
  sourceTable = {},
  transformationType = "virtual",
) {
  const sourceUsage = normalizeTableUsage(sourceTable);
  return {
    ...sourceUsage,
    queryable: true,
    analysisEligible: sourceUsage.analysisEligible !== false,
    templateEligible: sourceUsage.templateEligible !== false,
    reasons: [
      ...(sourceUsage.reasons || []),
      `VIRTUAL_TABLE_FROM_${String(transformationType || "virtual").toUpperCase()}`,
    ],
  };
}

function isBlank(value) {
  return value == null || String(value).trim() === "";
}

function text(value = "") {
  return String(value ?? "");
}

function toNumberOrNull(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  if (typeof value !== "string") return null;

  const normalized = value.replace(/,/g, "").trim();
  if (!normalized) return null;

  const num = Number(normalized);
  return Number.isFinite(num) ? num : null;
}

function isDateLike(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;

  if (typeof value !== "string") return false;

  const t = value.trim();
  if (!t) return false;

  return (
    /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(t) ||
    /^\d{4}[-/.]\d{1,2}$/.test(t) ||
    /^\d{4}\s*년\s*\d{1,2}\s*월?$/.test(t)
  );
}

function normalizeHeader(value = "") {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ");
}

function stripHeaderUnitSuffix(value = "") {
  return normalizeHeader(value)
    .replace(/\([^)]*\)/g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/（[^）]*）/g, "")
    .trim();
}

function normalizeHeaderForDiagnostics(value = "") {
  return stripHeaderUnitSuffix(value)
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "")
    .trim();
}

function canonicalKeyFromHeader(value = "", fallback = "") {
  const normalized = normalizeHeaderForDiagnostics(value)
    .replace(/[^\p{Letter}\p{Number}]+/gu, "_")
    .replace(/^_+|_+$/g, "");
  return normalized || fallback;
}

function compactText(value = "") {
  return String(value ?? "")
    .replace(/[_|/\\]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeMonthNumber(value = "") {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 1 || n > 12) return "";
  return String(n).padStart(2, "0");
}

function extractYear(value = "") {
  const matched = compactText(value).match(
    /(?:^|[^\d])((?:19|20)\d{2})(?:[^\d]|$)/,
  );
  return matched ? matched[1] : "";
}

function extractMonth(value = "") {
  const s = compactText(value);
  const explicit = s.match(/(?:^|[^\d])(0?[1-9]|1[0-2])\s*월/);
  if (explicit) return normalizeMonthNumber(explicit[1]);

  const yearMonth = s.match(
    /(?:19|20)\d{2}\s*[-./년]\s*(0?[1-9]|1[0-2])(?:\D|$)/,
  );
  if (yearMonth) return normalizeMonthNumber(yearMonth[1]);

  return "";
}

function extractQuarter(value = "") {
  const s = compactText(value).toUpperCase();
  const q = s.match(/(?:^|\s)Q\s*([1-4])(?:\s|$)/);
  if (q) return `Q${q[1]}`;

  const korean = s.match(/(?:^|\s)([1-4])\s*분기(?:\s|$)/);
  return korean ? `Q${korean[1]}` : "";
}

function removeTemporalTokens(value = "") {
  return compactText(value)
    .replace(/(?:19|20)\d{2}\s*[-./년]?\s*(0?[1-9]|1[0-2])?\s*월?/g, " ")
    .replace(/(?:^|\s)Q\s*[1-4](?:\s|$)/gi, " ")
    .replace(/(?:^|\s)[1-4]\s*분기(?:\s|$)/g, " ")
    .replace(/(?:^|\s)(0?[1-9]|1[0-2])\s*월(?:\s|$)/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseTemporalHeader(header = "") {
  const raw = compactText(header);
  if (!raw) return null;

  const year = extractYear(raw);
  const month = extractMonth(raw);
  const quarter = extractQuarter(raw);

  if (!year && !month && !quarter) return null;

  let periodLabel = raw;
  if (year && month) periodLabel = `${year}-${month}`;
  else if (year && quarter) periodLabel = `${year}-${quarter}`;
  else if (year) periodLabel = year;
  else if (month) periodLabel = `${month}월`;
  else if (quarter) periodLabel = quarter;

  const metricLabel = removeTemporalTokens(raw) || "지표값";

  return {
    raw,
    year,
    month,
    quarter,
    periodLabel,
    metricLabel,
  };
}

function parseTemporalValue(value = "") {
  const raw = compactText(value);
  if (!raw) return null;
  const parsed = parseTemporalHeader(raw);
  if (parsed) return parsed;

  const date = value instanceof Date ? value : new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    const year = String(date.getFullYear());
    const month = String(date.getMonth() + 1).padStart(2, "0");
    return {
      raw,
      year,
      month,
      quarter: "",
      periodLabel: `${year}-${month}`,
      metricLabel: "지표값",
    };
  }

  return null;
}

function isTemporalHeaderColumn(column = {}) {
  return Boolean(
    parseTemporalHeader(column.header || column.originalHeader || ""),
  );
}

function numericRatioForColumn(rows = [], column = {}) {
  const values = rows
    .map((row) => getRowValueByColumn(row, column, column.index ?? 0))
    .filter((value) => !isBlank(value));

  if (!values.length) return 0;

  const numeric = values.filter(
    (value) => toNumberOrNull(value) != null,
  ).length;
  return numeric / values.length;
}

function uniqueHeader(base = "", used = new Set()) {
  const fallback = String(base || "값").trim() || "값";
  let candidate = fallback;
  let index = 2;

  while (used.has(candidate)) {
    candidate = `${fallback}_${index}`;
    index += 1;
  }

  used.add(candidate);
  return candidate;
}

function makeVirtualColumn({
  header,
  type = "string",
  role = "dimension",
  sourceColumn = null,
} = {}) {
  return {
    header,
    originalHeader: header,
    key: header,
    canonicalKey: canonicalKeyFromHeader(header, header),
    type,
    role,
    sourceColumnHeader: sourceColumn?.header || null,
    sourceColumnKey: sourceColumn?.canonicalKey || sourceColumn?.key || null,
  };
}

function inferColumnType(values = []) {
  const sample = values
    .filter((v) => !isBlank(v))
    .slice(0, COLUMN_INFERENCE_THRESHOLDS.sampleSize);
  if (!sample.length) return "unknown";

  const numberCount = sample.filter((v) => toNumberOrNull(v) != null).length;
  const dateCount = sample.filter(isDateLike).length;
  const booleanCount = sample.filter(
    (v) =>
      typeof v === "boolean" ||
      BOOLEAN_VALUES.includes(String(v).trim().toLowerCase()),
  ).length;

  const ratio = (count) => count / sample.length;

  if (ratio(dateCount) >= COLUMN_INFERENCE_THRESHOLDS.dateRatio) return "date";
  if (ratio(numberCount) >= COLUMN_INFERENCE_THRESHOLDS.numberRatio)
    return "number";
  if (ratio(booleanCount) >= COLUMN_INFERENCE_THRESHOLDS.booleanRatio)
    return "boolean";

  return "string";
}

function hasNumericMeasureEvidence(
  header = "",
  type = "unknown",
  profile = {},
) {
  const normalizedHeader = normalizeHeaderForDiagnostics(header);
  const numericRatio = Number(profile.numericRatio || 0);
  const dateRatio = Number(profile.dateRatio || 0);

  return Boolean(
    type === "number" &&
    numericRatio >= 0.7 &&
    dateRatio < 0.5 &&
    NUMERIC_MEASURE_HEADER_PATTERN.test(normalizedHeader),
  );
}

function hasExplicitDateEvidence(header = "", type = "unknown", profile = {}) {
  const normalizedHeader = normalizeHeaderForDiagnostics(header);
  const dateRatio = Number(profile.dateRatio || 0);

  return Boolean(
    type === "date" ||
    dateRatio >= 0.5 ||
    EXPLICIT_DATE_HEADER_PATTERN.test(normalizedHeader),
  );
}

function inferColumnRole(header = "", type = "unknown", profile = {}) {
  const t = String(header).toLowerCase();

  // 숫자형 측정값은 헤더 일부에 "월" 등의 문자열이 포함돼도
  // 실제 날짜값 증거가 없으면 metric으로 우선 판정한다.
  if (hasNumericMeasureEvidence(header, type, profile)) return "metric";

  if (
    hasExplicitDateEvidence(header, type, profile) ||
    COLUMN_ROLE_PATTERNS.date.test(t)
  ) {
    return "date";
  }

  if (COLUMN_ROLE_PATTERNS.id.test(t)) return "id";
  if (COLUMN_ROLE_PATTERNS.status.test(t)) return "status";
  if (type === "number" && COLUMN_ROLE_PATTERNS.metric.test(t)) return "metric";
  if (type === "category") return "dimension";
  if (type === "string" || type === "boolean") return "dimension";

  return "unknown";
}

function getRowValueByColumn(row = {}, column = {}, index = 0) {
  if (!row) return undefined;

  if (Array.isArray(row)) {
    return row[index];
  }

  const keys = [
    column.key,
    column.canonicalKey,
    column.accessor,
    column.name,
    column.header,
    column.originalHeader,
  ].filter(Boolean);

  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
  }

  const normalizedTargets = keys
    .map(normalizeHeaderForDiagnostics)
    .filter(Boolean);
  if (!normalizedTargets.length) return undefined;

  const matchedKey = Object.keys(row).find((key) =>
    normalizedTargets.includes(normalizeHeaderForDiagnostics(key)),
  );

  return matchedKey ? row[matchedKey] : undefined;
}

function getColumnValues(table = {}, header = "", index = 0, column = {}) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const resolvedColumn = { ...column, header: column.header || header };

  return rows.map((row) => getRowValueByColumn(row, resolvedColumn, index));
}

function calculateEmptyRatio(rows = []) {
  if (!rows.length) return 1;

  let total = 0;
  let empty = 0;

  rows.forEach((row) => {
    Object.values(row || {}).forEach((value) => {
      total += 1;

      if (isBlank(value)) {
        empty += 1;
      }
    });
  });

  if (!total) return 1;

  return empty / total;
}

function calculateTypeConsistency(columns = []) {
  if (!columns.length) return 0;

  const valid = columns.filter(
    (column) => column.type && column.type !== "unknown",
  );

  return valid.length / columns.length;
}

function calculateHeaderConfidence(columns = []) {
  if (!columns.length) return 0;

  const validHeaders = columns.filter((column) => {
    const header = String(column.header || "").trim();

    return header && !/^column_\d+$/i.test(header) && header.length >= 2;
  });

  return validHeaders.length / columns.length;
}

function countDuplicateHeaders(columns = []) {
  const seen = new Map();
  let duplicateCount = 0;

  for (const column of columns) {
    const key = normalizeHeaderForDiagnostics(
      column.header || column.originalHeader || "",
    );
    if (!key) continue;
    const count = seen.get(key) || 0;
    if (count > 0) duplicateCount += 1;
    seen.set(key, count + 1);
  }

  return duplicateCount;
}

function buildWarnings({
  rows = [],
  columns = [],
  emptyRatio = 0,
  headerConfidence = 0,
  typeConsistency = 0,
}) {
  const warnings = [];

  if (!rows.length) {
    warnings.push("EMPTY_TABLE");
  }

  if (emptyRatio >= COLUMN_INFERENCE_THRESHOLDS.emptyRatioWarning) {
    warnings.push("MANY_EMPTY_CELLS");
  }

  if (headerConfidence < COLUMN_INFERENCE_THRESHOLDS.headerConfidenceWarning) {
    warnings.push("LOW_CONFIDENCE_HEADER");
  }

  if (typeConsistency < COLUMN_INFERENCE_THRESHOLDS.typeConsistencyWarning) {
    warnings.push("LOW_TYPE_CONSISTENCY");
  }

  if (columns.length <= 1) {
    warnings.push("TOO_FEW_COLUMNS");
  }

  if (countDuplicateHeaders(columns) > 0) {
    warnings.push("DUPLICATE_NORMALIZED_HEADERS");
  }

  return warnings;
}

function calculateConfidence({
  emptyRatio = 0,
  headerConfidence = 0,
  typeConsistency = 0,
}) {
  const weights = COLUMN_INFERENCE_THRESHOLDS.confidenceWeights;
  const score =
    headerConfidence * weights.headerConfidence +
    typeConsistency * weights.typeConsistency +
    (1 - emptyRatio) * weights.nonEmptyRatio;

  return Number(score.toFixed(2));
}

function analyzeColumnValues(values = []) {
  const sample = values.filter((v) => !isBlank(v));
  const totalRows = values.length;
  const nonEmptyCount = sample.length;
  const emptyRatio = totalRows > 0 ? 1 - nonEmptyCount / totalRows : 1;

  const numberCount = sample.filter((v) => toNumberOrNull(v) != null).length;
  const dateCount = sample.filter(isDateLike).length;
  const booleanCount = sample.filter(
    (v) =>
      typeof v === "boolean" ||
      BOOLEAN_VALUES.includes(String(v).trim().toLowerCase()),
  ).length;
  const textCount = Math.max(
    0,
    nonEmptyCount - numberCount - dateCount - booleanCount,
  );

  const ratio = (count) => (nonEmptyCount > 0 ? count / nonEmptyCount : 0);
  const normalizedValues = sample.map((v) => text(v).trim()).filter(Boolean);
  const uniqueValues = Array.from(new Set(normalizedValues));
  const uniqueCount = uniqueValues.length;
  const uniqueRatio = nonEmptyCount > 0 ? uniqueCount / nonEmptyCount : 0;

  const typeRatios = {
    number: ratio(numberCount),
    date: ratio(dateCount),
    boolean: ratio(booleanCount),
    text: ratio(textCount),
  };
  const dominantTypeRatio = Math.max(...Object.values(typeRatios), 0);

  return {
    totalRows,
    nonEmptyCount,
    emptyRatio: Number(emptyRatio.toFixed(3)),
    uniqueCount,
    uniqueRatio: Number(uniqueRatio.toFixed(3)),
    numericRatio: Number(typeRatios.number.toFixed(3)),
    dateRatio: Number(typeRatios.date.toFixed(3)),
    booleanRatio: Number(typeRatios.boolean.toFixed(3)),
    textRatio: Number(typeRatios.text.toFixed(3)),
    dominantTypeRatio: Number(dominantTypeRatio.toFixed(3)),
    sampleValues: uniqueValues.slice(0, 10),
  };
}

function confidenceFromColumnProfile(
  profile = {},
  type = "unknown",
  role = "unknown",
) {
  let typeConfidence = Number(profile.dominantTypeRatio || 0);
  let roleConfidence = 0.35;

  if (type && type !== "unknown") roleConfidence += 0.25;
  if (role && role !== "unknown") roleConfidence += 0.25;
  if (Number(profile.emptyRatio || 0) <= 0.2) roleConfidence += 0.1;
  if (Number(profile.uniqueRatio || 0) <= 0.6 && role === "dimension") {
    roleConfidence += 0.05;
  }
  if (role === "metric" && Number(profile.numericRatio || 0) >= 0.7) {
    roleConfidence += 0.1;
  }
  if (role === "date" && Number(profile.dateRatio || 0) >= 0.5) {
    roleConfidence += 0.1;
  }

  return {
    typeConfidence: Number(Math.min(1, typeConfidence).toFixed(2)),
    roleConfidence: Number(Math.min(1, roleConfidence).toFixed(2)),
  };
}

function normalizeColumn(column = {}, index = 0, table = {}) {
  const header =
    column.header ||
    column.name ||
    column.key ||
    column.label ||
    `column_${index + 1}`;

  const values = getColumnValues(table, header, column.index ?? index, column);
  const profile = analyzeColumnValues(values);
  const inferredType =
    column.type || column.valueType || inferColumnType(values);
  const inferredRole =
    column.role || inferColumnRole(header, inferredType, profile);
  const confidence = confidenceFromColumnProfile(
    profile,
    inferredType,
    inferredRole,
  );
  const normalizedHeader = normalizeHeaderForDiagnostics(header);

  return {
    header,
    originalHeader: column.originalHeader || header,
    normalizedHeader,
    canonicalKey:
      column.canonicalKey ||
      column.key ||
      canonicalKeyFromHeader(header, `column_${index + 1}`),
    index: column.index ?? index,
    type: inferredType,
    role: inferredRole,
    quality: Number.isFinite(Number(column.quality))
      ? Number(column.quality)
      : null,
    profile: {
      emptyRatio: profile.emptyRatio,
      nonEmptyCount: profile.nonEmptyCount,
      uniqueCount: profile.uniqueCount,
      uniqueRatio: profile.uniqueRatio,
      numericRatio: profile.numericRatio,
      dateRatio: profile.dateRatio,
      booleanRatio: profile.booleanRatio,
      textRatio: profile.textRatio,
      dominantTypeRatio: profile.dominantTypeRatio,
      sampleValues: profile.sampleValues,
    },
    diagnostics: {
      typeConfidence: confidence.typeConfidence,
      roleConfidence: confidence.roleConfidence,
      roleInferenceVersion: NUMERIC_MEASURE_ROLE_VERSION,
    },
  };
}

function countByRole(columns = []) {
  return columns.reduce((acc, column) => {
    const role = column.role || "unknown";
    acc[role] = (acc[role] || 0) + 1;
    return acc;
  }, {});
}

function hasRole(columns = [], role = "") {
  return columns.some((column) => column.role === role);
}

function countRoles(columns = [], roles = []) {
  return columns.filter((column) => roles.includes(column.role)).length;
}

function readyState(ready, reasons = []) {
  return {
    ready: Boolean(ready),
    reasons,
  };
}

function buildAnalysisReadiness(columns = []) {
  const hasMetric = hasRole(columns, "metric");
  const hasDate = hasRole(columns, "date");
  const hasDimension = hasRole(columns, "dimension");
  const hasStatus = hasRole(columns, "status");
  const dimensionLikeCount = countRoles(columns, ["dimension", "status"]);

  return {
    groupSummary: readyState(hasMetric && hasDimension, [
      hasMetric ? "HAS_METRIC" : "MISSING_METRIC",
      hasDimension ? "HAS_DIMENSION" : "MISSING_DIMENSION",
    ]),
    timeTrend: readyState(hasMetric && hasDate, [
      hasMetric ? "HAS_METRIC" : "MISSING_METRIC",
      hasDate ? "HAS_DATE" : "MISSING_DATE",
    ]),
    categoryCount: readyState(hasDimension || hasStatus, [
      hasDimension || hasStatus
        ? "HAS_CATEGORY_OR_STATUS"
        : "MISSING_CATEGORY_OR_STATUS",
    ]),
    topBottom: readyState(hasMetric && dimensionLikeCount >= 1, [
      hasMetric ? "HAS_METRIC" : "MISSING_METRIC",
      dimensionLikeCount >= 1 ? "HAS_LABEL" : "MISSING_LABEL",
    ]),
  };
}

function readinessKeys(readiness = {}) {
  return Object.entries(readiness)
    .filter(([, value]) => value?.ready)
    .map(([key]) => key);
}

function buildStructureSignals(columns = [], readiness = {}) {
  const roleCounts = countByRole(columns);
  const ready = readinessKeys(readiness);

  return {
    periodMetric: Boolean(readiness.timeTrend?.ready),
    categorySummary: Boolean(readiness.groupSummary?.ready),
    rosterStatus: Boolean(
      (roleCounts.dimension || 0) >= 2 ||
      ((roleCounts.dimension || 0) >= 1 && (roleCounts.status || 0) >= 1),
    ),
    analysisRecipeCount: ready.length,
    supportedAnalysisTypes: ready,
  };
}

function buildQueryabilityGrade({
  rows = [],
  columns = [],
  warnings = [],
  confidence = 0,
  readiness = {},
}) {
  const reasons = [];
  const supported = readinessKeys(readiness);

  if (!rows.length) reasons.push("NO_DATA_ROWS");
  if (columns.length <= 1) reasons.push("TOO_FEW_COLUMNS");
  if (warnings.includes("LOW_CONFIDENCE_HEADER"))
    reasons.push("LOW_CONFIDENCE_HEADER");
  if (warnings.includes("LOW_TYPE_CONSISTENCY"))
    reasons.push("LOW_TYPE_CONSISTENCY");
  if (!supported.length) reasons.push("NO_ANALYSIS_READY_RECIPE");

  if (!rows.length || columns.length <= 1 || Number(confidence || 0) < 0.35) {
    return { grade: "Q0", reasons };
  }

  if (warnings.includes("LOW_CONFIDENCE_HEADER") || !supported.length) {
    return { grade: "Q1", reasons };
  }

  if (Number(confidence || 0) >= 0.72 && supported.length >= 2) {
    return {
      grade: "Q3",
      reasons: reasons.length
        ? reasons
        : ["HIGH_CONFIDENCE_MULTI_RECIPE_READY"],
    };
  }

  return {
    grade: "Q2",
    reasons: reasons.length ? reasons : ["QUERYABLE_ANALYSIS_READY"],
  };
}

function buildDiagnostics({
  table = {},
  rows = [],
  columns = [],
  warnings = [],
  emptyRatio = 0,
  headerConfidence = 0,
  typeConsistency = 0,
  confidence = 0,
}) {
  const roleCounts = countByRole(columns);
  const readiness = buildAnalysisReadiness(columns);
  const structureSignals = buildStructureSignals(columns, readiness);
  const grade = buildQueryabilityGrade({
    rows,
    columns,
    warnings,
    confidence,
    readiness,
  });
  const excludedRows = Array.isArray(table.excludedRows)
    ? table.excludedRows
    : [];

  return {
    version: DIAGNOSTICS_VERSION,
    queryabilityGrade: grade.grade,
    queryabilityReasons: grade.reasons,
    metrics: {
      rowCount: rows.length,
      columnCount: columns.length,
      emptyRatio: Number(emptyRatio.toFixed(3)),
      headerConfidence: Number(headerConfidence.toFixed(3)),
      typeConsistency: Number(typeConsistency.toFixed(3)),
      confidence,
      excludedRowCount: excludedRows.length,
      summaryRowCount: Array.isArray(table.summaryRows)
        ? table.summaryRows.length
        : 0,
      isFallback: Boolean(table.isFallback),
      source: table.source || (table.isFallback ? "fallback" : "tableBlock"),
      isVirtual: Boolean(table.isVirtual),
      transformationType: table.transformation?.type || null,
      tableUsage: normalizeTableUsage(table),
    },
    transformation: table.transformation || null,
    tableUsage: normalizeTableUsage(table),
    roleCounts,
    analysisReadiness: readiness,
    structureSignals,
    inheritedQuality: {
      tableBlockScore: table.score ?? table.blockScore ?? null,
      dataQuality: table.dataQuality || null,
      headerQuality: table.headerQuality || null,
    },
  };
}

function findWideTemporalMetricColumns(table = {}) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];

  return columns
    .map((column) => ({
      column,
      temporal: parseTemporalHeader(
        column.header || column.originalHeader || "",
      ),
      numericRatio: numericRatioForColumn(rows, column),
    }))
    .filter(
      (item) =>
        item.temporal &&
        item.numericRatio >= 0.6 &&
        !["id", "status"].includes(String(item.column.role || "")),
    );
}

function findTemporalContextColumns(table = {}, wideColumns = []) {
  const wideSet = new Set(wideColumns.map((item) => item.column));
  return (table.columns || []).filter((column) => {
    if (wideSet.has(column)) return false;
    const header = column.header || column.originalHeader || "";
    const role = String(column.role || "");
    const parsed = parseTemporalHeader(header);
    return (
      role === "date" ||
      Boolean(parsed) ||
      /연도|년도|년월|연월|날짜|일자|date|year|period/i.test(header)
    );
  });
}

function selectWideDimensionColumns(
  table = {},
  wideColumns = [],
  contextColumns = [],
) {
  const wideSet = new Set(wideColumns.map((item) => item.column));
  const contextSet = new Set(contextColumns);
  const columns = Array.isArray(table.columns) ? table.columns : [];

  const preferred = columns.filter((column) => {
    if (wideSet.has(column) || contextSet.has(column)) return false;
    const role = String(column.role || "");
    if (["metric"].includes(role)) return false;
    return (
      ["dimension", "status", "id"].includes(role) ||
      ["string", "category", "boolean"].includes(String(column.type || ""))
    );
  });

  if (preferred.length) return preferred.slice(0, 4);

  return columns
    .filter((column) => !wideSet.has(column) && !contextSet.has(column))
    .slice(0, 3);
}

function mergeTemporalContext(headerTemporal = {}, rowTemporal = null) {
  const merged = {
    ...headerTemporal,
  };

  if (rowTemporal?.year && !merged.year) merged.year = rowTemporal.year;
  if (rowTemporal?.month && !merged.month) merged.month = rowTemporal.month;
  if (rowTemporal?.quarter && !merged.quarter)
    merged.quarter = rowTemporal.quarter;

  if (merged.year && merged.month)
    merged.periodLabel = `${merged.year}-${merged.month}`;
  else if (merged.year && merged.quarter)
    merged.periodLabel = `${merged.year}-${merged.quarter}`;
  else if (merged.year) merged.periodLabel = merged.year;
  else if (merged.month) merged.periodLabel = `${merged.month}월`;
  else if (merged.quarter) merged.periodLabel = merged.quarter;

  return merged;
}

function getRowTemporalContext(row = {}, contextColumns = []) {
  for (const column of contextColumns) {
    const value = getRowValueByColumn(row, column, column.index ?? 0);
    const parsed = parseTemporalValue(value);
    if (parsed) return parsed;
  }
  return null;
}

function buildWideToLongVirtualTable(table = {}, index = 0) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];
  if (!rows.length || columns.length < 3) return null;
  if (table?.transformation?.type === "wideToLong") return null;

  const wideColumns = findWideTemporalMetricColumns(table);
  if (wideColumns.length < 2) return null;

  const contextColumns = findTemporalContextColumns(table, wideColumns);
  const dimensionColumns = selectWideDimensionColumns(
    table,
    wideColumns,
    contextColumns,
  );

  if (!dimensionColumns.length && !contextColumns.length) return null;

  const usedHeaders = new Set();
  const dimensionSpecs = dimensionColumns.map((column) => ({
    column,
    outputHeader: uniqueHeader(
      column.header || column.originalHeader || "항목",
      usedHeaders,
    ),
  }));

  const periodHeader = uniqueHeader("기간", usedHeaders);
  const yearHeader = uniqueHeader("연도", usedHeaders);
  const monthHeader = uniqueHeader("월", usedHeaders);
  const quarterHeader = uniqueHeader("분기", usedHeaders);
  const metricNameHeader = uniqueHeader("지표명", usedHeaders);
  const metricValueHeader = uniqueHeader("지표값", usedHeaders);

  const outputRows = [];
  const metricLabels = new Set();

  for (const row of rows) {
    const rowTemporal = getRowTemporalContext(row, contextColumns);

    for (const item of wideColumns) {
      const rawValue = getRowValueByColumn(
        row,
        item.column,
        item.column.index ?? 0,
      );
      const value = toNumberOrNull(rawValue);
      if (value == null) continue;

      const temporal = mergeTemporalContext(item.temporal, rowTemporal);
      const metricLabel = temporal.metricLabel || "지표값";
      metricLabels.add(metricLabel);

      const out = {};
      for (const spec of dimensionSpecs) {
        out[spec.outputHeader] =
          getRowValueByColumn(row, spec.column, spec.column.index ?? 0) ?? "";
      }

      out[periodHeader] = temporal.periodLabel || item.temporal.raw;
      if (temporal.year) out[yearHeader] = temporal.year;
      if (temporal.month) out[monthHeader] = temporal.month;
      if (temporal.quarter) out[quarterHeader] = temporal.quarter;
      out[metricNameHeader] = metricLabel;
      out[metricValueHeader] = value;

      outputRows.push(out);
    }
  }

  if (!outputRows.length) return null;

  const virtualColumns = [
    ...dimensionSpecs.map((spec) =>
      makeVirtualColumn({
        header: spec.outputHeader,
        type: spec.column.type === "boolean" ? "boolean" : "category",
        role: spec.column.role === "status" ? "status" : "dimension",
        sourceColumn: spec.column,
      }),
    ),
    makeVirtualColumn({ header: periodHeader, type: "date", role: "date" }),
    makeVirtualColumn({ header: yearHeader, type: "string", role: "date" }),
    makeVirtualColumn({ header: monthHeader, type: "category", role: "date" }),
    makeVirtualColumn({
      header: quarterHeader,
      type: "category",
      role: "date",
    }),
    makeVirtualColumn({
      header: metricNameHeader,
      type: "category",
      role: "dimension",
    }),
    makeVirtualColumn({
      header: metricValueHeader,
      type: "number",
      role: "metric",
    }),
  ].filter((column) =>
    outputRows.some((row) =>
      Object.prototype.hasOwnProperty.call(row, column.header),
    ),
  );

  return {
    tableId: `${table.tableId || `table_${index + 1}`}#WIDE_LONG`,
    sourceTableId: table.tableId || null,
    sheetName: table.sheetName || "",
    tableType: "wide_to_long",
    source: "normalizedWideToLong",
    isVirtual: true,
    range: table.range || null,
    dataRange: table.dataRange || null,
    headerRows: table.headerRows || [],
    dataStartRow: table.dataStartRow ?? null,
    columns: virtualColumns,
    rows: outputRows,
    excludedRows: [],
    warnings: [],
    confidence: Math.min(
      0.92,
      Math.max(0.68, Number(table.confidence || 0.72)),
    ),
    tableUsage: inheritVirtualTableUsage(table, "wide_to_long"),
    transformation: {
      version: "wide_to_long_normalization_v1",
      type: "wideToLong",
      sourceTableId: table.tableId || null,
      sourceColumnCount: columns.length,
      sourceRowCount: rows.length,
      generatedRowCount: outputRows.length,
      dimensionColumns: dimensionSpecs.map((spec) => spec.column.header),
      temporalMetricColumns: wideColumns.map((item) => ({
        header: item.column.header,
        periodLabel: item.temporal.periodLabel,
        numericRatio: Number(item.numericRatio.toFixed(3)),
      })),
      metricLabels: Array.from(metricLabels).slice(0, 20),
    },
  };
}

function buildWideToLongVirtualTables(normalizedTables = []) {
  const out = [];

  for (let i = 0; i < normalizedTables.length; i += 1) {
    if (!isAnalysisEligibleTable(normalizedTables[i])) continue;
    const virtualTable = buildWideToLongVirtualTable(normalizedTables[i], i);
    if (virtualTable)
      out.push(
        normalizeTable(virtualTable, normalizedTables.length + out.length),
      );
  }

  return out;
}

function isCrossMetricHeaderCandidate(header = "") {
  const value = normalizeHeader(header);
  if (!value) return false;
  if (parseTemporalHeader(value)) return false;

  // 문장/주석형 긴 헤더는 cross axis라기보다 설명 컬럼일 가능성이 높다.
  if (value.length > 40) return false;
  if (/[.!?。]|입니다|합니다/.test(value)) return false;

  return true;
}

function findCrossMetricColumns(table = {}) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];

  return columns
    .map((column) => {
      const header = column.header || column.originalHeader || "";
      const role = String(column.role || "");
      const type = String(column.type || "");
      const numericRatio = numericRatioForColumn(rows, column);
      return {
        column,
        header,
        role,
        type,
        numericRatio,
      };
    })
    .filter((item) => {
      if (!isCrossMetricHeaderCandidate(item.header)) return false;
      if (["id", "status", "date"].includes(item.role)) return false;
      if (parseTemporalHeader(item.header)) return false;

      // cross table의 값 영역은 숫자 비율이 높아야 한다.
      // role/type은 참고만 하고 실제 값이 숫자로 읽혀야 변환한다.
      return item.numericRatio >= 0.6;
    });
}

function findCrossDimensionColumns(table = {}, crossColumns = []) {
  const crossSet = new Set(crossColumns.map((item) => item.column));
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];

  const preferred = columns.filter((column) => {
    if (crossSet.has(column)) return false;
    const header = column.header || column.originalHeader || "";
    const role = String(column.role || "");
    const type = String(column.type || "");
    const numericRatio = numericRatioForColumn(rows, column);

    if (parseTemporalHeader(header)) return false;
    if (role === "metric" || type === "number") return false;
    if (numericRatio >= 0.5) return false;

    return (
      ["dimension", "status", "id"].includes(role) ||
      ["string", "category", "boolean", "text"].includes(type)
    );
  });

  if (preferred.length) return preferred.slice(0, 5);

  // 역할 추론이 약한 테이블을 위한 보수적 fallback.
  return columns
    .filter((column) => !crossSet.has(column))
    .filter((column) => numericRatioForColumn(rows, column) < 0.5)
    .slice(0, 3);
}

function isLikelyCrossTable(
  table = {},
  crossColumns = [],
  dimensionColumns = [],
) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];

  if (!rows.length || columns.length < 3) return false;
  if (crossColumns.length < 2 || !dimensionColumns.length) return false;

  const crossRatio = crossColumns.length / columns.length;
  const allCrossHeadersShort = crossColumns.every(
    (item) => String(item.header || "").trim().length <= 40,
  );

  // 일반 테이블의 숫자 지표 2개를 모두 long 변환하는 오탐을 줄이기 위한 최소 구조 조건.
  // - 숫자 값 영역이 전체 컬럼의 일정 비율 이상이거나
  // - dimension 1~2개 + 값 컬럼 다수인 행렬형 구조여야 한다.
  return (
    allCrossHeadersShort && (crossRatio >= 0.4 || crossColumns.length >= 3)
  );
}

function buildCrossTableToLongVirtualTable(table = {}, index = 0) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];
  if (!rows.length || columns.length < 3) return null;
  if (table?.transformation?.type) return null;

  const crossColumns = findCrossMetricColumns(table);
  const dimensionColumns = findCrossDimensionColumns(table, crossColumns);

  if (!isLikelyCrossTable(table, crossColumns, dimensionColumns)) return null;

  const usedHeaders = new Set();
  const dimensionSpecs = dimensionColumns.map((column) => ({
    column,
    outputHeader: uniqueHeader(
      column.header || column.originalHeader || "행 항목",
      usedHeaders,
    ),
  }));

  const crossAxisHeader = uniqueHeader("교차항목", usedHeaders);
  const metricNameHeader = uniqueHeader("지표명", usedHeaders);
  const metricValueHeader = uniqueHeader("지표값", usedHeaders);

  const outputRows = [];

  for (const row of rows) {
    for (const item of crossColumns) {
      const rawValue = getRowValueByColumn(
        row,
        item.column,
        item.column.index ?? 0,
      );
      const value = toNumberOrNull(rawValue);
      if (value == null) continue;

      const out = {};
      for (const spec of dimensionSpecs) {
        out[spec.outputHeader] =
          getRowValueByColumn(row, spec.column, spec.column.index ?? 0) ?? "";
      }

      out[crossAxisHeader] = item.header || item.column.header || "항목";
      out[metricNameHeader] = "지표값";
      out[metricValueHeader] = value;
      outputRows.push(out);
    }
  }

  if (!outputRows.length) return null;

  const virtualColumns = [
    ...dimensionSpecs.map((spec) =>
      makeVirtualColumn({
        header: spec.outputHeader,
        type: spec.column.type === "boolean" ? "boolean" : "category",
        role: spec.column.role === "status" ? "status" : "dimension",
        sourceColumn: spec.column,
      }),
    ),
    makeVirtualColumn({
      header: crossAxisHeader,
      type: "category",
      role: "dimension",
    }),
    makeVirtualColumn({
      header: metricNameHeader,
      type: "category",
      role: "dimension",
    }),
    makeVirtualColumn({
      header: metricValueHeader,
      type: "number",
      role: "metric",
    }),
  ];

  return {
    tableId: `${table.tableId || `table_${index + 1}`}#CROSS_LONG`,
    sourceTableId: table.tableId || null,
    sheetName: table.sheetName || "",
    tableType: "cross_table_long",
    source: "normalizedCrossTableToLong",
    isVirtual: true,
    range: table.range || null,
    dataRange: table.dataRange || null,
    headerRows: table.headerRows || [],
    dataStartRow: table.dataStartRow ?? null,
    columns: virtualColumns,
    rows: outputRows,
    excludedRows: [],
    warnings: [],
    confidence: Math.min(0.9, Math.max(0.65, Number(table.confidence || 0.7))),
    tableUsage: inheritVirtualTableUsage(table, "cross_table_to_long"),
    transformation: {
      version: "cross_table_to_long_normalization_v1",
      type: "crossTableToLong",
      sourceTableId: table.tableId || null,
      sourceColumnCount: columns.length,
      sourceRowCount: rows.length,
      generatedRowCount: outputRows.length,
      dimensionColumns: dimensionSpecs.map((spec) => spec.column.header),
      crossMetricColumns: crossColumns.map((item) => ({
        header: item.header,
        numericRatio: Number(item.numericRatio.toFixed(3)),
      })),
      outputHeaders: {
        crossAxis: crossAxisHeader,
        metricName: metricNameHeader,
        metricValue: metricValueHeader,
      },
    },
  };
}

function buildCrossTableToLongVirtualTables(normalizedTables = []) {
  const out = [];

  for (let i = 0; i < normalizedTables.length; i += 1) {
    if (!isAnalysisEligibleTable(normalizedTables[i])) continue;
    const virtualTable = buildCrossTableToLongVirtualTable(
      normalizedTables[i],
      i,
    );
    if (virtualTable)
      out.push(
        normalizeTable(virtualTable, normalizedTables.length + out.length),
      );
  }

  return out;
}

function normalizeTable(table = {}, index = 0) {
  const columns = Array.isArray(table.columns)
    ? table.columns.map((column, columnIndex) =>
        normalizeColumn(column, columnIndex, table),
      )
    : [];

  const rows = Array.isArray(table.rows) ? table.rows : [];

  const emptyRatio = calculateEmptyRatio(rows);

  const headerConfidence = calculateHeaderConfidence(columns);

  const typeConsistency = calculateTypeConsistency(columns);

  const warnings = buildWarnings({
    rows,
    columns,
    emptyRatio,
    headerConfidence,
    typeConsistency,
  });

  const confidence = calculateConfidence({
    emptyRatio,
    headerConfidence,
    typeConsistency,
  });

  const diagnostics = buildDiagnostics({
    table,
    rows,
    columns,
    warnings,
    emptyRatio,
    headerConfidence,
    typeConsistency,
    confidence,
  });

  return {
    tableId: table.tableId || table.id || `table_${index + 1}`,
    sheetName: table.sheetName || table.sheet || "",
    tableType: table.tableType || "tabular",
    source: table.source || null,
    sourceTableId: table.sourceTableId || null,
    isVirtual: Boolean(table.isVirtual),
    transformation: table.transformation || null,
    tableUsage: normalizeTableUsage(table),
    headerRows: table.headerRows || [],
    dataStartRow: table.dataStartRow ?? null,
    range: table.range || null,
    dataRange: table.dataRange || null,
    columns,
    rows,
    excludedRows: Array.isArray(table.excludedRows) ? table.excludedRows : [],
    summaryRows: Array.isArray(table.summaryRows) ? table.summaryRows : [],
    dataQuality: table.dataQuality || null,
    warnings,
    confidence,
    queryabilityGrade: diagnostics.queryabilityGrade,
    queryabilityReasons: diagnostics.queryabilityReasons,
    diagnostics,
  };
}

function buildNormalizedQueryTables(queryTables = []) {
  if (!Array.isArray(queryTables)) return [];

  const normalizedTables = queryTables.map((table, index) =>
    normalizeTable(table, index),
  );
  const wideLongTables = buildWideToLongVirtualTables(normalizedTables);
  const crossLongTables = buildCrossTableToLongVirtualTables(normalizedTables);

  return [...normalizedTables, ...wideLongTables, ...crossLongTables];
}

module.exports = {
  buildNormalizedQueryTables,
  buildWideToLongVirtualTable,
  buildCrossTableToLongVirtualTable,
};
