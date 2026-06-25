const {
  COLUMN_ROLE_PATTERNS,
  BOOLEAN_VALUES,
  COLUMN_INFERENCE_THRESHOLDS,
} = require("./config/columnRoleConfig");

function isBlank(value) {
  return value == null || String(value).trim() === "";
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

  const text = value.trim();
  if (!text) return false;

  return (
    /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(text) ||
    /^\d{4}[-/.]\d{1,2}$/.test(text)
  );
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

function inferColumnRole(header = "", type = "unknown") {
  const text = String(header).toLowerCase();

  if (type === "date" || COLUMN_ROLE_PATTERNS.date.test(text)) return "date";
  if (COLUMN_ROLE_PATTERNS.id.test(text)) return "id";
  if (COLUMN_ROLE_PATTERNS.status.test(text)) return "status";
  if (type === "number" && COLUMN_ROLE_PATTERNS.metric.test(text))
    return "metric";
  if (type === "string" || type === "boolean") return "dimension";

  return "unknown";
}

function getColumnValues(table = {}, header = "", index = 0) {
  const rows = Array.isArray(table.rows) ? table.rows : [];

  return rows.map((row) => {
    if (row && Object.prototype.hasOwnProperty.call(row, header)) {
      return row[header];
    }

    if (Array.isArray(row)) {
      return row[index];
    }

    return undefined;
  });
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

function normalizeColumn(column = {}, index = 0, table = {}) {
  const header =
    column.header ||
    column.name ||
    column.key ||
    column.label ||
    `column_${index + 1}`;

  const values = getColumnValues(table, header, column.index ?? index);
  const inferredType =
    column.type || column.valueType || inferColumnType(values);
  const inferredRole = column.role || inferColumnRole(header, inferredType);

  return {
    header,
    originalHeader: column.originalHeader || header,
    index: column.index ?? index,
    type: inferredType,
    role: inferredRole,
    quality: Number.isFinite(Number(column.quality))
      ? Number(column.quality)
      : null,
  };
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

  return {
    tableId: table.tableId || table.id || `table_${index + 1}`,
    sheetName: table.sheetName || table.sheet || "",
    tableType: table.tableType || "tabular",
    headerRows: table.headerRows || [],
    dataStartRow: table.dataStartRow ?? null,
    range: table.range || null,
    columns,
    rows,
    excludedRows: Array.isArray(table.excludedRows) ? table.excludedRows : [],
    warnings,
    confidence,
  };
}

function buildNormalizedQueryTables(queryTables = []) {
  if (!Array.isArray(queryTables)) return [];

  return queryTables.map((table, index) => normalizeTable(table, index));
}

module.exports = {
  buildNormalizedQueryTables,
};
