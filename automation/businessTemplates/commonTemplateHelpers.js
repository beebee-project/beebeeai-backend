const { executeAnalysisRecipeCandidate } = require("../analysisRecipeExecutor");

function normalizeHeader(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function normalizeText(value = "") {
  return String(value ?? "").trim();
}

function headerMatches(header = "", hints = []) {
  const h = normalizeHeader(header);

  if (!h) return false;

  return hints.some((hint) => {
    const normalizedHint = normalizeHeader(hint);
    return (
      normalizedHint &&
      (h.includes(normalizedHint) || normalizedHint.includes(h))
    );
  });
}

function getColumnHeader(column = {}) {
  return (
    column.header ||
    column.originalHeader ||
    column.name ||
    column.key ||
    column.accessor ||
    ""
  );
}

function getColumns(table = {}) {
  return Array.isArray(table.columns) ? table.columns : [];
}

function findTableForTemplate(
  normalizedQueryTables = [],
  templateCandidate = {},
) {
  const firstCandidate = Array.isArray(templateCandidate.candidates)
    ? templateCandidate.candidates[0]
    : null;

  if (firstCandidate?.tableId) {
    const matched = normalizedQueryTables.find(
      (table) => table.tableId === firstCandidate.tableId,
    );

    if (matched) return matched;
  }

  return (
    normalizedQueryTables.find((table) => table.isPrimary) ||
    normalizedQueryTables[0] ||
    null
  );
}

function findColumn(table = {}, hints = [], options = {}) {
  const columns = getColumns(table);

  const matchedByHint = columns.find((col) =>
    headerMatches(getColumnHeader(col), hints),
  );

  if (matchedByHint) return matchedByHint;

  if (options.type) {
    const matchedByType = columns.find(
      (col) =>
        col.type === options.type ||
        col.dominantType === options.type ||
        col.semanticType === options.type,
    );

    if (matchedByType) return matchedByType;
  }

  if (options.role) {
    const matchedByRole = columns.find(
      (col) =>
        col.role === options.role ||
        col.inferredRole === options.role ||
        col.semanticRole === options.role,
    );

    if (matchedByRole) return matchedByRole;
  }

  return null;
}

function findColumnHeader(table = {}, hints = [], options = {}) {
  const column = findColumn(table, hints, options);
  return column ? getColumnHeader(column) : "";
}

function findNumericHeaders(table = {}, hints = []) {
  const columns = getColumns(table);

  return columns
    .filter((col) => {
      const header = getColumnHeader(col);

      const isNumber =
        col.type === "number" ||
        col.dominantType === "number" ||
        col.role === "metric" ||
        col.inferredRole === "metric";

      const matchedHint = hints.length ? headerMatches(header, hints) : true;

      return isNumber && matchedHint;
    })
    .map(getColumnHeader)
    .filter(Boolean);
}

function makeTemplateCandidate({
  recipeType,
  title,
  tableId,
  columns = {},
  meta = {},
}) {
  return {
    recipeType,
    title,
    tableId,
    columns,
    meta,
  };
}

function executeTemplateSections({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const candidates = Array.isArray(templateCandidate.candidates)
    ? templateCandidate.candidates
    : [];

  return candidates
    .map((candidate, index) => {
      const result = executeAnalysisRecipeCandidate({
        normalizedQueryTables,
        candidate,
      });

      if (!result?.ok) return null;

      return {
        sectionId:
          candidate.sectionId ||
          candidate.recipeType ||
          candidate.type ||
          candidate.recipeId ||
          `section_${index + 1}`,
        title:
          candidate.title ||
          candidate.name ||
          candidate.label ||
          `섹션 ${index + 1}`,
        candidate,
        result,
      };
    })
    .filter(Boolean);
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  if (value == null || value === "") return null;

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");

  if (!cleaned || cleaned === "-" || cleaned === "." || cleaned === "-.") {
    return null;
  }

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function getRowValue(row = {}, header = "") {
  if (!row || !header) return undefined;

  if (Object.prototype.hasOwnProperty.call(row, header)) {
    return row[header];
  }

  const normalizedHeader = normalizeHeader(header);

  const matchedKey = Object.keys(row).find(
    (key) => normalizeHeader(key) === normalizedHeader,
  );

  return matchedKey ? row[matchedKey] : undefined;
}

function getRows(table = {}) {
  return Array.isArray(table.rows) ? table.rows : [];
}

function createVirtualTable({
  sourceTable = {},
  tableId = "",
  tableName = "",
  columns = [],
  rows = [],
}) {
  return {
    ...sourceTable,
    tableId: tableId || `${sourceTable.tableId || "table"}_virtual`,
    tableName:
      tableName || sourceTable.tableName || sourceTable.sheetName || "",
    columns,
    rows,
    rowCount: rows.length,
    isVirtual: true,
  };
}

module.exports = {
  normalizeHeader,
  normalizeText,
  headerMatches,
  getColumnHeader,
  getColumns,
  findTableForTemplate,
  findColumn,
  findColumnHeader,
  findNumericHeaders,
  makeTemplateCandidate,
  executeTemplateSections,
  toNumber,
  getRowValue,
  getRows,
  createVirtualTable,
};
