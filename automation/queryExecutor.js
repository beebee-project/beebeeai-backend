function toComparable(value) {
  if (value == null || value === "") return null;

  if (typeof value === "number") return value;

  const n = Number(String(value).replace(/,/g, "").trim());
  if (Number.isFinite(n)) return n;

  return String(value).trim();
}

function matchFilter(row, filter = {}) {
  const raw = row[filter.columnKey];
  const left = toComparable(raw);
  const right = toComparable(filter.value);

  switch (filter.operator) {
    case "=":
      return String(left ?? "").trim() === String(right ?? "").trim();
    case "!=":
      return String(left ?? "").trim() !== String(right ?? "").trim();
    case ">":
      return Number(left) > Number(right);
    case ">=":
      return Number(left) >= Number(right);
    case "<":
      return Number(left) < Number(right);
    case "<=":
      return Number(left) <= Number(right);
    default:
      return true;
  }
}

function applyFilters(rows = [], filters = []) {
  if (!filters.length) return rows;
  return rows.filter((row) => filters.every((f) => matchFilter(row, f)));
}

function numericValues(rows = [], columnKey) {
  return rows
    .map((r) => toComparable(r[columnKey]))
    .filter((v) => typeof v === "number" && Number.isFinite(v));
}

function aggregate(rows = [], operation = "list", metric = null) {
  if (operation === "count") return rows.length;

  if (!metric?.columnKey) return rows;

  const values = numericValues(rows, metric.columnKey);

  if (operation === "sum") {
    return values.reduce((a, b) => a + b, 0);
  }

  if (operation === "average") {
    if (!values.length) return null;
    return values.reduce((a, b) => a + b, 0) / values.length;
  }

  if (operation === "max") {
    if (!values.length) return null;
    return Math.max(...values);
  }

  if (operation === "min") {
    if (!values.length) return null;
    return Math.min(...values);
  }

  return rows;
}

function groupRows(rows = [], groupBy) {
  const groups = new Map();

  for (const row of rows) {
    const key = String(row[groupBy.columnKey] ?? "");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }

  return groups;
}

function executeQueryIntent(queryTables = [], intent = {}) {
  const table =
    queryTables.find((t) => t.tableId === intent.table?.tableId) ||
    queryTables.find((t) => t.isPrimary) ||
    queryTables[0];

  if (!table) {
    return {
      ok: false,
      error: "실행 가능한 테이블이 없습니다.",
    };
  }

  const rows = Array.isArray(table.rows) ? table.rows : [];
  const filteredRows = applyFilters(rows, intent.filters || []);

  if (intent.groupBy?.columnKey) {
    const groups = groupRows(filteredRows, intent.groupBy);
    const resultRows = [];

    for (const [groupValue, groupRowsValue] of groups.entries()) {
      resultRows.push({
        [intent.groupBy.header || intent.groupBy.columnKey]: groupValue,
        operation: intent.operation,
        metric: intent.metric?.header || null,
        value: aggregate(groupRowsValue, intent.operation, intent.metric),
        rowCount: groupRowsValue.length,
      });
    }

    return {
      ok: true,
      table: intent.table,
      operation: intent.operation,
      metric: intent.metric,
      groupBy: intent.groupBy,
      filters: intent.filters || [],
      rowCount: filteredRows.length,
      resultType: "grouped",
      rows: resultRows,
    };
  }

  const value = aggregate(filteredRows, intent.operation, intent.metric);

  return {
    ok: true,
    table: intent.table,
    operation: intent.operation,
    metric: intent.metric,
    groupBy: null,
    filters: intent.filters || [],
    rowCount: filteredRows.length,
    resultType: intent.operation === "list" ? "rows" : "scalar",
    value,
    rows: intent.operation === "list" ? filteredRows.slice(0, 100) : [],
  };
}

module.exports = { executeQueryIntent };
