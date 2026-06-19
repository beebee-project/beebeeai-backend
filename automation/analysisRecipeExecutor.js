const { ANALYSIS_OUTPUT_LABELS } = require("./analysisOutputLabelConfig");

function normalizeText(value = "") {
  return String(value || "").trim();
}

function getValue(row = {}, key = "") {
  if (!row || !key) return undefined;
  return row[key];
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

function normalizeHeader(value = "") {
  return String(value)
    .toLowerCase()
    .replace(/[\s_]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function resolveCandidateColumn(candidate = {}, key = "") {
  const pluralMap = {
    metric: "metrics",
    dimension: "dimensions",
    date: "dates",
  };

  return (
    candidate.columns?.[key] ||
    candidate[key] ||
    candidate[`${key}Header`] ||
    candidate[pluralMap[key]]?.[0] ||
    ""
  );
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

function groupSummary({ rows = [], dimension = "", metric = "" }) {
  const map = new Map();

  rows.forEach((row) => {
    const label =
      normalizeText(getRowValue(row, dimension)) ||
      ANALYSIS_OUTPUT_LABELS.emptyLabel;

    const rawValue = getRowValue(row, metric);
    const value = toNumber(rawValue);

    if (!map.has(label)) {
      map.set(label, {
        [dimension]: label,
        count: 0,
        sum: 0,
        average: null,
      });
    }

    const item = map.get(label);
    item.count += 1;

    if (value != null) {
      item.sum += value;
    }
  });

  return Array.from(map.values()).map((item) => ({
    ...item,
    average: item.count ? item.sum / item.count : null,
  }));
}

function categoryCount({ rows = [], dimension = "" }) {
  const map = new Map();

  rows.forEach((row) => {
    const label =
      normalizeText(getRowValue(row, dimension)) ||
      ANALYSIS_OUTPUT_LABELS.emptyLabel;
    map.set(label, (map.get(label) || 0) + 1);
  });

  return Array.from(map.entries()).map(([label, count]) => ({
    [dimension]: label,
    count,
  }));
}

function topBottom({ rows = [], metric = "", dimension = "" }) {
  const values = rows
    .map((row) => ({
      label:
        normalizeText(getRowValue(row, dimension)) ||
        ANALYSIS_OUTPUT_LABELS.itemLabel,
      value: toNumber(getRowValue(row, metric)),
      row,
    }))
    .filter((item) => item.value != null)
    .sort((a, b) => b.value - a.value);

  return [
    ...values.slice(0, 5).map((item) => ({
      type: ANALYSIS_OUTPUT_LABELS.top,
      label: item.label,
      value: item.value,
    })),
    ...values
      .slice(-5)
      .reverse()
      .map((item) => ({
        type: ANALYSIS_OUTPUT_LABELS.bottom,
        label: item.label,
        value: item.value,
      })),
  ];
}

function normalizeMonth(value) {
  const text = normalizeText(value);
  if (!text) return null;

  const date = value instanceof Date ? value : new Date(text);
  if (Number.isNaN(date.getTime())) return null;

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");

  return `${year}-${month}`;
}

function timeTrend({ rows = [], date = "", metric = "" }) {
  const map = new Map();

  rows.forEach((row) => {
    const month = normalizeMonth(getRowValue(row, date));
    const value = toNumber(getRowValue(row, metric));

    if (!month || value == null) return;

    if (!map.has(month)) {
      map.set(month, {
        period: month,
        count: 0,
        sum: 0,
        average: null,
      });
    }

    const item = map.get(month);
    item.count += 1;
    item.sum += value;
  });

  return Array.from(map.values())
    .sort((a, b) => String(a.period).localeCompare(String(b.period)))
    .map((item) => ({
      ...item,
      average: item.count ? item.sum / item.count : null,
    }));
}

function findTable(normalizedQueryTables = [], tableId = "") {
  return normalizedQueryTables.find((table) => table.tableId === tableId);
}

function executeAnalysisRecipeCandidate({
  normalizedQueryTables = [],
  candidate = {},
}) {
  const table = findTable(normalizedQueryTables, candidate.tableId);

  if (!table) {
    return {
      ok: false,
      code: "TABLE_NOT_FOUND",
      message: "선택한 분석 후보의 표를 찾지 못했습니다.",
    };
  }

  const rows = Array.isArray(table.rows) ? table.rows : [];

  const metric = resolveCandidateColumn(candidate, "metric");
  const dimension = resolveCandidateColumn(candidate, "dimension");
  const date = resolveCandidateColumn(candidate, "date");

  let resultRows = [];

  if (candidate.recipeType === "group_summary") {
    resultRows = groupSummary({ rows, dimension, metric });
  } else if (candidate.recipeType === "category_count") {
    resultRows = categoryCount({ rows, dimension });
  } else if (candidate.recipeType === "top_bottom") {
    resultRows = topBottom({ rows, dimension, metric });
  } else if (candidate.recipeType === "time_trend") {
    resultRows = timeTrend({ rows, date, metric });
  } else {
    return {
      ok: false,
      code: "UNSUPPORTED_RECIPE_TYPE",
      message: "지원하지 않는 분석 후보입니다.",
    };
  }

  return {
    ok: true,
    recipeType: candidate.recipeType,
    title: candidate.title,
    tableId: table.tableId,
    sheetName: table.sheetName,
    columns: candidate.columns,
    rows: resultRows,
    rowCount: resultRows.length,
  };
}

module.exports = {
  executeAnalysisRecipeCandidate,
};
