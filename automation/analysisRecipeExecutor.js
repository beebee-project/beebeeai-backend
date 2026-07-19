const { ANALYSIS_OUTPUT_LABELS } = require("./analysisOutputLabelConfig");
const { ANALYSIS_RECIPE_TYPES } = require("./config/analysisRecipeConfig");

function normalizeText(value = "") {
  return String(value || "").trim();
}

function normalizeHeader(value = "") {
  return String(value)
    .toLowerCase()
    .replace(/[\s_]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
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

function resolveCandidateColumn(candidate = {}, key = "") {
  const pluralMap = {
    metric: "metrics",
    dimension: "dimensions",
    dimension2: "dimensions",
    date: "dates",
  };

  if (key === "dimension2") {
    return (
      candidate.columns?.dimension2 ||
      candidate.dimension2 ||
      candidate.dimension2Header ||
      candidate[pluralMap[key]]?.[1] ||
      ""
    );
  }

  return (
    candidate.columns?.[key] ||
    candidate[key] ||
    candidate[`${key}Header`] ||
    candidate[pluralMap[key]]?.[0] ||
    ""
  );
}

function comparableText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase();
}

function valuesEqual(left, right) {
  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (leftNumber != null && rightNumber != null) {
    return leftNumber === rightNumber;
  }
  return comparableText(left) === comparableText(right);
}

function matchesCandidateFilter(row = {}, filter = {}) {
  const header = filter.header || filter.column || filter.field || "";
  if (!header) return true;

  const operator = String(filter.operator || "equals").toLowerCase();
  const actual = getRowValue(row, header);
  const expected = filter.value;
  const expectedValues = Array.isArray(expected) ? expected : [expected];

  if (operator === "isblank")
    return actual == null || String(actual).trim() === "";
  if (operator === "notblank")
    return !(actual == null || String(actual).trim() === "");
  if (operator === "equals" || operator === "eq") {
    return expectedValues.some((value) => valuesEqual(actual, value));
  }
  if (operator === "notequals" || operator === "neq") {
    return expectedValues.every((value) => !valuesEqual(actual, value));
  }
  if (operator === "in") {
    return expectedValues.some((value) => valuesEqual(actual, value));
  }
  if (operator === "notin") {
    return expectedValues.every((value) => !valuesEqual(actual, value));
  }
  if (operator === "includes" || operator === "contains") {
    const actualText = comparableText(actual);
    return expectedValues.some((value) =>
      actualText.includes(comparableText(value)),
    );
  }

  const actualNumber = toNumber(actual);
  const expectedNumber = toNumber(expectedValues[0]);
  if (actualNumber == null || expectedNumber == null) return false;
  if (operator === "gt") return actualNumber > expectedNumber;
  if (operator === "gte") return actualNumber >= expectedNumber;
  if (operator === "lt") return actualNumber < expectedNumber;
  if (operator === "lte") return actualNumber <= expectedNumber;

  return false;
}

function applyCandidateFilters(rows = [], filters = []) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const safeFilters = Array.isArray(filters)
    ? filters.filter((filter) => filter && typeof filter === "object")
    : [];
  if (!safeFilters.length) return safeRows;

  return safeRows.filter((row) =>
    safeFilters.every((filter) => matchesCandidateFilter(row, filter)),
  );
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null || value === "") return null;

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/[^\d.-]/g, "");

  if (!cleaned || cleaned === "-" || cleaned === "." || cleaned === "-.") {
    return null;
  }

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function safeDivide(numerator, denominator) {
  const a = Number(numerator);
  const b = Number(denominator);
  if (!Number.isFinite(a) || !Number.isFinite(b) || b === 0) return null;
  return a / b;
}

function normalizePeriod(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  const text = normalizeText(value);
  if (!text) return null;

  const compact = text.replace(/\s+/g, "");

  let match = compact.match(/(\d{4})[.\-/년_]*(\d{1,2})/);
  if (match) {
    return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
  }

  match = compact.match(/(\d{4})\s*(?:년|년도)?$/);
  if (match) return match[1];

  const date = new Date(text);
  if (!Number.isNaN(date.getTime())) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  return text;
}

function sortByPeriod(rows = [], key = "period") {
  return rows
    .slice()
    .sort((a, b) =>
      String(a?.[key] || "").localeCompare(String(b?.[key] || "")),
    );
}

function aggregateRows({
  rows = [],
  dimension = "",
  metric = "",
  operation = "summary",
}) {
  const map = new Map();

  rows.forEach((row) => {
    const label =
      normalizeText(getRowValue(row, dimension)) ||
      ANALYSIS_OUTPUT_LABELS.emptyLabel;
    const value = metric ? toNumber(getRowValue(row, metric)) : null;

    if (!map.has(label)) {
      map.set(label, {
        [dimension]: label,
        count: 0,
        numericCount: 0,
        sum: 0,
        average: null,
      });
    }

    const item = map.get(label);
    item.count += 1;
    if (value != null) {
      item.numericCount += 1;
      item.sum += value;
    }
  });

  const result = Array.from(map.values()).map((item) => ({
    ...item,
    average: item.numericCount ? item.sum / item.numericCount : null,
  }));

  if (operation === "sum") {
    return result.map((item) => ({
      [dimension]: item[dimension],
      value: item.sum,
      sum: item.sum,
      count: item.count,
    }));
  }

  if (operation === "average") {
    return result.map((item) => ({
      [dimension]: item[dimension],
      value: item.average,
      average: item.average,
      count: item.count,
    }));
  }

  if (operation === "count") {
    return result.map((item) => ({
      [dimension]: item[dimension],
      value: item.count,
      count: item.count,
    }));
  }

  return result;
}

function categoryCount({ rows = [], dimension = "" }) {
  return aggregateRows({ rows, dimension, operation: "count" });
}

function topBottom({ rows = [], metric = "", dimension = "" }) {
  const values = rows
    .map((row, index) => ({
      label:
        normalizeText(getRowValue(row, dimension)) ||
        `${ANALYSIS_OUTPUT_LABELS.itemLabel}_${index + 1}`,
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

function timeAggregate({
  rows = [],
  date = "",
  metric = "",
  operation = "summary",
}) {
  const map = new Map();

  rows.forEach((row) => {
    const period = normalizePeriod(getRowValue(row, date));
    if (!period) return;

    const value = metric ? toNumber(getRowValue(row, metric)) : null;

    if (!map.has(period)) {
      map.set(period, {
        period,
        count: 0,
        numericCount: 0,
        sum: 0,
        average: null,
      });
    }

    const item = map.get(period);
    item.count += 1;
    if (value != null) {
      item.numericCount += 1;
      item.sum += value;
    }
  });

  const result = sortByPeriod(Array.from(map.values())).map((item) => ({
    ...item,
    average: item.numericCount ? item.sum / item.numericCount : null,
  }));

  if (operation === "sum") {
    return result.map((item) => ({ ...item, value: item.sum }));
  }
  if (operation === "average") {
    return result.map((item) => ({ ...item, value: item.average }));
  }
  if (operation === "count") {
    return result.map((item) => ({ ...item, value: item.count }));
  }

  return result;
}

function timeGrowth(args = {}) {
  const base = timeAggregate({ ...args, operation: "sum" });
  return base.map((item, index) => {
    const previous = index > 0 ? base[index - 1].value : null;
    const change = previous == null ? null : item.value - previous;
    const growthRate =
      previous == null ? null : safeDivide(change, previous) * 100;
    return {
      ...item,
      previous,
      change,
      growthRate,
    };
  });
}

function cumulativeSum(args = {}) {
  const base = timeAggregate({ ...args, operation: "sum" });
  let cumulative = 0;
  return base.map((item) => {
    cumulative += Number(item.value || 0);
    return {
      ...item,
      cumulative,
    };
  });
}

function compositionRatio({ rows = [], dimension = "", metric = "" }) {
  const base = metric
    ? aggregateRows({ rows, dimension, metric, operation: "sum" })
    : categoryCount({ rows, dimension });
  const total = base.reduce(
    (sum, row) => sum + Number(row.value || row.sum || row.count || 0),
    0,
  );

  return base
    .map((row) => {
      const value = Number(row.value || row.sum || row.count || 0);
      const ratio = total ? value / total : null;
      return {
        ...row,
        value,
        total,
        ratio,
        ratioPercent: ratio == null ? null : ratio * 100,
      };
    })
    .sort((a, b) => Number(b.value || 0) - Number(a.value || 0));
}

function crossAggregate({
  rows = [],
  dimension = "",
  dimension2 = "",
  metric = "",
  operation = "count",
}) {
  const map = new Map();

  rows.forEach((row) => {
    const label1 =
      normalizeText(getRowValue(row, dimension)) ||
      ANALYSIS_OUTPUT_LABELS.emptyLabel;
    const label2 =
      normalizeText(getRowValue(row, dimension2)) ||
      ANALYSIS_OUTPUT_LABELS.emptyLabel;
    const key = `${label1}||${label2}`;
    const value = metric ? toNumber(getRowValue(row, metric)) : null;

    if (!map.has(key)) {
      map.set(key, {
        [dimension]: label1,
        [dimension2]: label2,
        count: 0,
        sum: 0,
      });
    }

    const item = map.get(key);
    item.count += 1;
    if (value != null) item.sum += value;
  });

  return Array.from(map.values()).map((item) => ({
    ...item,
    value: operation === "sum" ? item.sum : item.count,
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

  const sourceRows = Array.isArray(table.rows) ? table.rows : [];
  const filters = Array.isArray(candidate.filters) ? candidate.filters : [];
  const rows = applyCandidateFilters(sourceRows, filters);
  const metric = resolveCandidateColumn(candidate, "metric");
  const dimension = resolveCandidateColumn(candidate, "dimension");
  const dimension2 = resolveCandidateColumn(candidate, "dimension2");
  const date = resolveCandidateColumn(candidate, "date");
  const recipeType =
    candidate.recipeType || candidate.recipeId || candidate.type;

  let resultRows = [];
  let resultType = "grouped";
  let operation = candidate.operation || recipeType;

  if (
    recipeType === ANALYSIS_RECIPE_TYPES.GROUP_SUMMARY ||
    recipeType === "group_summary"
  ) {
    resultRows = aggregateRows({
      rows,
      dimension,
      metric,
      operation: "summary",
    });
    operation = "summary";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.GROUP_SUM) {
    resultRows = aggregateRows({ rows, dimension, metric, operation: "sum" });
    operation = "sum";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.GROUP_AVG) {
    resultRows = aggregateRows({
      rows,
      dimension,
      metric,
      operation: "average",
    });
    operation = "average";
  } else if (
    recipeType === ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT ||
    recipeType === ANALYSIS_RECIPE_TYPES.GROUP_COUNT ||
    recipeType === ANALYSIS_RECIPE_TYPES.STATUS_COUNT ||
    recipeType === "category_count"
  ) {
    resultRows = categoryCount({ rows, dimension });
    operation = "count";
  } else if (
    recipeType === ANALYSIS_RECIPE_TYPES.TOP_BOTTOM ||
    recipeType === "top_bottom"
  ) {
    resultRows = topBottom({ rows, dimension, metric });
    operation = "topBottom";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.COMPOSITION_RATIO) {
    resultRows = compositionRatio({ rows, dimension, metric });
    operation = "compositionRatio";
  } else if (
    recipeType === ANALYSIS_RECIPE_TYPES.TIME_TREND ||
    recipeType === "time_trend"
  ) {
    resultRows = timeAggregate({ rows, date, metric, operation: "summary" });
    resultType = "timeSeries";
    operation = "timeTrend";
  } else if (
    recipeType === ANALYSIS_RECIPE_TYPES.TIME_SUM ||
    recipeType === ANALYSIS_RECIPE_TYPES.WIDE_TIME_TREND
  ) {
    resultRows = timeAggregate({ rows, date, metric, operation: "sum" });
    resultType = "timeSeries";
    operation =
      recipeType === ANALYSIS_RECIPE_TYPES.WIDE_TIME_TREND
        ? "wideTimeTrend"
        : "timeSum";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.TIME_AVG) {
    resultRows = timeAggregate({ rows, date, metric, operation: "average" });
    resultType = "timeSeries";
    operation = "timeAverage";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.TIME_COUNT) {
    resultRows = timeAggregate({ rows, date, metric: "", operation: "count" });
    resultType = "timeSeries";
    operation = "timeCount";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.TIME_GROWTH) {
    resultRows = timeGrowth({ rows, date, metric });
    resultType = "timeSeries";
    operation = "growthRate";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.CUMULATIVE_SUM) {
    resultRows = cumulativeSum({ rows, date, metric });
    resultType = "timeSeries";
    operation = "cumulativeSum";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.CROSS_COUNT) {
    resultRows = crossAggregate({
      rows,
      dimension,
      dimension2,
      operation: "count",
    });
    resultType = "crossTable";
    operation = "crossCount";
  } else if (recipeType === ANALYSIS_RECIPE_TYPES.CROSS_SUM) {
    resultRows = crossAggregate({
      rows,
      dimension,
      dimension2,
      metric,
      operation: "sum",
    });
    resultType = "crossTable";
    operation = "crossSum";
  } else {
    return {
      ok: false,
      code: "UNSUPPORTED_RECIPE_TYPE",
      message: "지원하지 않는 분석 후보입니다.",
    };
  }

  return {
    ok: true,
    recipeType,
    resultType,
    operation,
    title: candidate.title,
    tableId: table.tableId,
    sheetName: table.sheetName,
    sourceTableId:
      candidate.sourceTableId || table.sourceTableId || table.tableId,
    sourceSheetName: candidate.sourceSheetName || "",
    columns: candidate.columns,
    groupBy: dimension ? { header: dimension } : null,
    groupBy2: dimension2 ? { header: dimension2 } : null,
    metric: metric
      ? {
          header: metric,
          displayHeader: candidate.metricDisplayHeader || metric,
        }
      : null,
    date: date ? { header: date } : null,
    filters,
    measureIsolation: candidate.measureIsolation || null,
    sourceRowCount: sourceRows.length,
    filteredRowCount: rows.length,
    rows: resultRows,
    rowCount: resultRows.length,
  };
}

module.exports = {
  executeAnalysisRecipeCandidate,
  aggregateRows,
  timeAggregate,
  compositionRatio,
  crossAggregate,
  applyCandidateFilters,
  matchesCandidateFilter,
};
