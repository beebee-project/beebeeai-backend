function toComparable(value) {
  if (value == null || value === "") return null;

  if (typeof value === "number") return value;

  const n = Number(String(value).replace(/,/g, "").trim());
  if (Number.isFinite(n)) return n;

  return String(value).trim();
}

function extractYear(value) {
  if (value == null || value === "") return "";

  const d = value instanceof Date ? value : new Date(value);
  if (!Number.isNaN(d.getTime())) return String(d.getFullYear());

  const m = String(value).match(/(\d{4})/);
  return m ? m[1] : "";
}

function toDate(value) {
  if (value == null || value === "") return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  const d = new Date(value);
  if (!Number.isNaN(d.getTime())) return d;

  return null;
}

function extractMonth(value) {
  const d = toDate(value);
  if (!d) return "";

  return String(d.getMonth() + 1).padStart(2, "0");
}

function extractQuarter(value) {
  const d = toDate(value);
  if (!d) return "";

  return `Q${Math.floor(d.getMonth() / 3) + 1}`;
}

function extractYearMonth(value) {
  const d = toDate(value);
  if (!d) return "";

  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");

  return `${y}-${m}`;
}

function applyDeriveSteps(rows = [], steps = []) {
  const deriveSteps = steps.filter((s) => s.type === "derive");

  if (!deriveSteps.length) return rows;

  return rows.map((row) => {
    const next = { ...row };

    for (const step of deriveSteps) {
      const sourceValue = row[step.sourceColumnKey];

      if (step.fn === "year") {
        next[step.outputKey] = extractYear(sourceValue);
      }

      if (step.fn === "month") {
        next[step.outputKey] = extractMonth(sourceValue);
      }

      if (step.fn === "quarter") {
        next[step.outputKey] = extractQuarter(sourceValue);
      }

      if (step.fn === "yearMonth") {
        next[step.outputKey] = extractYearMonth(sourceValue);
      }
    }

    return next;
  });
}

function rateValue(rows = [], rateStep = {}) {
  const numerator = rateStep.numerator || {};
  const denominator = rateStep.denominator || {};
  const multiplier = Number(rateStep.multiplier || 1);

  const denominatorCount =
    denominator.type === "count" ? rows.length : rows.length;

  if (!denominatorCount) return null;

  let numeratorCount = 0;

  if (numerator.type === "exists") {
    numeratorCount = rows.filter((r) => {
      const v = r[numerator.columnKey];
      return v != null && String(v).trim() !== "";
    }).length;
  } else if (numerator.type === "positive") {
    numeratorCount = rows.filter((r) => {
      const v = r[numerator.columnKey];

      if (v == null || v === "") return false;

      if (typeof v === "boolean") return v === true;

      if (typeof v === "number") return v > 0;

      const s = String(v).trim();

      if (/^(true|yes|y|1)$/i.test(s)) return true;
      if (/완료|성공|전환|구매|가입|불량|오류|실패|달성|처리/.test(s)) {
        return true;
      }

      const n = Number(s.replace(/,/g, ""));
      return Number.isFinite(n) && n > 0;
    }).length;
  } else {
    return null;
  }

  return (numeratorCount / denominatorCount) * multiplier;
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

function applySortRows(rows = [], sortStep = null) {
  if (!sortStep?.by) return rows;

  const dir = sortStep.direction === "asc" ? 1 : -1;

  return [...rows].sort((a, b) => {
    const av = toComparable(a[sortStep.by] ?? a.value);
    const bv = toComparable(b[sortStep.by] ?? b.value);

    if (typeof av === "number" && typeof bv === "number") {
      return (av - bv) * dir;
    }

    return String(av ?? "").localeCompare(String(bv ?? "")) * dir;
  });
}

function applyLimitRows(rows = [], limitStep = null) {
  if (!limitStep?.count) return rows;
  return rows.slice(0, Number(limitStep.count));
}

function applyCompareRows(rows = [], compareStep = null) {
  if (!compareStep || compareStep.method !== "growthRate") return rows;

  const outputHeader = compareStep.outputHeader || "증감률";

  const sorted = [...rows].sort((a, b) => {
    const av = String(a[compareStep.groupBy?.header] ?? "");
    const bv = String(b[compareStep.groupBy?.header] ?? "");
    return av.localeCompare(bv);
  });

  return sorted.map((row, idx) => {
    if (idx === 0) {
      return {
        ...row,
        기준값: null,
        비교값: row.value,
        [outputHeader]: null,
      };
    }

    const prev = sorted[idx - 1];
    const prevValue = Number(prev.value);
    const currValue = Number(row.value);

    const growth =
      Number.isFinite(prevValue) &&
      prevValue !== 0 &&
      Number.isFinite(currValue)
        ? ((currValue - prevValue) / prevValue) *
          Number(compareStep.multiplier || 100)
        : null;

    return {
      ...row,
      기준값: prevValue,
      비교값: currValue,
      [outputHeader]: growth,
    };
  });
}

function executeSinglePipeline(table = {}, intent = {}, steps = []) {
  const rows = applyDeriveSteps(
    Array.isArray(table.rows) ? table.rows : [],
    steps,
  );

  const filterStep = steps.find((s) => s.type === "filter");
  const filters = filterStep?.filters || [];
  const filteredRows = applyFilters(rows, filters);

  const aggregateStep = steps.find((s) => s.type === "aggregate");
  const rateStep = steps.find((s) => s.type === "rate");
  const groupByStep = steps.find((s) => s.type === "groupBy");
  const sortStep = steps.find((s) => s.type === "sort");
  const limitStep = steps.find((s) => s.type === "limit");
  const compareStep = steps.find((s) => s.type === "compare");

  const operation = rateStep
    ? "rate"
    : aggregateStep?.operation || intent.operation;

  const metric = aggregateStep?.metric || intent.metric;

  const groupBy = groupByStep
    ? {
        columnKey: groupByStep.columnKey,
        header: groupByStep.header,
      }
    : intent.groupBy;

  if (groupBy?.columnKey) {
    const groups = groupRows(filteredRows, groupBy);
    const resultRows = [];

    for (const [groupValue, groupRowsValue] of groups.entries()) {
      resultRows.push({
        [groupBy.header || groupBy.columnKey]: groupValue,
        operation: compareStep ? "growthRate" : operation,
        metric: rateStep?.outputHeader || metric?.header || null,
        value: rateStep
          ? rateValue(groupRowsValue, rateStep)
          : aggregate(groupRowsValue, operation, metric),
        rowCount: groupRowsValue.length,
      });
    }

    const comparedRows = applyCompareRows(resultRows, compareStep);
    const finalRows = applyLimitRows(
      applySortRows(comparedRows, sortStep),
      limitStep,
    );

    return {
      ok: true,
      operation: compareStep ? "growthRate" : operation,
      metric: rateStep
        ? { header: rateStep.outputHeader, type: "rate" }
        : metric,
      groupBy,
      filters,
      rowCount: filteredRows.length,
      resultType: "grouped",
      rows: finalRows,
    };
  }

  const value = rateStep
    ? rateValue(filteredRows, rateStep)
    : aggregate(filteredRows, operation, metric);

  const finalListRows =
    operation === "list"
      ? applyLimitRows(applySortRows(filteredRows, sortStep), limitStep)
      : [];

  return {
    ok: true,
    operation,
    metric: rateStep ? { header: rateStep.outputHeader, type: "rate" } : metric,
    groupBy: null,
    filters,
    rowCount: filteredRows.length,
    resultType: operation === "list" ? "rows" : "scalar",
    value,
    rows: finalListRows.length
      ? finalListRows
      : operation === "list"
        ? filteredRows.slice(0, 100)
        : [],
  };
}

function executePipelines(table = {}, intent = {}, plan = {}) {
  const pipelines = Array.isArray(plan.pipelines) ? plan.pipelines : [];

  const results = pipelines.map((pipeline) => {
    const result = executeSinglePipeline(table, intent, pipeline.steps || []);

    return {
      id: pipeline.id,
      label: pipeline.label || pipeline.id,
      ...result,
    };
  });

  return {
    ok: results.every((r) => r.ok),
    resultType: "multi",
    pipelines: results,
  };
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

  const plan = intent.plan || null;
  const steps = plan?.steps || [];

  if (Array.isArray(plan?.pipelines) && plan.pipelines.length) {
    const multiResult = executePipelines(table, intent, plan);

    return {
      ok: multiResult.ok,
      table: intent.table,
      operation: intent.operation,
      metric: intent.metric,
      groupBy: intent.groupBy,
      filters: intent.filters || [],
      plan,
      rowCount: Array.isArray(table.rows) ? table.rows.length : 0,
      resultType: "multi",
      pipelines: multiResult.pipelines,
    };
  }

  const rows = applyDeriveSteps(
    Array.isArray(table.rows) ? table.rows : [],
    steps,
  );

  const filterStep = steps.find((s) => s.type === "filter");
  const filters = filterStep?.filters || intent.filters || [];
  const filteredRows = applyFilters(rows, filters);

  const aggregateStep = steps.find((s) => s.type === "aggregate");
  const rateStep = steps.find((s) => s.type === "rate");
  const groupByStep = steps.find((s) => s.type === "groupBy");
  const sortStep = steps.find((s) => s.type === "sort");
  const limitStep = steps.find((s) => s.type === "limit");
  const deriveStep = steps.find((s) => s.type === "derive");
  const compareStep = steps.find((s) => s.type === "compare");

  const operation = rateStep
    ? "rate"
    : aggregateStep?.operation || intent.operation;

  const metric = aggregateStep?.metric || intent.metric;

  const groupBy = groupByStep
    ? {
        columnKey: groupByStep.columnKey,
        header: groupByStep.header,
      }
    : intent.groupBy;

  if (groupBy?.columnKey) {
    const groups = groupRows(filteredRows, groupBy);
    const resultRows = [];

    for (const [groupValue, groupRowsValue] of groups.entries()) {
      resultRows.push({
        [groupBy.header || groupBy.columnKey]: groupValue,
        operation: compareStep ? "growthRate" : operation,
        metric: rateStep?.outputHeader || metric?.header || null,
        value: rateStep
          ? rateValue(groupRowsValue, rateStep)
          : aggregate(groupRowsValue, operation, metric),
        rowCount: groupRowsValue.length,
      });
    }

    const comparedRows = applyCompareRows(resultRows, compareStep);

    const finalRows = applyLimitRows(
      applySortRows(comparedRows, sortStep),
      limitStep,
    );

    return {
      ok: true,
      table: intent.table,
      operation: compareStep ? "growthRate" : operation,
      metric: rateStep
        ? { header: rateStep.outputHeader, type: "rate" }
        : metric,
      groupBy,
      filters,
      plan,
      rowCount: filteredRows.length,
      resultType: "grouped",
      rows: finalRows,
    };
  }

  const value = rateStep
    ? rateValue(filteredRows, rateStep)
    : aggregate(filteredRows, operation, metric);

  const finalListRows =
    operation === "list"
      ? applyLimitRows(applySortRows(filteredRows, sortStep), limitStep)
      : [];

  return {
    ok: true,
    table: intent.table,
    operation: compareStep ? "growthRate" : operation,
    metric: rateStep ? { header: rateStep.outputHeader, type: "rate" } : metric,
    groupBy: null,
    filters,
    rowCount: filteredRows.length,
    resultType: operation === "list" ? "rows" : "scalar",
    value,
    rows: finalListRows.length
      ? finalListRows
      : operation === "list"
        ? filteredRows.slice(0, 100)
        : [],
    plan,
  };
}

module.exports = { executeQueryIntent };
