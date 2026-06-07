function toNumber(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function inferLabelKey(row = {}) {
  const keys = Object.keys(row || {});

  return (
    keys.find((key) => /name|이름|title|제목|label|항목|id$/i.test(key)) ||
    keys.find((key) => typeof row[key] === "string" && row[key].trim()) ||
    keys[0] ||
    "항목"
  );
}

function formatValue(v) {
  const n = toNumber(v);
  if (n == null) return String(v ?? "");

  return Number.isInteger(n)
    ? n.toLocaleString()
    : Number(n.toFixed(2)).toLocaleString();
}

function numericKeys(row = {}, exclude = []) {
  const excluded = new Set(exclude);

  return Object.keys(row || {}).filter((key) => {
    if (excluded.has(key)) return false;
    const n = toNumber(row[key]);
    return n != null;
  });
}

function getPrimaryNumericKey(result = {}) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  if (!rows.length) return null;

  const groupHeader = getGroupHeader(result);

  if (
    result.resultType === "grouped" &&
    result.operation !== "multiAggregate" &&
    result.operation !== "pipelineCombine"
  ) {
    return "value";
  }

  return (
    numericKeys(rows[0], [groupHeader, "rowCount", "기준값", "비교값"])[0] ||
    null
  );
}

function topBottomHighlight(
  rows = [],
  labelKey = "",
  valueKey = "",
  valueLabel = valueKey,
) {
  const values = rows
    .map((r) => ({
      label: r[labelKey],
      value: toNumber(r[valueKey]),
    }))
    .filter((r) => r.value != null);

  const label = valueLabel || valueKey || "값";

  if (!values.length) return [];

  const max = values.reduce((a, b) => (b.value > a.value ? b : a));
  const min = values.reduce((a, b) => (b.value < a.value ? b : a));

  if (max.label === min.label) {
    return [`${max.label}의 ${label} 값은 ${formatValue(max.value)}입니다.`];
  }

  return [
    `${label} 기준 최대는 ${max.label}(${formatValue(max.value)})입니다.`,
    `${label} 기준 최소는 ${min.label}(${formatValue(min.value)})입니다.`,
  ];
}

function growthHighlight(result = {}) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  const groupHeader = getGroupHeader(result);

  const growthRows = rows
    .map((r) => ({
      label: r[groupHeader],
      value: toNumber(r["증감률"]),
    }))
    .filter((r) => r.value != null);

  if (!growthRows.length) return [];

  const max = growthRows.reduce((a, b) => (b.value > a.value ? b : a));
  const min = growthRows.reduce((a, b) => (b.value < a.value ? b : a));

  return [
    `가장 높은 증감률은 ${max.label}의 ${formatValue(max.value)}%입니다.`,
    `가장 낮은 증감률은 ${min.label}의 ${formatValue(min.value)}%입니다.`,
  ];
}

function windowHighlight(result = {}) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  if (!rows.length) return [];

  const groupHeader = getGroupHeader(result);
  const first = rows[0] || {};
  const windowKey = Object.keys(first).find((key) =>
    /누적|이동평균|rolling|cumulative/i.test(key),
  );

  if (!windowKey) return [];

  const last = rows[rows.length - 1];

  return [
    `${windowKey}의 마지막 값은 ${last?.[groupHeader] ?? "마지막 구간"} 기준 ${formatValue(last?.[windowKey])}입니다.`,
  ];
}

function pivotHighlight(result = {}) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  const columns = result.pivot?.columns || [];

  if (!rows.length || !columns.length) return [];

  const filledCount = rows.reduce(
    (sum, row) =>
      sum +
      columns.filter((col) => {
        const v = row[col];
        return v !== "" && v != null;
      }).length,
    0,
  );

  const totalCells = rows.length * columns.length;
  const density = totalCells ? (filledCount / totalCells) * 100 : 0;

  return [
    `교차표는 ${rows.length}개 행과 ${columns.length}개 열 항목으로 구성되었습니다.`,
    `값이 채워진 셀 비율은 ${formatValue(density)}%입니다.`,
  ];
}

function getGroupHeader(result = {}) {
  if (result.groupBy?.header) return result.groupBy.header;
  if (result.pivot?.rowGroup?.header) return result.pivot.rowGroup.header;

  if (result.resultType === "rows") {
    return inferLabelKey(result.rows?.[0] || {});
  }

  return "그룹";
}

function getMetricHeader(result = {}) {
  return result.metric?.header || "값";
}

function buildTitle(result = {}, context = {}) {
  const message = context.message || "";
  if (message) return `${message} 분석 결과`;

  const groupHeader = getGroupHeader(result);
  const metricHeader = getMetricHeader(result);

  if (result.resultType === "pivot") {
    return `${groupHeader} 기준 교차 분석 결과`;
  }

  return `${groupHeader}별 ${metricHeader} 분석 결과`;
}

function buildSummary(result = {}) {
  const rowCount = Array.isArray(result.rows) ? result.rows.length : 0;
  const operation = result.operation || "분석";

  return `${operation} 결과 ${rowCount}건이 생성되었습니다.`;
}

function buildHighlights(result = {}) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  if (!rows.length) return [];

  const groupHeader = getGroupHeader(result);

  if (result.resultType === "pivot") {
    return pivotHighlight(result);
  }

  const highlights = [];

  if (
    result.operation === "multiAggregate" ||
    result.operation === "pipelineCombine"
  ) {
    const metricKeys = numericKeys(rows[0], [groupHeader, "rowCount"]);

    highlights.push(
      `${groupHeader} 기준으로 ${rows.length}개 그룹이 생성되었습니다.`,
    );

    highlights.push(`${metricKeys.length}개 지표가 함께 계산되었습니다.`);

    const primaryKey = metricKeys[0];
    if (primaryKey) {
      highlights.push(
        ...topBottomHighlight(rows, groupHeader, primaryKey, primaryKey),
      );
    }

    return highlights.slice(0, 5);
  }

  highlights.push(
    `${groupHeader} 기준으로 ${rows.length}개 결과가 생성되었습니다.`,
  );

  highlights.push(...growthHighlight(result));
  highlights.push(...windowHighlight(result));

  const primaryKey = getPrimaryNumericKey(result);
  if (primaryKey) {
    const valueLabel =
      primaryKey === "value" ? result.metric?.header || "값" : primaryKey;

    highlights.push(
      ...topBottomHighlight(rows, groupHeader, primaryKey, valueLabel),
    );
  }

  return highlights.slice(0, 5);
}

function buildNarrativeSections(result = {}, context = {}) {
  return {
    version: "report_narrative_v1",
    title: buildTitle(result, context),
    summary: buildSummary(result),
    highlights: buildHighlights(result),
    resultType: result.resultType || "",
    operation: result.operation || "",
  };
}

module.exports = {
  buildNarrativeSections,
};
