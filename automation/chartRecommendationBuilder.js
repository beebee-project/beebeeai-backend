function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null || value === "") return null;

  const n = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

function inferLabelKey(row = {}) {
  const keys = Object.keys(row || {});

  return (
    keys.find((key) =>
      /name|이름|성명|title|제목|label|항목|부서|팀|구분|분류|연월|월|일자|date|id$/i.test(
        key,
      ),
    ) ||
    keys.find((key) => typeof row[key] === "string" && row[key].trim()) ||
    keys[0] ||
    "항목"
  );
}

function getRows(result = {}) {
  return Array.isArray(result.rows) ? result.rows : [];
}

function getGroupHeader(result = {}) {
  const rows = getRows(result);
  const first = rows[0] || {};

  if (result.groupBy?.header) return result.groupBy.header;
  if (result.pivot?.rowGroup?.header) return result.pivot.rowGroup.header;

  const inferred = inferLabelKey(first);
  return inferred || "항목";
}

function getMetricHeader(result = {}, fallback = "값") {
  const rows = getRows(result);
  const first = rows[0] || {};

  if (result.metric?.header) return result.metric.header;
  if (Object.prototype.hasOwnProperty.call(first, "value")) return "value";

  return fallback;
}

function getNumericKeys(row = {}, groupHeader = "") {
  return Object.keys(row || {}).filter((key) => {
    if (key === groupHeader) return false;
    if (["rowCount", "operation", "metric", "type"].includes(key)) return false;

    return toNumber(row[key]) != null;
  });
}

function hasGrowthField(rows = []) {
  return rows.some((r) => Object.prototype.hasOwnProperty.call(r, "증감률"));
}

function hasWindowField(rows = []) {
  return rows.some((r) =>
    Object.keys(r || {}).some((key) =>
      /누적|이동평균|rolling|cumulative/i.test(key),
    ),
  );
}

function pickValueField(rows = [], groupHeader = "", preferred = "") {
  const first = rows[0] || {};
  if (preferred && Object.prototype.hasOwnProperty.call(first, preferred)) {
    return preferred;
  }

  if (Object.prototype.hasOwnProperty.call(first, "value")) return "value";

  const numericKeys = getNumericKeys(first, groupHeader);
  return (
    numericKeys.find((key) => /평균|average|avg/i.test(key)) ||
    numericKeys.find((key) =>
      /합계|sum|total|금액|매출|연봉|비용|amount|revenue/i.test(key),
    ) ||
    numericKeys.find((key) => /count|건수|개수|인원|수/i.test(key)) ||
    numericKeys[0] ||
    null
  );
}

function chartInsight({ rows = [], categoryField = "", valueField = "" } = {}) {
  if (!rows.length || !categoryField || !valueField) return "";

  const values = rows
    .map((row) => ({
      label: row[categoryField],
      value: toNumber(row[valueField]),
    }))
    .filter((row) => row.label != null && row.value != null);

  if (!values.length) return "";

  const top = values.reduce((a, b) => (b.value > a.value ? b : a));
  return `${categoryField} 기준 ${valueField} 최상위 항목은 ${top.label}입니다.`;
}

function recommendChartSpec(result = {}) {
  const rows = getRows(result);
  const groupHeader = getGroupHeader(result);
  const metricHeader = getMetricHeader(result);

  if (!rows.length) return null;

  if (result.resultType === "pivot") {
    const seriesFields =
      result.pivot?.columns || getNumericKeys(rows[0], groupHeader);

    if (!seriesFields.length) return null;

    return {
      version: "chart_spec_v1",
      recommendedType: seriesFields.length <= 5 ? "stacked_bar" : "pivot_table",
      title: `${groupHeader} 기준 교차 분석`,
      categoryField: groupHeader,
      seriesFields,
      rowCount: rows.length,
      insight: `${groupHeader} 기준으로 ${seriesFields.length}개 항목을 비교합니다.`,
    };
  }

  if (
    result.operation === "pipelineCombine" ||
    result.operation === "multiAggregate"
  ) {
    const numericKeys = getNumericKeys(rows[0], groupHeader);

    if (!numericKeys.length) return null;

    return {
      version: "chart_spec_v1",
      recommendedType: numericKeys.length <= 3 ? "grouped_bar" : "table",
      title: `${groupHeader} 기준 복수 지표 분석`,
      categoryField: groupHeader,
      seriesFields: numericKeys,
      rowCount: rows.length,
      insight: `${groupHeader}별 ${numericKeys.length}개 지표를 함께 비교합니다.`,
    };
  }

  if (hasGrowthField(rows)) {
    return {
      version: "chart_spec_v1",
      recommendedType: "line",
      title: `${groupHeader} 기준 증감 분석`,
      categoryField: groupHeader,
      valueField: "증감률",
      rowCount: rows.length,
      insight: chartInsight({
        rows,
        categoryField: groupHeader,
        valueField: "증감률",
      }),
    };
  }

  if (hasWindowField(rows)) {
    const windowKey = Object.keys(rows[0] || {}).find((key) =>
      /누적|이동평균|rolling|cumulative/i.test(key),
    );

    const valueField =
      windowKey || pickValueField(rows, groupHeader, metricHeader);
    if (!valueField) return null;

    return {
      version: "chart_spec_v1",
      recommendedType: "line",
      title: `${groupHeader} 기준 추세 분석`,
      categoryField: groupHeader,
      valueField,
      rowCount: rows.length,
      insight: chartInsight({ rows, categoryField: groupHeader, valueField }),
    };
  }

  if (result.resultType === "rows") {
    const valueField = pickValueField(rows, groupHeader, metricHeader);
    if (!valueField) return null;

    return {
      version: "chart_spec_v1",
      recommendedType: "horizontal_bar",
      title: `${valueField} 기준 상위 항목`,
      categoryField: groupHeader,
      valueField,
      rowCount: rows.length,
      insight: chartInsight({ rows, categoryField: groupHeader, valueField }),
    };
  }

  if (result.resultType === "grouped" || rows.length >= 2) {
    const valueField = pickValueField(rows, groupHeader, metricHeader);
    if (!valueField) return null;

    return {
      version: "chart_spec_v1",
      recommendedType: rows.length > 8 ? "horizontal_bar" : "bar",
      title: `${groupHeader}별 ${valueField}`,
      categoryField: groupHeader,
      valueField,
      rowCount: rows.length,
      insight: chartInsight({ rows, categoryField: groupHeader, valueField }),
    };
  }

  return null;
}

module.exports = {
  recommendChartSpec,
};
