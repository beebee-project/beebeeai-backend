function getGroupHeader(result = {}) {
  return result.groupBy?.header || result.pivot?.rowGroup?.header || "그룹";
}

function getMetricHeader(result = {}) {
  return result.metric?.header || "값";
}

function getNumericKeys(row = {}, groupHeader = "") {
  return Object.keys(row).filter((key) => {
    if (key === groupHeader) return false;
    if (key === "rowCount") return false;
    if (["operation", "metric"].includes(key)) return false;

    const value = row[key];
    return typeof value === "number" && Number.isFinite(value);
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

function recommendChartSpec(result = {}) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  const groupHeader = getGroupHeader(result);
  const metricHeader = getMetricHeader(result);

  if (!rows.length) return null;

  if (result.resultType === "pivot") {
    const seriesFields = result.pivot?.columns || [];

    return {
      version: "chart_spec_v1",
      recommendedType: seriesFields.length <= 5 ? "stacked_bar" : "pivot_table",
      title: `${groupHeader} 기준 교차 분석`,
      categoryField: groupHeader,
      seriesFields,
      rowCount: rows.length,
    };
  }

  if (
    result.operation === "pipelineCombine" ||
    result.operation === "multiAggregate"
  ) {
    const numericKeys = getNumericKeys(rows[0], groupHeader);

    return {
      version: "chart_spec_v1",
      recommendedType: numericKeys.length <= 3 ? "grouped_bar" : "table",
      title: `${groupHeader} 기준 복수 지표 분석`,
      categoryField: groupHeader,
      seriesFields: numericKeys,
      rowCount: rows.length,
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
    };
  }

  if (hasWindowField(rows)) {
    const windowKey = Object.keys(rows[0] || {}).find((key) =>
      /누적|이동평균|rolling|cumulative/i.test(key),
    );

    return {
      version: "chart_spec_v1",
      recommendedType: "line",
      title: `${groupHeader} 기준 추세 분석`,
      categoryField: groupHeader,
      valueField: windowKey || metricHeader,
      rowCount: rows.length,
    };
  }

  if (result.resultType === "rows") {
    const numericKeys = getNumericKeys(rows[0], groupHeader);
    const valueField = numericKeys[0] || metricHeader;

    return {
      version: "chart_spec_v1",
      recommendedType: "horizontal_bar",
      title: `${valueField} 기준 상위 항목`,
      categoryField: groupHeader,
      valueField,
      rowCount: rows.length,
    };
  }

  if (result.resultType === "grouped") {
    return {
      version: "chart_spec_v1",
      recommendedType: "bar",
      title: `${groupHeader}별 ${metricHeader}`,
      categoryField: groupHeader,
      valueField: metricHeader,
      rowCount: rows.length,
    };
  }

  return null;
}

module.exports = {
  recommendChartSpec,
};
