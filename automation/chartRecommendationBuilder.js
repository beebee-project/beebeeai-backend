const {
  buildChartTitle,
  buildFieldLabels,
  buildInsightText,
  decorateChartSpec,
} = require("./reportDisplayUtils");

function inferLabelKey(row = {}) {
  const keys = Object.keys(row || {});

  return (
    keys.find((key) => /name|이름|title|제목|label|항목|id$/i.test(key)) ||
    keys.find((key) => typeof row[key] === "string" && row[key].trim()) ||
    keys[0] ||
    "항목"
  );
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
  return result.metric?.header || result.metricHeader || "값";
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

function attachDisplayMeta(spec = null, result = {}, context = {}) {
  if (!spec) return null;
  const rows = Array.isArray(result.rows) ? result.rows : [];
  const seriesFields = Array.isArray(spec.seriesFields)
    ? spec.seriesFields
    : spec.valueField
      ? [spec.valueField]
      : [];
  const displayContext = {
    title: spec.title,
    sectionTitle: context.sectionTitle || result.title,
    metricHeader: getMetricHeader(result),
    categoryField: spec.categoryField,
    rows,
  };

  const fieldLabels = buildFieldLabels(
    [spec.categoryField, ...seriesFields],
    displayContext,
  );

  return decorateChartSpec(
    {
      ...spec,
      fieldLabels,
      title: buildChartTitle({
        categoryField: spec.categoryField,
        valueField: spec.valueField,
        seriesFields,
        rows,
        context: displayContext,
      }),
      insight: buildInsightText({
        categoryField: spec.categoryField,
        valueField: spec.valueField,
        seriesFields,
        rows,
        context: { ...displayContext, fieldLabels },
      }),
    },
    result,
    displayContext,
  );
}

function recommendChartSpec(result = {}) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  const groupHeader = getGroupHeader(result);
  const metricHeader = getMetricHeader(result);

  if (!rows.length) return null;

  if (result.resultType === "pivot") {
    const seriesFields = result.pivot?.columns || [];

    return attachDisplayMeta(
      {
        version: "chart_spec_v1",
        recommendedType:
          seriesFields.length <= 5 ? "stacked_bar" : "pivot_table",
        title: `${groupHeader} 기준 교차 분석`,
        categoryField: groupHeader,
        seriesFields,
        rowCount: rows.length,
      },
      result,
    );
  }

  if (
    result.operation === "pipelineCombine" ||
    result.operation === "multiAggregate"
  ) {
    const numericKeys = getNumericKeys(rows[0], groupHeader);

    return attachDisplayMeta(
      {
        version: "chart_spec_v1",
        recommendedType: numericKeys.length <= 3 ? "grouped_bar" : "table",
        title: `${groupHeader} 기준 복수 지표 분석`,
        categoryField: groupHeader,
        seriesFields: numericKeys,
        rowCount: rows.length,
      },
      result,
    );
  }

  if (hasGrowthField(rows)) {
    return attachDisplayMeta(
      {
        version: "chart_spec_v1",
        recommendedType: "line",
        title: `${groupHeader} 기준 증감 분석`,
        categoryField: groupHeader,
        valueField: "증감률",
        rowCount: rows.length,
      },
      result,
    );
  }

  if (hasWindowField(rows)) {
    const windowKey = Object.keys(rows[0] || {}).find((key) =>
      /누적|이동평균|rolling|cumulative/i.test(key),
    );

    return attachDisplayMeta(
      {
        version: "chart_spec_v1",
        recommendedType: "line",
        title: `${groupHeader} 기준 추세 분석`,
        categoryField: groupHeader,
        valueField: windowKey || metricHeader,
        rowCount: rows.length,
      },
      result,
    );
  }

  if (result.resultType === "rows") {
    const numericKeys = getNumericKeys(rows[0], groupHeader);
    const valueField = numericKeys[0] || metricHeader;

    return attachDisplayMeta(
      {
        version: "chart_spec_v1",
        recommendedType: "horizontal_bar",
        title: `${valueField} 기준 상위 항목`,
        categoryField: groupHeader,
        valueField,
        rowCount: rows.length,
      },
      result,
    );
  }

  if (result.resultType === "grouped") {
    return attachDisplayMeta(
      {
        version: "chart_spec_v1",
        recommendedType: "bar",
        title: `${groupHeader}별 ${metricHeader}`,
        categoryField: groupHeader,
        valueField: metricHeader,
        rowCount: rows.length,
      },
      result,
    );
  }

  return null;
}

module.exports = {
  recommendChartSpec,
};
