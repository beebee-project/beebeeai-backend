const AGGREGATE_SUMMARY_OUTPUT_SCHEMA_VERSION =
  "aggregate_summary_output_schema_v1";

function normalizeToken(value = "") {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
}

function hasOwn(row = {}, key = "") {
  return Boolean(row && key && Object.prototype.hasOwnProperty.call(row, key));
}

function finiteOrBlank(value) {
  if (value == null || value === "") return "";
  const number = Number(value);
  return Number.isFinite(number) ? number : "";
}

function isTopBottomLikeResult(result = {}) {
  const recipeType = normalizeToken(result.recipeType || result.recipeId || "");
  const operation = normalizeToken(result.operation || "");

  return recipeType === "topbottom" || operation === "topbottom";
}

function isGroupSummaryResult(result = {}) {
  const recipeType = normalizeToken(result.recipeType || result.recipeId || "");
  const operation = normalizeToken(result.operation || "");

  return recipeType === "groupsummary" || operation === "summary";
}

function isTimeTrendSummaryResult(result = {}) {
  const recipeType = normalizeToken(result.recipeType || result.recipeId || "");
  const operation = normalizeToken(result.operation || "");

  return (
    recipeType === "timetrend" ||
    operation === "timetrend" ||
    operation === "timesummary"
  );
}

function isAggregateSummaryResult(result = {}) {
  return isGroupSummaryResult(result) || isTimeTrendSummaryResult(result);
}

function resolveResultGroupHeader(result = {}) {
  if (isTimeTrendSummaryResult(result)) {
    return result.date?.header || "기간";
  }
  return result.groupBy?.header || "그룹";
}

function resolveResultLabel(
  result = {},
  row = {},
  groupHeader = resolveResultGroupHeader(result),
) {
  if (hasOwn(row, groupHeader)) return row[groupHeader];

  if (isTimeTrendSummaryResult(result)) {
    return row.period ?? row.date ?? row.label ?? "";
  }

  return row.label ?? row.item ?? row.name ?? "";
}

function resolveResultOperation(result = {}, row = {}) {
  return row.operation ?? row.type ?? result.operation ?? "";
}

function resolveResultMetric(result = {}, row = {}, fallback = "") {
  return (
    row.metric ??
    result.metric?.displayHeader ??
    result.metric?.header ??
    fallback ??
    ""
  );
}

function resolveResultRowCount(result = {}, row = {}) {
  const direct = row.rowCount ?? row.count;
  if (direct != null && direct !== "") return direct;
  return isTopBottomLikeResult(result) ? 1 : "";
}

function resolveResultNumericCount(row = {}) {
  const direct = row.numericCount ?? row.validCount;
  if (direct != null && direct !== "") return direct;
  return resolveResultRowCount({}, row);
}

function resolveSummarySum(row = {}) {
  if (hasOwn(row, "sum")) return finiteOrBlank(row.sum);
  if (normalizeToken(row.valueKind || "") === "sum" && hasOwn(row, "value")) {
    return finiteOrBlank(row.value);
  }
  return finiteOrBlank(row.value);
}

function resolveSummaryAverage(row = {}) {
  if (hasOwn(row, "average")) return finiteOrBlank(row.average);
  if (hasOwn(row, "avg")) return finiteOrBlank(row.avg);
  if (
    normalizeToken(row.valueKind || "") === "average" &&
    hasOwn(row, "value")
  ) {
    return finiteOrBlank(row.value);
  }
  return "";
}

function resolvePrimaryResultValue(result = {}, row = {}) {
  if (isAggregateSummaryResult(result)) {
    return resolveSummarySum(row);
  }
  return finiteOrBlank(row.value);
}

function buildAggregateSummaryResultRow({
  labelHeader = "그룹",
  label = "",
  operation = "summary",
  metric = "",
  sum = 0,
  average = null,
  count = 0,
  numericCount = 0,
  extra = {},
} = {}) {
  const safeCount = Number.isFinite(Number(count)) ? Number(count) : 0;
  const safeNumericCount = Number.isFinite(Number(numericCount))
    ? Number(numericCount)
    : 0;
  const safeSum = Number.isFinite(Number(sum)) ? Number(sum) : 0;
  const safeAverage =
    average == null || average === ""
      ? null
      : Number.isFinite(Number(average))
        ? Number(average)
        : null;

  return {
    ...extra,
    [labelHeader || "그룹"]: label,
    operation,
    metric,
    value: safeSum,
    valueKind: "sum",
    sum: safeSum,
    average: safeAverage,
    count: safeCount,
    numericCount: safeNumericCount,
    rowCount: safeCount,
  };
}

function buildAggregateSummaryWorkbookRows(result = {}) {
  const groupHeader = resolveResultGroupHeader(result);
  const metricHeader = resolveResultMetric(
    result,
    {},
    result.metric?.header || "값",
  );

  return (result.rows || []).map((row) => ({
    [groupHeader]: resolveResultLabel(result, row, groupHeader),
    작업: resolveResultOperation(result, row),
    지표: resolveResultMetric(result, row, metricHeader),
    합계: resolveSummarySum(row),
    평균: resolveSummaryAverage(row),
    행수: resolveResultRowCount(result, row),
    유효값수: resolveResultNumericCount(row),
  }));
}

function buildGenericGroupedWorkbookRows(result = {}) {
  const groupHeader = result.groupBy?.header || "그룹";
  const extraKeys = ["기준값", "비교값", "증감률"].filter((key) =>
    (result.rows || []).some((row) => hasOwn(row, key)),
  );

  return (result.rows || []).map((row) => {
    const output = {
      [groupHeader]: resolveResultLabel(result, row, groupHeader),
      작업: resolveResultOperation(result, row),
      지표: resolveResultMetric(result, row, result.metric?.header || "값"),
      값: row.value,
      행수: resolveResultRowCount(result, row),
    };

    for (const key of extraKeys) {
      output[key] = row[key];
    }

    return output;
  });
}

function buildWorkbookRowsFromAnalysisResult(result = {}) {
  if (isAggregateSummaryResult(result)) {
    return buildAggregateSummaryWorkbookRows(result);
  }

  if (result.resultType === "grouped") {
    if (
      result.operation === "multiAggregate" ||
      result.operation === "pipelineCombine"
    ) {
      return result.rows || [];
    }
    return buildGenericGroupedWorkbookRows(result);
  }

  if (result.resultType === "scalar") {
    return [
      {
        지표: result.metric?.header || result.operation,
        값: result.value,
        행수: result.rowCount,
      },
    ];
  }

  if (result.resultType === "pivot") {
    return result.rows || [];
  }

  return result.rows || [];
}

function buildChartDataRowsFromAnalysisResult(result = {}) {
  if (result.resultType !== "grouped" && !isTimeTrendSummaryResult(result)) {
    return [];
  }

  if (
    result.resultType === "grouped" &&
    (result.operation === "multiAggregate" ||
      result.operation === "pipelineCombine")
  ) {
    return result.rows || [];
  }

  const groupHeader = resolveResultGroupHeader(result);
  const metricHeader =
    result.metric?.displayHeader || result.metric?.header || "값";

  return (result.rows || []).map((row) => ({
    [groupHeader]: resolveResultLabel(result, row, groupHeader),
    [metricHeader]: resolvePrimaryResultValue(result, row),
    행수: resolveResultRowCount(result, row),
  }));
}

function buildInsightValueRowsFromAnalysisResult(result = {}) {
  if (result.resultType !== "grouped" && !isTimeTrendSummaryResult(result)) {
    return [];
  }

  const groupHeader = resolveResultGroupHeader(result);

  return (result.rows || [])
    .map((row) => ({
      label: resolveResultLabel(result, row, groupHeader),
      value: resolvePrimaryResultValue(result, row),
    }))
    .filter((row) => Number.isFinite(Number(row.value)))
    .map((row) => ({
      label: row.label,
      value: Number(row.value),
    }));
}

module.exports = {
  AGGREGATE_SUMMARY_OUTPUT_SCHEMA_VERSION,
  isTopBottomLikeResult,
  isGroupSummaryResult,
  isTimeTrendSummaryResult,
  isAggregateSummaryResult,
  resolveResultGroupHeader,
  resolveResultLabel,
  resolveResultOperation,
  resolveResultMetric,
  resolveResultRowCount,
  resolveResultNumericCount,
  resolvePrimaryResultValue,
  buildAggregateSummaryResultRow,
  buildAggregateSummaryWorkbookRows,
  buildWorkbookRowsFromAnalysisResult,
  buildChartDataRowsFromAnalysisResult,
  buildInsightValueRowsFromAnalysisResult,
};
