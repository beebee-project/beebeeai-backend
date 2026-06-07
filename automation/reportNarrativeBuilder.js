function toNumber(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function getGroupHeader(result = {}) {
  return result.groupBy?.header || result.pivot?.rowGroup?.header || "그룹";
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
    const columnCount = result.pivot?.columns?.length || 0;

    return [
      `행 기준은 ${result.pivot?.rowGroup?.header || groupHeader}입니다.`,
      `열 기준은 ${result.pivot?.columnGroup?.header || "항목"}입니다.`,
      `총 ${rows.length}개 행과 ${columnCount}개 열 항목으로 구성되었습니다.`,
    ];
  }

  if (
    result.operation === "multiAggregate" ||
    result.operation === "pipelineCombine"
  ) {
    const metricKeys = Object.keys(rows[0] || {}).filter(
      (k) => k !== groupHeader && k !== "rowCount",
    );

    return [
      `${groupHeader} 기준으로 ${rows.length}개 그룹이 생성되었습니다.`,
      `${metricKeys.length}개 지표가 함께 계산되었습니다.`,
    ];
  }

  const valueRows = rows
    .map((r) => ({
      label: r[groupHeader],
      value: toNumber(r.value),
    }))
    .filter((r) => r.value != null);

  if (!valueRows.length) {
    return [`${groupHeader} 기준으로 ${rows.length}개 결과가 생성되었습니다.`];
  }

  const max = valueRows.reduce((a, b) => (b.value > a.value ? b : a));
  const min = valueRows.reduce((a, b) => (b.value < a.value ? b : a));

  return [
    `가장 높은 값은 ${max.label}의 ${max.value}입니다.`,
    `가장 낮은 값은 ${min.label}의 ${min.value}입니다.`,
  ];
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
