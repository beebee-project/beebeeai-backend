const ANALYSIS_RECIPE_RULES = [
  {
    recipeType: "group_summary",
    canBuild: ({ primaryDimension, primaryMetric }) =>
      !!primaryDimension && !!primaryMetric,
    build: ({ table, primaryDimension, primaryMetric, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_group_summary`,
        title: `${primaryDimension.header}별 ${primaryMetric.header} 요약`,
        description: `${primaryDimension.header} 기준으로 ${primaryMetric.header}의 합계, 평균, 건수를 분석합니다.`,
        recipeType: "group_summary",
        table,
        metric: primaryMetric,
        dimension: primaryDimension,
      }),
  },
  {
    recipeType: "time_trend",
    canBuild: ({ primaryDate, primaryMetric }) =>
      !!primaryDate && !!primaryMetric,
    build: ({ table, primaryDate, primaryMetric, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_time_trend`,
        title: `${primaryDate.header} 기준 ${primaryMetric.header} 추이`,
        description: `${primaryDate.header}를 기준으로 ${primaryMetric.header}의 기간별 흐름을 분석합니다.`,
        recipeType: "time_trend",
        table,
        metric: primaryMetric,
        date: primaryDate,
      }),
  },
  {
    recipeType: "category_count",
    canBuild: ({ primaryDimension }) => !!primaryDimension,
    build: ({ table, primaryDimension, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_category_count`,
        title: `${primaryDimension.header}별 건수`,
        description: `${primaryDimension.header} 기준으로 데이터 건수를 집계합니다.`,
        recipeType: "category_count",
        table,
        dimension: primaryDimension,
      }),
  },
  {
    recipeType: "top_bottom",
    canBuild: ({ primaryMetric }) => !!primaryMetric,
    build: ({ table, primaryDimension, primaryMetric, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_top_bottom`,
        title: `${primaryMetric.header} 상위/하위 항목`,
        description: `${primaryMetric.header} 기준으로 높은 항목과 낮은 항목을 확인합니다.`,
        recipeType: "top_bottom",
        table,
        metric: primaryMetric,
        dimension: primaryDimension,
      }),
  },
];

function getColumnsByRole(columns = [], role = "") {
  return columns.filter((column) => column.role === role);
}

function makeCandidate({
  id,
  title,
  description,
  recipeType,
  table,
  metric,
  dimension,
  date,
}) {
  return {
    id,
    title,
    description,
    recipeType,
    tableId: table.tableId,
    sheetName: table.sheetName,
    confidence: table.confidence,
    columns: {
      metric: metric?.header || null,
      dimension: dimension?.header || null,
      date: date?.header || null,
    },
  };
}

function uniqueColumns(columns = []) {
  const seen = new Set();
  return columns.filter((column) => {
    const key = column.key || column.header;
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildTableCandidates(table = {}) {
  const candidates = [];

  if (Number(table.confidence) < 0.45) return candidates;

  const columns = Array.isArray(table.columns) ? table.columns : [];

  const metrics = getColumnsByRole(columns, "metric");
  const dimensions = uniqueColumns([
    ...getColumnsByRole(columns, "dimension"),
    ...columns.filter(
      (column) =>
        column.type === "category" &&
        !["metric", "date", "id"].includes(column.role),
    ),
  ]);
  const dates = getColumnsByRole(columns, "date");
  const statuses = getColumnsByRole(columns, "status");

  console.log("[recipe-column-roles]", {
    tableId: table.tableId,
    columns: columns.map((c) => ({
      header: c.header,
      type: c.type,
      role: c.role,
      uniqueCount: c.uniqueCount,
      uniqueRatio: c.uniqueRatio,
    })),
    metrics: metrics.map((c) => c.header),
    dimensions: dimensions.map((c) => c.header),
    dates: dates.map((c) => c.header),
    statuses: statuses.map((c) => c.header),
  });

  const primaryMetric = metrics[0];
  const primaryDimension = dimensions[0] || statuses[0];
  const primaryDate = dates[0];

  ANALYSIS_RECIPE_RULES.forEach((rule) => {
    const context = {
      table,
      columns,
      metrics,
      dimensions,
      dates,
      statuses,
      primaryMetric,
      primaryDimension,
      primaryDate,
      makeCandidate,
    };

    if (rule.canBuild(context)) {
      candidates.push(rule.build(context));
    }
  });

  return candidates;
}

function buildAnalysisRecipeCandidates(normalizedQueryTables = []) {
  if (!Array.isArray(normalizedQueryTables)) return [];

  const candidates = normalizedQueryTables.flatMap(buildTableCandidates);

  console.log(
    "[analysisRecipeCandidates]",
    candidates.map((c) => ({
      recipeType: c.recipeType,
      title: c.title,
      tableId: c.tableId,
    })),
  );

  return candidates;
}

module.exports = {
  buildAnalysisRecipeCandidates,
};
