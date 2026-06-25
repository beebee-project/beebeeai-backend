const {
  ANALYSIS_RECIPE_RULES,
  ANALYSIS_RECIPE_OPTIONS,
  hasRequiredRecipeContext,
} = require("./config/analysisRecipeConfig");

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

  if (Number(table.confidence) < ANALYSIS_RECIPE_OPTIONS.minTableConfidence) {
    return candidates;
  }

  const columns = Array.isArray(table.columns) ? table.columns : [];

  const metrics = getColumnsByRole(columns, "metric");
  const dimensions = uniqueColumns([
    ...getColumnsByRole(columns, "dimension"),
    ...columns.filter(
      (column) =>
        ANALYSIS_RECIPE_OPTIONS.categoryTypesForDimensionFallback.includes(
          column.type,
        ) &&
        !ANALYSIS_RECIPE_OPTIONS.excludedRolesForDimensionFallback.includes(
          column.role,
        ),
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

    if (hasRequiredRecipeContext(rule, context)) {
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
