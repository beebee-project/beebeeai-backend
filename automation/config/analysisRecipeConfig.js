const ANALYSIS_RECIPE_RULES = [
  {
    recipeType: "group_summary",
    required: ["primaryDimension", "primaryMetric"],
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
    required: ["primaryDate", "primaryMetric"],
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
    required: ["primaryDimension"],
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
    required: ["primaryMetric"],
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

const ANALYSIS_RECIPE_OPTIONS = Object.freeze({
  minTableConfidence: 0.45,
  categoryTypesForDimensionFallback: ["category"],
  excludedRolesForDimensionFallback: ["metric", "date", "id"],
});

function hasRequiredRecipeContext(rule = {}, context = {}) {
  return (rule.required || []).every((key) => Boolean(context[key]));
}

module.exports = {
  ANALYSIS_RECIPE_RULES,
  ANALYSIS_RECIPE_OPTIONS,
  hasRequiredRecipeContext,
};
