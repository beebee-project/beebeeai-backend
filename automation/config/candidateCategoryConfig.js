const AUTOMATION_CATEGORY_DEFS = [
  {
    categoryId: "summary",
    title: "요약 집계",
    description: "부서, 항목, 상태 등 기준별 건수·합계·평균을 집계합니다.",
    internalOnly: true,
    recipeTypes: [
      "group_summary",
      "category_count",
      "groupAggregate",
      "multiAggregate",
      "pipelineCombine",
    ],
  },
  {
    categoryId: "trend",
    title: "추이 분석",
    description: "월별·연도별 변화, 누적합계, 이동평균, 성장률을 분석합니다.",
    internalOnly: true,
    recipeTypes: [
      "time_trend",
      "cumulativeSum",
      "rollingAverage",
      "growthRate",
    ],
  },
  {
    categoryId: "ranking",
    title: "순위 / TOP 분석",
    description: "상위·하위 항목을 확인합니다.",
    internalOnly: true,
    recipeTypes: ["top_bottom", "list"],
  },
  {
    categoryId: "cross_summary",
    title: "교차 분석",
    description: "두 기준을 교차하여 요약합니다.",
    internalOnly: true,
    recipeTypes: ["pivot"],
  },
];

module.exports = {
  AUTOMATION_CATEGORY_DEFS,
};
