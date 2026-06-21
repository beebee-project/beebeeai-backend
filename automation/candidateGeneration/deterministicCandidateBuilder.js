const {
  buildAnalysisRecipeCandidates,
} = require("../analysisRecipeCandidateBuilder");
const {
  buildBusinessTemplateCandidates,
} = require("../businessTemplateConfig");

function getRecipeType(candidate = {}) {
  return candidate.recipeType || candidate.type || candidate.recipeId || "";
}

function isRecipeType(candidate = {}, types = []) {
  return types.includes(getRecipeType(candidate));
}

function buildAutomationCategoryCandidates(analysisRecipeCandidates = []) {
  const list = Array.isArray(analysisRecipeCandidates)
    ? analysisRecipeCandidates
    : [];

  const groupTypes = [
    "group_summary",
    "category_count",
    "groupAggregate",
    "multiAggregate",
    "pipelineCombine",
  ];

  const trendTypes = [
    "time_trend",
    "cumulativeSum",
    "rollingAverage",
    "growthRate",
  ];

  const rankingTypes = ["top_bottom", "list"];
  const pivotTypes = ["pivot"];

  const groupCandidates = list.filter((c) => isRecipeType(c, groupTypes));
  const trendCandidates = list.filter((c) => isRecipeType(c, trendTypes));
  const rankingCandidates = list.filter((c) => isRecipeType(c, rankingTypes));
  const pivotCandidates = list.filter((c) => isRecipeType(c, pivotTypes));

  const categories = [];

  if (groupCandidates.length) {
    categories.push({
      categoryId: "summary",
      title: "요약 집계",
      description: "부서, 항목, 상태 등 기준별 건수·합계·평균을 집계합니다.",
      internalOnly: true,
      candidates: groupCandidates,
    });
  }

  if (trendCandidates.length) {
    categories.push({
      categoryId: "trend",
      title: "추이 분석",
      description: "월별·연도별 변화, 누적합계, 이동평균, 성장률을 분석합니다.",
      internalOnly: true,
      candidates: trendCandidates,
    });
  }

  if (rankingCandidates.length) {
    categories.push({
      categoryId: "ranking",
      title: "순위 / TOP 분석",
      description: "상위·하위 항목을 확인합니다.",
      internalOnly: true,
      candidates: rankingCandidates,
    });
  }

  if (pivotCandidates.length) {
    categories.push({
      categoryId: "cross_summary",
      title: "교차 분석",
      description: "두 기준을 교차하여 요약합니다.",
      internalOnly: true,
      candidates: pivotCandidates,
    });
  }

  return categories;
}

function buildDeterministicCandidateBundle({
  normalizedQueryTables = [],
  source = "deterministic",
} = {}) {
  const safeTables = Array.isArray(normalizedQueryTables)
    ? normalizedQueryTables
    : [];

  const analysisRecipeCandidates = buildAnalysisRecipeCandidates(safeTables);
  const categoryCandidates = buildAutomationCategoryCandidates(
    analysisRecipeCandidates,
  );
  const businessTemplateCandidates = buildBusinessTemplateCandidates(
    analysisRecipeCandidates,
  );

  return {
    analysisRecipeCandidates,
    categoryCandidates,
    businessTemplateCandidates,
    candidateGeneration: {
      version: "candidate_generation_v1",
      source,
      deterministic: {
        used: true,
        counts: {
          normalizedTables: safeTables.length,
          analysisRecipeCandidates: analysisRecipeCandidates.length,
          categoryCandidates: categoryCandidates.length,
          businessTemplateCandidates: businessTemplateCandidates.length,
        },
      },
      aiReranker: {
        enabled: false,
        used: false,
        skippedReason: "NOT_REQUESTED",
      },
      validation: {
        used: false,
      },
      generatedAt: new Date().toISOString(),
    },
  };
}

module.exports = {
  buildDeterministicCandidateBundle,
  buildAutomationCategoryCandidates,
};
