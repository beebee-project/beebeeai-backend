const BUSINESS_TEMPLATE_DEFS = [
  {
    templateId: "hr_monthly_report",
    title: "월간 인사 보고서",
    description:
      "인원 현황, 부서별 집계, 추이, 상위/하위 항목을 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["category_count", "group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 100,
  },
  {
    templateId: "research_budget_report",
    title: "연구비 집행 현황",
    description:
      "예산·집행액·항목별 집계와 집행 추이를 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 90,
  },
  {
    templateId: "sales_report",
    title: "매출 분석 보고서",
    description:
      "매출 합계, 기간별 추이, 상위 항목을 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 80,
  },
];

function getRecipeType(candidate = {}) {
  return candidate.recipeType || candidate.type || candidate.recipeId || "";
}

function buildBusinessTemplateCandidate(def, analysisCandidates = []) {
  const matchedRequired = def.requiredRecipeTypes
    .map((type) => analysisCandidates.find((c) => getRecipeType(c) === type))
    .filter(Boolean);

  if (matchedRequired.length < def.requiredRecipeTypes.length) {
    return null;
  }

  const matchedOptional = def.optionalRecipeTypes
    .map((type) => analysisCandidates.find((c) => getRecipeType(c) === type))
    .filter(Boolean);

  const matchedCandidates = [...matchedRequired, ...matchedOptional];

  return {
    templateId: def.templateId,
    title: def.title,
    description: def.description,
    outputTypes: def.outputTypes,
    priority: def.priority,
    confidence: Math.min(
      1,
      (matchedRequired.length + matchedOptional.length * 0.5) /
        (def.requiredRecipeTypes.length + def.optionalRecipeTypes.length * 0.5),
    ),
    candidates: matchedCandidates,
    primaryCandidate: matchedCandidates[0] || null,
  };
}

function buildBusinessTemplateCandidates(analysisCandidates = []) {
  if (!Array.isArray(analysisCandidates)) return [];

  return BUSINESS_TEMPLATE_DEFS.map((def) =>
    buildBusinessTemplateCandidate(def, analysisCandidates),
  )
    .filter(Boolean)
    .sort((a, b) => b.priority - a.priority || b.confidence - a.confidence);
}

module.exports = {
  BUSINESS_TEMPLATE_DEFS,
  buildBusinessTemplateCandidates,
};
