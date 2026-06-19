const { executeAnalysisRecipeCandidate } = require("./analysisRecipeExecutor");

function executeTemplateSections({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const candidates = Array.isArray(templateCandidate.candidates)
    ? templateCandidate.candidates
    : [];

  return candidates
    .map((candidate, index) => {
      const result = executeAnalysisRecipeCandidate({
        normalizedQueryTables,
        candidate,
      });

      if (!result?.ok) return null;

      return {
        sectionId:
          candidate.recipeType ||
          candidate.type ||
          candidate.recipeId ||
          `section_${index + 1}`,
        title:
          candidate.title ||
          candidate.name ||
          candidate.label ||
          `섹션 ${index + 1}`,
        candidate,
        result,
      };
    })
    .filter(Boolean);
}

function executeHrMonthlyReport(args) {
  return executeTemplateSections(args);
}

function executeBusinessTemplate({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const templateId = templateCandidate.templateId;

  if (!templateId) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_ID_REQUIRED",
      message: "templateId가 필요합니다.",
    };
  }

  let sections = [];

  switch (templateId) {
    case "hr_monthly_report":
      sections = executeHrMonthlyReport({
        normalizedQueryTables,
        templateCandidate,
      });
      break;

    default:
      sections = executeTemplateSections({
        normalizedQueryTables,
        templateCandidate,
      });
      break;
  }

  if (!sections.length) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_EXECUTION_EMPTY",
      message: "실행 가능한 템플릿 섹션이 없습니다.",
    };
  }

  return {
    ok: true,
    resultType: "businessTemplate",
    templateId,
    title: templateCandidate.title || templateId,
    description: templateCandidate.description || "",
    sections,
  };
}

module.exports = {
  executeBusinessTemplate,
};
