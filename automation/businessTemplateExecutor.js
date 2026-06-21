const {
  executeTemplateSections,
} = require("./businessTemplates/commonTemplateHelpers");

const {
  executeSalesReport,
} = require("./businessTemplates/salesReportBuilder");

const {
  executeResearchBudgetReport,
} = require("./businessTemplates/researchBudgetReportBuilder");

const {
  executeHrMonthlyReport,
} = require("./businessTemplates/hrMonthlyReportBuilder");

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
    case "sales_report":
      sections = executeSalesReport({
        normalizedQueryTables,
        templateCandidate,
      });
      break;

    case "research_budget_report":
      sections = executeResearchBudgetReport({
        normalizedQueryTables,
        templateCandidate,
      });
      break;

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
