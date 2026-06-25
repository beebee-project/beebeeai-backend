const { executeTemplateSections } = require("./commonTemplateHelpers");
const { executeSalesReport } = require("./salesReportBuilder");
const {
  executeResearchBudgetReport,
} = require("./researchBudgetReportBuilder");
const { executeHrMonthlyReport } = require("./hrMonthlyReportBuilder");

const BUSINESS_TEMPLATE_EXECUTORS = Object.freeze({
  sales_report: executeSalesReport,
  research_budget_report: executeResearchBudgetReport,
  hr_monthly_report: executeHrMonthlyReport,
});

function getBusinessTemplateExecutor(templateId = "") {
  return BUSINESS_TEMPLATE_EXECUTORS[templateId] || executeTemplateSections;
}

function hasBusinessTemplateExecutor(templateId = "") {
  return Boolean(BUSINESS_TEMPLATE_EXECUTORS[templateId]);
}

module.exports = {
  BUSINESS_TEMPLATE_EXECUTORS,
  getBusinessTemplateExecutor,
  hasBusinessTemplateExecutor,
};
