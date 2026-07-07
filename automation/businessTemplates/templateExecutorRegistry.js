const { executeTemplateSections } = require("./commonTemplateHelpers");
const { executeSalesReport } = require("./salesReportBuilder");
const {
  executeResearchBudgetReport,
} = require("./researchBudgetReportBuilder");
const { executeHrMonthlyReport } = require("./hrMonthlyReportBuilder");
const { executeStatusRateReport } = require("./statusRateTemplateBuilder");

const BUSINESS_TEMPLATE_EXECUTORS = Object.freeze({
  sales_report: executeSalesReport,
  research_budget_report: executeResearchBudgetReport,
  hr_monthly_report: executeHrMonthlyReport,

  // statusRate common builder templates
  purchase_inspection_status: executeStatusRateReport,
  service_contract_execution_status: executeStatusRateReport,
  recruitment_applicant_management: executeStatusRateReport,
  research_participant_status: executeStatusRateReport,
  purchase_order_status: executeStatusRateReport,
  asset_equipment_management: executeStatusRateReport,
  customer_inquiry_analysis: executeStatusRateReport,
  event_applicant_status: executeStatusRateReport,
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
