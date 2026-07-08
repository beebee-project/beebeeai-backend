const { executeTemplateSections } = require("./commonTemplateHelpers");
const { executeSalesReport } = require("./salesReportBuilder");
const {
  executeResearchBudgetReport,
} = require("./researchBudgetReportBuilder");
const { executeHrMonthlyReport } = require("./hrMonthlyReportBuilder");
const { executeStatusRateReport } = require("./statusRateTemplateBuilder");
const { executeSurveyScoreReport } = require("./surveyScoreTemplateBuilder");
const {
  executeInventoryFlowReport,
} = require("./inventoryFlowTemplateBuilder");

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
  education_completion_status: executeStatusRateReport,
  maintenance_request_status: executeStatusRateReport,
  safety_inspection_status: executeStatusRateReport,
  delivery_logistics_status: executeStatusRateReport,
  project_progress_status: executeStatusRateReport,
  task_issue_tracking_report: executeStatusRateReport,
  grant_application_status: executeStatusRateReport,
  attendance_check_report: executeStatusRateReport,
  quality_defect_analysis: executeStatusRateReport,

  // surveyScore common builder templates
  survey_satisfaction_analysis: executeSurveyScoreReport,
  education_feedback_report: executeSurveyScoreReport,
  event_satisfaction_report: executeSurveyScoreReport,
  course_evaluation_report: executeSurveyScoreReport,
  customer_satisfaction_report: executeSurveyScoreReport,
  internal_survey_score_report: executeSurveyScoreReport,

  // inventoryFlow common builder templates
  inventory_stock_status: executeInventoryFlowReport,
  inventory_inout_flow_report: executeInventoryFlowReport,
  asset_lifecycle_report: executeInventoryFlowReport,
  equipment_rental_status: executeInventoryFlowReport,
  supply_usage_report: executeInventoryFlowReport,
  warehouse_movement_report: executeInventoryFlowReport,
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
