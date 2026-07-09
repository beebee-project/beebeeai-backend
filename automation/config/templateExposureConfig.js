const TEMPLATE_EXPOSURE_VERSION = "template_exposure_config_v1";

const TEMPLATE_EXPOSURE_LEVELS = Object.freeze({
  PRIMARY_VERIFIED: "primaryVerified",
  GENERAL_VERIFIED: "generalVerified",
  CONDITIONAL: "conditional",
  QUARANTINED: "quarantined",
});

const EXPOSURE_LEVEL_DEFS = Object.freeze({
  [TEMPLATE_EXPOSURE_LEVELS.PRIMARY_VERIFIED]: Object.freeze({
    exposureLevel: TEMPLATE_EXPOSURE_LEVELS.PRIMARY_VERIFIED,
    label: "검증 완료",
    shortLabel: "강추천",
    description:
      "전용 executor 또는 공통 builder 기반으로 회귀 산출물 검증을 통과한 템플릿입니다.",
    priority: 100,
    displayOrder: 10,
    hideFromFrontend: false,
  }),
  [TEMPLATE_EXPOSURE_LEVELS.GENERAL_VERIFIED]: Object.freeze({
    exposureLevel: TEMPLATE_EXPOSURE_LEVELS.GENERAL_VERIFIED,
    label: "검증 완료",
    shortLabel: "일반 노출",
    description:
      "정의/힌트 기반 템플릿 중 회귀 산출물 검증을 통과한 템플릿입니다.",
    priority: 80,
    displayOrder: 20,
    hideFromFrontend: false,
  }),
  [TEMPLATE_EXPOSURE_LEVELS.CONDITIONAL]: Object.freeze({
    exposureLevel: TEMPLATE_EXPOSURE_LEVELS.CONDITIONAL,
    label: "조건부 노출",
    shortLabel: "조건부",
    description:
      "카탈로그에는 포함되어 있으나 아직 직접 산출물 회귀 검증이 없는 템플릿입니다.",
    priority: 40,
    displayOrder: 60,
    hideFromFrontend: false,
  }),
  [TEMPLATE_EXPOSURE_LEVELS.QUARANTINED]: Object.freeze({
    exposureLevel: TEMPLATE_EXPOSURE_LEVELS.QUARANTINED,
    label: "격리 보류",
    shortLabel: "보류",
    description:
      "fetch failed 또는 기대값 불안정 이력이 있어 stage logging 후 재투입할 템플릿입니다.",
    priority: 0,
    displayOrder: 99,
    hideFromFrontend: true,
    disabledReason:
      "현재 버전에서는 안정성 확인 전까지 프론트 추천에서 제외됩니다.",
  }),
});

const TEMPLATE_DOMAIN_DISPLAY = Object.freeze({
  sales: Object.freeze({
    domain: "sales",
    label: "매출·영업",
    shortLabel: "매출",
    displayOrder: 10,
    description: "매출, 판매수량, 상품·거래처별 실적 분석",
  }),
  budget: Object.freeze({
    domain: "budget",
    label: "예산·지출·정산",
    shortLabel: "예산",
    displayOrder: 20,
    description: "연구비, 카드, 출장비, 회의비, 계약·정산 분석",
  }),
  hr: Object.freeze({
    domain: "hr",
    label: "인사·조직·명단",
    shortLabel: "인사",
    displayOrder: 30,
    description: "인력, 명단, 채용, 참여자, 부서·직급 현황",
  }),
  inventory: Object.freeze({
    domain: "inventory",
    label: "재고·물류·자산",
    shortLabel: "재고",
    displayOrder: 40,
    description: "재고, 입출고, 발주, 자산, 장비, 창고 이동 분석",
  }),
  operation: Object.freeze({
    domain: "operation",
    label: "운영·상태 관리",
    shortLabel: "운영",
    displayOrder: 50,
    description: "이수, 완료, 처리, 점검, 배송, 이슈 상태 분석",
  }),
  survey: Object.freeze({
    domain: "survey",
    label: "설문·평가·만족도",
    shortLabel: "설문",
    displayOrder: 60,
    description: "만족도, 점수, 평점, 문항별 평균과 분포 분석",
  }),
  project: Object.freeze({
    domain: "project",
    label: "프로젝트·성과지표",
    shortLabel: "성과",
    displayOrder: 70,
    description: "프로젝트 진행, 지원사업, KPI 성과 관리",
  }),
  customer: Object.freeze({
    domain: "customer",
    label: "고객·민원·CS",
    shortLabel: "CS",
    displayOrder: 80,
    description: "고객 문의, 민원 유형, 처리 상태 분석",
  }),
  publicData: Object.freeze({
    domain: "publicData",
    label: "공공데이터·통계표",
    shortLabel: "공공통계",
    displayOrder: 90,
    description: "공공 통계표, wide/month table, 다단 헤더 데이터 분석",
  }),
  general: Object.freeze({
    domain: "general",
    label: "일반 자동화",
    shortLabel: "일반",
    displayOrder: 100,
    description: "도메인 분류가 어려운 일반 데이터 자동화",
  }),
});

const PRIMARY_VERIFIED_TEMPLATE_IDS = Object.freeze([
  "sales_report",
  "research_budget_report",
  "hr_monthly_report",
  "purchase_inspection_status",
  "service_contract_execution_status",
  "recruitment_applicant_management",
  "research_participant_status",
  "purchase_order_status",
  "asset_equipment_management",
  "customer_inquiry_analysis",
  "event_applicant_status",
  "education_completion_status",
  "maintenance_request_status",
  "safety_inspection_status",
  "delivery_logistics_status",
  "project_progress_status",
  "task_issue_tracking_report",
  "grant_application_status",
  "attendance_check_report",
  "survey_satisfaction_analysis",
  "education_feedback_report",
  "event_satisfaction_report",
  "course_evaluation_report",
  "customer_satisfaction_report",
  "inventory_stock_status",
  "inventory_inout_flow_report",
  "asset_lifecycle_report",
  "equipment_rental_status",
  "supply_usage_report",
  "warehouse_movement_report",
]);

const GENERAL_VERIFIED_TEMPLATE_IDS = Object.freeze([
  "corporate_card_usage_report",
  "travel_expense_settlement_analysis",
  "meeting_expense_usage_report",
  "monthly_expense_report",
  "employee_roster_status",
  "vendor_performance_report",
  "revenue_cost_profit_report",
  "kpi_performance_dashboard",
  "energy_usage_report",
]);

const QUARANTINED_TEMPLATE_IDS = Object.freeze([]);

const CONDITIONAL_TEMPLATE_IDS = Object.freeze([
  "purchase_analysis_report",
  "sales_quantity_analysis",
  "branch_performance_report",
  "regional_performance_report",
  "product_sales_analysis",
  "account_sales_analysis",
  "project_expense_settlement_summary",
  "contract_payment_schedule_report",
  "public_monthly_statistics_report",
  "quality_defect_analysis",
  "internal_survey_score_report",
]);

const TEMPLATE_EXPOSURE_OVERRIDES = Object.freeze(
  Object.fromEntries([
    ...PRIMARY_VERIFIED_TEMPLATE_IDS.map((templateId) => [
      templateId,
      TEMPLATE_EXPOSURE_LEVELS.PRIMARY_VERIFIED,
    ]),
    ...GENERAL_VERIFIED_TEMPLATE_IDS.map((templateId) => [
      templateId,
      TEMPLATE_EXPOSURE_LEVELS.GENERAL_VERIFIED,
    ]),
    ...CONDITIONAL_TEMPLATE_IDS.map((templateId) => [
      templateId,
      TEMPLATE_EXPOSURE_LEVELS.CONDITIONAL,
    ]),
    ...QUARANTINED_TEMPLATE_IDS.map((templateId) => [
      templateId,
      TEMPLATE_EXPOSURE_LEVELS.QUARANTINED,
    ]),
  ]),
);

function normalizeText(value = "") {
  return String(value || "").trim();
}

function getExposureLevelDef(level = "") {
  return (
    EXPOSURE_LEVEL_DEFS[normalizeText(level)] ||
    EXPOSURE_LEVEL_DEFS[TEMPLATE_EXPOSURE_LEVELS.CONDITIONAL]
  );
}

function getTemplateDomainDisplay(domain = "") {
  const key = normalizeText(domain) || "general";
  return TEMPLATE_DOMAIN_DISPLAY[key] || TEMPLATE_DOMAIN_DISPLAY.general;
}

function getTemplateExposureLevel(templateId = "") {
  const key = normalizeText(templateId);
  return (
    TEMPLATE_EXPOSURE_OVERRIDES[key] || TEMPLATE_EXPOSURE_LEVELS.CONDITIONAL
  );
}

function getTemplateExposure(templateId = "", template = {}) {
  const exposureLevel = getTemplateExposureLevel(
    templateId || template.templateId,
  );
  const levelDef = getExposureLevelDef(exposureLevel);
  const domainDisplay = getTemplateDomainDisplay(template.domain);
  const title = normalizeText(
    template.title || template.templateId || templateId,
  );
  const domainOrder = Number(domainDisplay.displayOrder || 100);
  const exposureOrder = Number(levelDef.displayOrder || 99);

  return {
    version: TEMPLATE_EXPOSURE_VERSION,
    templateId: normalizeText(templateId || template.templateId),
    exposureLevel,
    label: levelDef.label,
    shortLabel: levelDef.shortLabel,
    description: levelDef.description,
    priority: Number(levelDef.priority || 0),
    displayOrder: domainOrder * 100 + exposureOrder,
    exposureDisplayOrder: exposureOrder,
    domainDisplayOrder: domainOrder,
    hideFromFrontend: levelDef.hideFromFrontend === true,
    disabledReason: levelDef.disabledReason || "",
    domain: domainDisplay.domain,
    domainLabel: template.domainLabel || domainDisplay.label,
    domainShortLabel: domainDisplay.shortLabel,
    domainDescription: domainDisplay.description,
    implementationLevel: template.implementationLevel || "",
    implementationLevelLabel: template.implementationLevelLabel || "",
    displayTitle: title,
    displaySubtitle:
      template.description || domainDisplay.description || levelDef.description,
  };
}

function getTemplateExposureBadges(template = {}) {
  const exposure = getTemplateExposure(template.templateId, template);
  return [
    exposure.domainShortLabel || exposure.domainLabel,
    exposure.shortLabel || exposure.label,
    template.implementationLevelLabel ||
      exposure.implementationLevelLabel ||
      "",
  ].filter(Boolean);
}

function buildTemplateExposureSummary(templates = []) {
  const list = Array.isArray(templates) ? templates : [];
  const exposureCounts = {};
  const domainCounts = {};

  for (const template of list) {
    const exposure = getTemplateExposure(template.templateId, template);
    exposureCounts[exposure.exposureLevel] =
      (exposureCounts[exposure.exposureLevel] || 0) + 1;
    domainCounts[exposure.domain] = (domainCounts[exposure.domain] || 0) + 1;
  }

  return {
    version: TEMPLATE_EXPOSURE_VERSION,
    totalCount: list.length,
    exposureCounts,
    domainCounts,
    hiddenCount: exposureCounts[TEMPLATE_EXPOSURE_LEVELS.QUARANTINED] || 0,
  };
}

module.exports = {
  TEMPLATE_EXPOSURE_VERSION,
  TEMPLATE_EXPOSURE_LEVELS,
  EXPOSURE_LEVEL_DEFS,
  TEMPLATE_DOMAIN_DISPLAY,
  PRIMARY_VERIFIED_TEMPLATE_IDS,
  GENERAL_VERIFIED_TEMPLATE_IDS,
  CONDITIONAL_TEMPLATE_IDS,
  QUARANTINED_TEMPLATE_IDS,
  TEMPLATE_EXPOSURE_OVERRIDES,
  getExposureLevelDef,
  getTemplateDomainDisplay,
  getTemplateExposureLevel,
  getTemplateExposure,
  getTemplateExposureBadges,
  buildTemplateExposureSummary,
};
