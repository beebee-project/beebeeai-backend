const TEMPLATE_DOMAIN_VERSION = "template_domain_config_v1";

const TEMPLATE_DOMAINS = Object.freeze({
  SALES: "sales",
  BUDGET: "budget",
  HR: "hr",
  INVENTORY: "inventory",
  SURVEY: "survey",
  OPERATION: "operation",
  CUSTOMER: "customer",
  PROJECT: "project",
  PUBLIC_DATA: "publicData",
  GENERAL: "general",
});

const IMPLEMENTATION_LEVELS = Object.freeze({
  CUSTOM: "custom",
  DEFINITION_ONLY: "definitionOnly",
  STATUS_RATE: "statusRate",
  SURVEY_SCORE: "surveyScore",
  INVENTORY_FLOW: "inventoryFlow",
  COMPOSITE: "composite",
});

const TEMPLATE_DOMAIN_DEFS = Object.freeze({
  [TEMPLATE_DOMAINS.SALES]: Object.freeze({
    domain: TEMPLATE_DOMAINS.SALES,
    label: "매출·영업",
    group: "businessPerformance",
    priorityBand: 90,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.DEFINITION_ONLY,
    description: "매출, 매입, 판매수량, 지점·지역·상품·거래처별 실적 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.BUDGET]: Object.freeze({
    domain: TEMPLATE_DOMAINS.BUDGET,
    label: "예산·지출·정산",
    group: "financeOps",
    priorityBand: 80,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.DEFINITION_ONLY,
    description: "연구비, 예산, 지출, 정산, 카드, 출장비, 구매·계약 집행 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.HR]: Object.freeze({
    domain: TEMPLATE_DOMAINS.HR,
    label: "인사·조직·명단",
    group: "peopleOps",
    priorityBand: 75,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.DEFINITION_ONLY,
    description: "직원·참여자 명단, 부서·직급·상태, 입퇴사, 교육, 평가, 급여 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.INVENTORY]: Object.freeze({
    domain: TEMPLATE_DOMAINS.INVENTORY,
    label: "재고·물류·운영",
    group: "operationFlow",
    priorityBand: 55,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.INVENTORY_FLOW,
    description: "재고, 입출고, 발주, 배송, 장비·비품, 생산·품질·안전 운영 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.SURVEY]: Object.freeze({
    domain: TEMPLATE_DOMAINS.SURVEY,
    label: "설문·평가·만족도",
    group: "feedbackScore",
    priorityBand: 50,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.SURVEY_SCORE,
    description: "설문, 만족도, 평가 점수, 문항별 평균·분포 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.OPERATION]: Object.freeze({
    domain: TEMPLATE_DOMAINS.OPERATION,
    label: "운영·상태 관리",
    group: "operationStatus",
    priorityBand: 60,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.STATUS_RATE,
    description: "교육 이수, 배송, 유지보수, 안전점검, 처리 상태와 완료율 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.CUSTOMER]: Object.freeze({
    domain: TEMPLATE_DOMAINS.CUSTOMER,
    label: "고객·민원·CS",
    group: "customerOps",
    priorityBand: 60,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.DEFINITION_ONLY,
    description: "고객 문의, 민원, 처리상태, 유형별·기간별 건수 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.PROJECT]: Object.freeze({
    domain: TEMPLATE_DOMAINS.PROJECT,
    label: "프로젝트·성과지표",
    group: "projectOps",
    priorityBand: 58,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.STATUS_RATE,
    description: "과제·프로젝트 상태, 담당자별 건수, 진행률, 목표 대비 성과 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.PUBLIC_DATA]: Object.freeze({
    domain: TEMPLATE_DOMAINS.PUBLIC_DATA,
    label: "공공데이터·통계표",
    group: "publicStatistics",
    priorityBand: 65,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.DEFINITION_ONLY,
    description: "공공 통계표, wide/month table, 다중 시트, 다단 헤더 정리와 추이 분석 영역입니다.",
  }),
  [TEMPLATE_DOMAINS.GENERAL]: Object.freeze({
    domain: TEMPLATE_DOMAINS.GENERAL,
    label: "일반 자동화",
    group: "general",
    priorityBand: 40,
    defaultImplementationLevel: IMPLEMENTATION_LEVELS.DEFINITION_ONLY,
    description: "특정 업무 도메인으로 분류되지 않는 일반 요약·추이·구성비 분석 영역입니다.",
  }),
});

const IMPLEMENTATION_LEVEL_DEFS = Object.freeze({
  [IMPLEMENTATION_LEVELS.CUSTOM]: Object.freeze({
    level: IMPLEMENTATION_LEVELS.CUSTOM,
    label: "커스텀 executor",
    description: "전용 businessTemplate executor가 있어 고품질 섹션을 직접 생성합니다.",
  }),
  [IMPLEMENTATION_LEVELS.DEFINITION_ONLY]: Object.freeze({
    level: IMPLEMENTATION_LEVELS.DEFINITION_ONLY,
    label: "정의/힌트 기반",
    description: "별도 executor 없이 일반 분석 후보와 공통 템플릿 실행 흐름을 재사용합니다.",
  }),
  [IMPLEMENTATION_LEVELS.STATUS_RATE]: Object.freeze({
    level: IMPLEMENTATION_LEVELS.STATUS_RATE,
    label: "상태·율 공통 builder",
    description: "이수율, 취소율, 완료율, 진행률 등 상태 기반 비율 분석 builder가 필요합니다.",
  }),
  [IMPLEMENTATION_LEVELS.SURVEY_SCORE]: Object.freeze({
    level: IMPLEMENTATION_LEVELS.SURVEY_SCORE,
    label: "설문·점수 공통 builder",
    description: "문항별 평균, 응답 분포, 만족도/평가점수 분석 builder가 필요합니다.",
  }),
  [IMPLEMENTATION_LEVELS.INVENTORY_FLOW]: Object.freeze({
    level: IMPLEMENTATION_LEVELS.INVENTORY_FLOW,
    label: "재고·흐름 공통 builder",
    description: "입고/출고/현재고/누적 흐름 분석 builder가 필요합니다.",
  }),
  [IMPLEMENTATION_LEVELS.COMPOSITE]: Object.freeze({
    level: IMPLEMENTATION_LEVELS.COMPOSITE,
    label: "복합 조합형",
    description: "여러 공통 builder와 일반 recipe를 조합해 생성합니다.",
  }),
});

function normalizeDomain(domain = "") {
  const raw = String(domain || "").trim();
  return Object.prototype.hasOwnProperty.call(TEMPLATE_DOMAIN_DEFS, raw)
    ? raw
    : TEMPLATE_DOMAINS.GENERAL;
}

function normalizeImplementationLevel(level = "", domain = "") {
  const raw = String(level || "").trim();
  if (Object.prototype.hasOwnProperty.call(IMPLEMENTATION_LEVEL_DEFS, raw)) {
    return raw;
  }

  const domainDef = TEMPLATE_DOMAIN_DEFS[normalizeDomain(domain)];
  return domainDef?.defaultImplementationLevel || IMPLEMENTATION_LEVELS.DEFINITION_ONLY;
}

function getTemplateDomainDef(domain = "") {
  return TEMPLATE_DOMAIN_DEFS[normalizeDomain(domain)];
}

function getImplementationLevelDef(level = "", domain = "") {
  const normalized = normalizeImplementationLevel(level, domain);
  return IMPLEMENTATION_LEVEL_DEFS[normalized];
}

function enrichTemplateDefinition(def = {}) {
  const domain = normalizeDomain(def.domain);
  const domainDef = getTemplateDomainDef(domain);
  const implementationLevel = normalizeImplementationLevel(
    def.implementationLevel,
    domain,
  );
  const implementationDef = getImplementationLevelDef(implementationLevel, domain);

  return {
    ...def,
    domain,
    domainLabel: def.domainLabel || domainDef.label,
    domainGroup: def.domainGroup || domainDef.group,
    implementationLevel,
    implementationLevelLabel:
      def.implementationLevelLabel || implementationDef.label,
    preferredRecipeTypes: Array.isArray(def.preferredRecipeTypes)
      ? def.preferredRecipeTypes
      : Array.isArray(def.requiredAnyRecipeTypes)
        ? def.requiredAnyRecipeTypes
        : [],
    templateTags: Array.isArray(def.templateTags) ? def.templateTags : [],
    templateDomainVersion: TEMPLATE_DOMAIN_VERSION,
  };
}

function getTemplateDomainDefinitions() {
  return TEMPLATE_DOMAIN_DEFS;
}

function getImplementationLevelDefinitions() {
  return IMPLEMENTATION_LEVEL_DEFS;
}

module.exports = {
  TEMPLATE_DOMAIN_VERSION,
  TEMPLATE_DOMAINS,
  IMPLEMENTATION_LEVELS,
  TEMPLATE_DOMAIN_DEFS,
  IMPLEMENTATION_LEVEL_DEFS,
  normalizeDomain,
  normalizeImplementationLevel,
  getTemplateDomainDef,
  getImplementationLevelDef,
  enrichTemplateDefinition,
  getTemplateDomainDefinitions,
  getImplementationLevelDefinitions,
};
