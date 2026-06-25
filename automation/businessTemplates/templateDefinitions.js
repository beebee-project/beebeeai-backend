const BUSINESS_TEMPLATE_DEFS = [
  {
    templateId: "sales_report",
    title: "매출 분석 보고서",
    description:
      "기간별 매출, 수량, 상위 항목, 평균 판매금액을 보고서 형태로 생성합니다.",
    requiredAnyRecipeTypes: ["time_trend", "group_summary", "top_bottom"],
    requiredAnyHeaderHints: [
      "매출",
      "순매출액",
      "판매",
      "수량",
      "카드매출",
      "revenue",
      "sales",
    ],
    optionalHeaderHints: [
      "연도",
      "월",
      "연월",
      "제품",
      "상품",
      "지역",
      "업종",
      "거래처",
    ],
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    priority: 100,
  },
  {
    templateId: "research_budget_report",
    title: "연구비 집행 현황",
    description:
      "연구비, 집행액, 항목별·기관별·연도별 현황을 보고서 형태로 생성합니다.",
    requiredAnyRecipeTypes: ["group_summary", "time_trend", "top_bottom"],
    requiredAnyHeaderHints: [
      "연구비",
      "집행",
      "정부출연금",
      "과제",
      "항목명",
      "기관분류",
      "전문기관",
    ],
    optionalHeaderHints: [
      "예산",
      "현금",
      "현물",
      "사업명",
      "연구기관",
      "연구책임자",
      "진행년도",
      "예산년도",
    ],
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    priority: 90,
  },
  {
    templateId: "hr_monthly_report",
    title: "월간 인사 보고서",
    description:
      "명단, 상태, 부서별·직급별 현황, 입사 추이, 연봉 요약을 보고서 형태로 생성합니다.",
    requiredAnyRecipeTypes: ["category_count", "group_summary", "top_bottom"],
    requiredAnyHeaderHints: [
      "부서",
      "소속",
      "직급",
      "직위",
      "재직",
      "상태",
      "성명",
      "이름",
      "입사",
      "연봉",
      "급여",
    ],
    optionalHeaderHints: [
      "직원",
      "사원",
      "인사",
      "평가",
      "조직",
      "팀",
      "근무",
      "퇴사",
    ],
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    priority: 80,
  },
];

function getBusinessTemplateDefinitions() {
  return BUSINESS_TEMPLATE_DEFS;
}

function findBusinessTemplateDefinition(templateId = "") {
  return (
    BUSINESS_TEMPLATE_DEFS.find((def) => def.templateId === templateId) || null
  );
}

module.exports = {
  BUSINESS_TEMPLATE_DEFS,
  getBusinessTemplateDefinitions,
  findBusinessTemplateDefinition,
};
