const {
  findTableForTemplate,
  executeTemplateSections,
} = require("./commonTemplateHelpers");
const {
  buildPeriodMetricReportSections,
} = require("../structuralBuilders/periodMetricReportBuilder");
const {
  buildCategorySummaryReportSections,
} = require("../structuralBuilders/categorySummaryReportBuilder");

function executionConfig() {
  return {
    hints: {
      metric: [
        "집행금액(현금",
        "현금집행금액",
        "집행금액현금",
        "집행액현금",
        "집행금액",
        "집행액",
        "사용금액",
        "지출금액",
        "연구비집행액",
        "amount",
      ],
      year: ["진행년도", "집행년도", "연도", "예산년도", "년도", "year"],
      period: ["진행년도", "집행년도", "연도", "예산년도", "년도", "year"],
      item: [
        "항목명",
        "비목",
        "세목",
        "집행항목",
        "이용항목",
        "항목",
        "expense",
        "category",
      ],
      category: ["기관분류", "기관유형", "기관구분", "분류", "유형"],
    },
    sectionIds: {
      year: "yearly_execution_trend",
      topBottom: "top_bottom_execution",
    },
    sectionTypes: {
      year: "yearly_execution_trend",
      topBottom: "top_bottom_execution",
    },
    dimensions: [
      {
        sectionId: "expense_category_execution",
        sectionType: "expense_category_execution",
        hints: [
          "항목명",
          "비목",
          "세목",
          "집행항목",
          "이용항목",
          "항목",
          "expense",
          "category",
        ],
      },
      {
        sectionId: "organization_type_execution",
        sectionType: "organization_type_execution",
        hints: ["기관분류", "기관유형", "기관구분", "분류", "유형"],
      },
      {
        sectionId: "agency_execution",
        sectionType: "agency_execution",
        hints: [
          "전문기관명",
          "기관명",
          "연구기관",
          "기관",
          "수행기관",
          "agency",
          "organization",
        ],
      },
    ],
    rankingDimensionHints: [
      "항목명",
      "비목",
      "세목",
      "집행항목",
      "전문기관명",
      "기관명",
    ],
  };
}

function allocatedBudgetConfig() {
  return {
    hints: {
      metric: [
        "총 연구비",
        "총연구비",
        "연구비총액",
        "연구개발비총액",
        "총사업비",
        "총액",
        "정부출연금",
        "정부지원금",
        "국고지원금",
        "출연금",
        "정부",
      ],
      year: ["예산년도", "연도", "년도", "year"],
      period: ["예산년도", "연도", "년도", "year"],
      item: ["과제명", "세부과제명", "연구과제명", "project"],
      category: ["사업명", "내역사업명", "프로그램", "program", "business"],
    },
    sectionIds: {
      year: "yearly_budget_trend",
      topBottom: "top_bottom_project_budget",
    },
    sectionTypes: {
      year: "yearly_budget_trend",
      topBottom: "top_bottom_project_budget",
    },
    dimensions: [
      {
        sectionId: "program_budget",
        sectionType: "program_budget",
        hints: ["사업명", "내역사업명", "프로그램", "program", "business"],
      },
      {
        sectionId: "organization_type_budget",
        sectionType: "organization_type_budget",
        hints: ["기관유형", "기관분류", "기관구분", "유형", "분류"],
      },
      {
        sectionId: "organization_budget",
        sectionType: "organization_budget",
        hints: ["연구기관", "수행기관", "기관명", "기관", "organization"],
      },
      {
        sectionId: "researcher_budget",
        sectionType: "researcher_budget",
        hints: ["연구책임자", "책임자", "담당자", "pi"],
      },
    ],
    rankingDimensionHints: ["과제명", "세부과제명", "연구과제명", "사업명"],
  };
}

function executeResearchBudgetReport({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const table = findTableForTemplate(normalizedQueryTables, templateCandidate);

  if (!table?.tableId) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const executionSections = buildPeriodMetricReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: executionConfig(),
  });

  const allocatedSections = buildPeriodMetricReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: allocatedBudgetConfig(),
  });

  const categoryFallbackSections = buildCategorySummaryReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: {
      metricHints: ["집행금액", "집행액", "총 연구비", "정부출연금", "금액"],
      dimensions: [
        {
          sectionId: "expense_category_summary",
          sectionType: "expense_category_summary",
          hints: ["항목명", "비목", "세목", "항목"],
        },
        {
          sectionId: "organization_summary",
          sectionType: "organization_summary",
          hints: ["기관분류", "기관유형", "연구기관", "기관명", "기관"],
        },
      ],
      topBottom: {
        sectionId: "research_budget_top_bottom",
        sectionType: "research_budget_top_bottom",
        dimensionHints: ["과제명", "항목명", "기관명", "사업명"],
      },
    },
  });

  const sections =
    executionSections.length >= 2
      ? executionSections
      : allocatedSections.length >= 2
        ? allocatedSections
        : [
            ...executionSections,
            ...allocatedSections,
            ...categoryFallbackSections,
          ];

  if (!sections.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  return sections;
}

module.exports = {
  executeResearchBudgetReport,
};
