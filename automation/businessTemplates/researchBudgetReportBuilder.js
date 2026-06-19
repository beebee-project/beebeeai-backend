const {
  findTableForTemplate,
  findColumnHeader,
  makeTemplateCandidate,
  executeTemplateSections,
} = require("./commonTemplateHelpers");

function buildExecutionBudgetCandidates({ table }) {
  const tableId = table.tableId;

  const yearHeader = findColumnHeader(table, [
    "진행년도",
    "집행년도",
    "연도",
    "예산년도",
    "년도",
    "year",
  ]);

  const agencyHeader = findColumnHeader(table, [
    "전문기관명",
    "기관명",
    "연구기관",
    "기관",
    "수행기관",
    "agency",
    "organization",
  ]);

  const orgTypeHeader = findColumnHeader(table, [
    "기관분류",
    "기관유형",
    "기관구분",
    "분류",
    "유형",
  ]);

  const expenseCategoryHeader = findColumnHeader(table, [
    "항목명",
    "비목",
    "세목",
    "집행항목",
    "이용항목",
    "항목",
    "expense",
    "category",
  ]);

  const cashAmountHeader = findColumnHeader(
    table,
    ["집행금액(현금", "현금집행금액", "집행금액현금", "집행액현금", "현금"],
    { type: "number" },
  );

  const inKindAmountHeader = findColumnHeader(
    table,
    ["집행금액(현물", "현물집행금액", "집행금액현물", "집행액현물", "현물"],
    { type: "number" },
  );

  const executionAmountHeader =
    cashAmountHeader ||
    findColumnHeader(
      table,
      ["집행금액", "집행액", "사용금액", "지출금액", "연구비집행액", "amount"],
      { type: "number" },
    );

  const candidates = [];

  if (expenseCategoryHeader && executionAmountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${expenseCategoryHeader}별 ${executionAmountHeader} 요약`,
        tableId,
        columns: {
          dimension: expenseCategoryHeader,
          metric: executionAmountHeader,
        },
        meta: {
          sectionType: "expense_category_execution",
        },
      }),
    );
  }

  if (orgTypeHeader && executionAmountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${orgTypeHeader}별 ${executionAmountHeader} 요약`,
        tableId,
        columns: {
          dimension: orgTypeHeader,
          metric: executionAmountHeader,
        },
        meta: {
          sectionType: "organization_type_execution",
        },
      }),
    );
  }

  if (agencyHeader && executionAmountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${agencyHeader}별 ${executionAmountHeader} 요약`,
        tableId,
        columns: {
          dimension: agencyHeader,
          metric: executionAmountHeader,
        },
        meta: {
          sectionType: "agency_execution",
        },
      }),
    );
  }

  if (yearHeader && executionAmountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${yearHeader}별 ${executionAmountHeader} 추이`,
        tableId,
        columns: {
          dimension: yearHeader,
          metric: executionAmountHeader,
        },
        meta: {
          sectionType: "yearly_execution_trend",
        },
      }),
    );
  }

  if (expenseCategoryHeader && inKindAmountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${expenseCategoryHeader}별 ${inKindAmountHeader} 요약`,
        tableId,
        columns: {
          dimension: expenseCategoryHeader,
          metric: inKindAmountHeader,
        },
        meta: {
          sectionType: "expense_category_in_kind_execution",
        },
      }),
    );
  }

  if (expenseCategoryHeader && executionAmountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "top_bottom",
        title: `${executionAmountHeader} 상위/하위 항목`,
        tableId,
        columns: {
          dimension: expenseCategoryHeader,
          metric: executionAmountHeader,
        },
        meta: {
          sectionType: "top_bottom_execution",
        },
      }),
    );
  }

  return candidates;
}

function buildAllocatedBudgetCandidates({ table }) {
  const tableId = table.tableId;

  const yearHeader = findColumnHeader(table, [
    "예산년도",
    "연도",
    "년도",
    "year",
  ]);

  const projectHeader = findColumnHeader(table, [
    "과제명",
    "세부과제명",
    "연구과제명",
    "project",
  ]);

  const programHeader = findColumnHeader(table, [
    "사업명",
    "내역사업명",
    "프로그램",
    "program",
    "business",
  ]);

  const organizationHeader = findColumnHeader(table, [
    "연구기관",
    "수행기관",
    "기관명",
    "기관",
    "organization",
  ]);

  const orgTypeHeader = findColumnHeader(table, [
    "기관유형",
    "기관분류",
    "기관구분",
    "유형",
    "분류",
  ]);

  const researcherHeader = findColumnHeader(table, [
    "연구책임자",
    "책임자",
    "담당자",
    "pi",
  ]);

  const totalBudgetHeader = findColumnHeader(
    table,
    [
      "총 연구비",
      "총연구비",
      "연구비총액",
      "연구개발비총액",
      "총사업비",
      "총액",
    ],
    { type: "number" },
  );

  const governmentBudgetHeader = findColumnHeader(
    table,
    ["정부출연금", "정부지원금", "국고지원금", "출연금", "정부"],
    { type: "number" },
  );

  const budgetMetricHeader = totalBudgetHeader || governmentBudgetHeader;

  const candidates = [];

  if (programHeader && budgetMetricHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${programHeader}별 ${budgetMetricHeader} 요약`,
        tableId,
        columns: {
          dimension: programHeader,
          metric: budgetMetricHeader,
        },
        meta: {
          sectionType: "program_budget",
        },
      }),
    );
  }

  if (orgTypeHeader && budgetMetricHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${orgTypeHeader}별 ${budgetMetricHeader} 요약`,
        tableId,
        columns: {
          dimension: orgTypeHeader,
          metric: budgetMetricHeader,
        },
        meta: {
          sectionType: "organization_type_budget",
        },
      }),
    );
  }

  if (organizationHeader && budgetMetricHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${organizationHeader}별 ${budgetMetricHeader} 요약`,
        tableId,
        columns: {
          dimension: organizationHeader,
          metric: budgetMetricHeader,
        },
        meta: {
          sectionType: "organization_budget",
        },
      }),
    );
  }

  if (yearHeader && budgetMetricHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${yearHeader}별 ${budgetMetricHeader} 추이`,
        tableId,
        columns: {
          dimension: yearHeader,
          metric: budgetMetricHeader,
        },
        meta: {
          sectionType: "yearly_budget_trend",
        },
      }),
    );
  }

  if (projectHeader && budgetMetricHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "top_bottom",
        title: `${budgetMetricHeader} 상위/하위 과제`,
        tableId,
        columns: {
          dimension: projectHeader,
          metric: budgetMetricHeader,
        },
        meta: {
          sectionType: "top_bottom_project_budget",
        },
      }),
    );
  }

  if (researcherHeader && budgetMetricHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${researcherHeader}별 ${budgetMetricHeader} 요약`,
        tableId,
        columns: {
          dimension: researcherHeader,
          metric: budgetMetricHeader,
        },
        meta: {
          sectionType: "researcher_budget",
        },
      }),
    );
  }

  if (governmentBudgetHeader && totalBudgetHeader && programHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${programHeader}별 ${governmentBudgetHeader} 요약`,
        tableId,
        columns: {
          dimension: programHeader,
          metric: governmentBudgetHeader,
        },
        meta: {
          sectionType: "program_government_budget",
        },
      }),
    );
  }

  return candidates;
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

  const executionCandidates = buildExecutionBudgetCandidates({ table });
  const allocatedCandidates = buildAllocatedBudgetCandidates({ table });

  let candidates = [];

  if (executionCandidates.length >= 2) {
    candidates = executionCandidates;
  } else if (allocatedCandidates.length >= 2) {
    candidates = allocatedCandidates;
  } else {
    candidates = [...executionCandidates, ...allocatedCandidates];
  }

  if (!candidates.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  return executeTemplateSections({
    normalizedQueryTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });
}

module.exports = {
  executeResearchBudgetReport,
};
