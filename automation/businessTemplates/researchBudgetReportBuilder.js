const {
  findTableForTemplate,
  executeTemplateSections,
  findColumnHeader,
  getRows,
  getRowValue,
  makeTemplateSection,
  toNumber,
} = require("./commonTemplateHelpers");
const {
  buildPeriodMetricReportSections,
  selectExecutionTable,
} = require("../structuralBuilders/periodMetricReportBuilder");
const {
  buildCategorySummaryReportSections,
} = require("../structuralBuilders/categorySummaryReportBuilder");

const RESEARCH_BUDGET_REPORT_VERSION = "research_budget_report_v2";

const RESEARCH_BUDGET_HINTS = {
  executionAmount: [
    "집행금액(현금",
    "현금집행금액",
    "집행금액현금",
    "집행금액",
    "집행액",
    "사용금액",
    "지출금액",
    "연구비집행액",
    "연구개발비집행액",
    "집행",
    "execution",
    "spent",
    "expense",
    "amount",
  ],
  cashExecution: [
    "집행금액(현금",
    "현금집행금액",
    "집행금액현금",
    "집행액현금",
    "현금집행액",
    "현금",
    "cash",
  ],
  inKindExecution: [
    "집행금액(현물",
    "현물집행금액",
    "집행금액현물",
    "집행액현물",
    "현물집행액",
    "현물",
    "in kind",
    "inkind",
  ],
  budgetAmount: [
    "총 연구비",
    "총연구비",
    "연구비총액",
    "연구개발비총액",
    "총사업비",
    "총액",
    "당해년도연구비",
    "당해년도 연구비",
    "예산액",
    "예산",
    "budget",
    "amount",
  ],
  governmentAmount: [
    "정부출연금",
    "정부 지원금",
    "정부지원금",
    "국고지원금",
    "정부연구비",
    "정부",
    "출연금",
    "government",
    "grant",
  ],
  privateAmount: ["민간부담금", "민간부담액", "기업부담금", "민간", "private"],
  metric: [
    "집행금액",
    "집행액",
    "사용금액",
    "지출금액",
    "총 연구비",
    "총연구비",
    "연구비",
    "연구개발비",
    "정부출연금",
    "정부지원금",
    "출연금",
    "예산",
    "금액",
    "amount",
    "budget",
  ],
  year: [
    "진행년도",
    "집행년도",
    "예산년도",
    "기준년도",
    "연도",
    "년도",
    "year",
  ],
  period: [
    "진행년도",
    "집행년도",
    "예산년도",
    "기준년도",
    "연도",
    "년도",
    "연구기간",
    "시작일",
    "종료일",
    "기간",
    "year",
    "date",
  ],
  expenseItem: [
    "항목명",
    "비목",
    "세목",
    "집행항목",
    "이용항목",
    "항목",
    "expense",
    "category",
  ],
  program: [
    "사업명",
    "내역사업명",
    "프로그램",
    "세부사업",
    "사업",
    "program",
    "business",
  ],
  project: [
    "과제명",
    "세부과제명",
    "연구과제명",
    "과제고유번호",
    "과제번호",
    "project",
  ],
  organizationType: [
    "기관분류",
    "기관유형",
    "기관구분",
    "기관종류",
    "구분",
    "분류",
    "유형",
    "organization type",
  ],
  agency: ["전문기관명", "전문기관", "관리기관", "agency"],
  organization: [
    "연구기관",
    "수행기관",
    "기관명",
    "기관",
    "주관기관",
    "organization",
    "institute",
  ],
  researcher: ["연구책임자", "책임자", "담당자", "pi", "researcher"],
};

function normalizeText(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function sameHeader(a = "", b = "") {
  return normalizeText(a) === normalizeText(b);
}

function findFirstColumnHeader(table = {}, hintGroups = []) {
  for (const hints of hintGroups) {
    const matched = findColumnHeader(table, hints || []);
    if (matched) return matched;
  }
  return "";
}

function resolveBudgetColumns(table = {}, config = {}) {
  const hints = config.hints || RESEARCH_BUDGET_HINTS;
  const cashExecutionHeader = findColumnHeader(
    table,
    hints.cashExecution || [],
  );
  const inKindExecutionHeader = findColumnHeader(
    table,
    hints.inKindExecution || [],
  );
  const executionHeader =
    cashExecutionHeader ||
    findFirstColumnHeader(table, [
      hints.executionAmount,
      ["집행금액", "집행액", "사용금액", "지출금액", "금액"],
    ]);
  const budgetHeader = findFirstColumnHeader(table, [
    hints.budgetAmount,
    ["총 연구비", "총연구비", "연구비", "예산", "금액"],
  ]);
  const governmentHeader = findColumnHeader(
    table,
    hints.governmentAmount || [],
  );
  const privateHeader = findColumnHeader(table, hints.privateAmount || []);
  const metricHeader =
    executionHeader ||
    budgetHeader ||
    governmentHeader ||
    findColumnHeader(table, hints.metric || []);

  const yearHeader = findColumnHeader(table, hints.year || []);
  const periodHeader =
    yearHeader || findColumnHeader(table, hints.period || []);
  const expenseItemHeader = findColumnHeader(table, hints.expenseItem || []);
  const programHeader = findColumnHeader(table, hints.program || []);
  const projectHeader = findColumnHeader(table, hints.project || []);
  const organizationTypeHeader = findColumnHeader(
    table,
    hints.organizationType || [],
  );
  const agencyHeader = findColumnHeader(table, hints.agency || []);
  const organizationHeader = findColumnHeader(table, hints.organization || []);
  const researcherHeader = findColumnHeader(table, hints.researcher || []);

  return {
    executionHeader,
    cashExecutionHeader,
    inKindExecutionHeader,
    budgetHeader,
    governmentHeader,
    privateHeader,
    metricHeader,
    yearHeader,
    periodHeader,
    expenseItemHeader,
    programHeader,
    projectHeader,
    organizationTypeHeader,
    agencyHeader,
    organizationHeader,
    researcherHeader,
  };
}

function normalizePeriodValue(value = "") {
  const raw = String(value ?? "").trim();
  if (!raw) return "";

  const year = raw.match(/(19|20)\d{2}/);
  if (year) return year[0];

  return raw;
}

function periodSortValue(value = "") {
  const raw = normalizePeriodValue(value);
  const year = raw.match(/(19|20)\d{2}/);
  if (year) return Number(year[0]);
  const n = Number(String(raw).replace(/,/g, ""));
  return Number.isFinite(n) ? n : raw;
}

function comparePeriodValues(a, b) {
  const av = periodSortValue(a);
  const bv = periodSortValue(b);
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  return String(av).localeCompare(String(bv), "ko");
}

function getDimensionValue(row = {}, header = "") {
  const value = getRowValue(row, header);
  const normalized = String(value ?? "").trim();
  return normalized || "미분류";
}

function addIfNumber(target = {}, field = "", value) {
  if (value == null) return false;
  target[field] = (target[field] || 0) + value;
  return true;
}

function getBudgetAmounts(row = {}, columns = {}) {
  const cash = columns.cashExecutionHeader
    ? toNumber(getRowValue(row, columns.cashExecutionHeader))
    : null;
  const inKind = columns.inKindExecutionHeader
    ? toNumber(getRowValue(row, columns.inKindExecutionHeader))
    : null;
  const execution = columns.executionHeader
    ? toNumber(getRowValue(row, columns.executionHeader))
    : null;
  const budget = columns.budgetHeader
    ? toNumber(getRowValue(row, columns.budgetHeader))
    : null;
  const government = columns.governmentHeader
    ? toNumber(getRowValue(row, columns.governmentHeader))
    : null;
  const privateAmount = columns.privateHeader
    ? toNumber(getRowValue(row, columns.privateHeader))
    : null;

  const hasCashInKind = cash != null || inKind != null;
  const executionTotal = hasCashInKind
    ? (cash || 0) + (inKind || 0)
    : execution;
  const primaryAmount =
    executionTotal != null
      ? executionTotal
      : budget != null
        ? budget
        : government;

  return {
    cash,
    inKind,
    execution: executionTotal,
    budget,
    government,
    privateAmount,
    primaryAmount,
  };
}

function aggregateBy({ table, dimensionHeader, columns }) {
  if (!dimensionHeader) return [];

  const map = new Map();
  for (const row of getRows(table)) {
    const key = getDimensionValue(row, dimensionHeader);
    const amounts = getBudgetAmounts(row, columns);
    if (amounts.primaryAmount == null && amounts.budget == null) continue;

    const current = map.get(key) || {
      label: key,
      집행액: 0,
      현금집행액: 0,
      현물집행액: 0,
      총연구비: 0,
      정부출연금: 0,
      민간부담금: 0,
      건수: 0,
    };

    addIfNumber(current, "집행액", amounts.execution);
    addIfNumber(current, "현금집행액", amounts.cash);
    addIfNumber(current, "현물집행액", amounts.inKind);
    addIfNumber(current, "총연구비", amounts.budget);
    addIfNumber(current, "정부출연금", amounts.government);
    addIfNumber(current, "민간부담금", amounts.privateAmount);
    current.건수 += 1;
    map.set(key, current);
  }

  return Array.from(map.values());
}

function primaryValue(row = {}) {
  if (row.집행액) return row.집행액;
  if (row.총연구비) return row.총연구비;
  if (row.정부출연금) return row.정부출연금;
  if (row.현금집행액) return row.현금집행액;
  return 0;
}

function makeBudgetSection({
  sectionId,
  sectionType,
  title,
  table,
  rows,
  columns = {},
  chartHint = {},
  narrativeHint = {},
}) {
  if (!Array.isArray(rows) || !rows.length) return null;

  return makeTemplateSection({
    sectionId,
    sectionType,
    title,
    candidate: {
      recipeType: "research_budget_report_v2_custom",
      sectionType,
      title,
      tableId: table.tableId,
      columns,
      meta: {
        researchBudgetReportVersion: RESEARCH_BUDGET_REPORT_VERSION,
        sectionType,
      },
    },
    result: {
      ok: true,
      recipeType: "research_budget_report_v2_custom",
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
      meta: {
        researchBudgetReportVersion: RESEARCH_BUDGET_REPORT_VERSION,
      },
    },
    chartHint,
    narrativeHint: {
      ...narrativeHint,
      researchBudgetReportVersion: RESEARCH_BUDGET_REPORT_VERSION,
    },
  });
}

function buildBudgetTrendSection({ table, columns }) {
  const { periodHeader } = columns;
  if (!periodHeader) return null;

  const map = new Map();
  for (const row of getRows(table)) {
    const period = normalizePeriodValue(getRowValue(row, periodHeader));
    if (!period) continue;

    const amounts = getBudgetAmounts(row, columns);
    if (amounts.primaryAmount == null && amounts.budget == null) continue;

    const current = map.get(period) || {
      기간: period,
      집행액: 0,
      현금집행액: 0,
      현물집행액: 0,
      총연구비: 0,
      정부출연금: 0,
      민간부담금: 0,
      과제수: 0,
    };
    addIfNumber(current, "집행액", amounts.execution);
    addIfNumber(current, "현금집행액", amounts.cash);
    addIfNumber(current, "현물집행액", amounts.inKind);
    addIfNumber(current, "총연구비", amounts.budget);
    addIfNumber(current, "정부출연금", amounts.government);
    addIfNumber(current, "민간부담금", amounts.privateAmount);
    current.과제수 += 1;
    map.set(period, current);
  }

  const rows = Array.from(map.values()).sort((a, b) =>
    comparePeriodValues(a.기간, b.기간),
  );

  let previous = null;
  let cumulative = 0;
  rows.forEach((row) => {
    const value = primaryValue(row);
    cumulative += value;
    row.누적금액 = cumulative;
    row.전기금액 = previous;
    row.전기대비증감 = previous == null ? null : value - previous;
    row.전기대비증감률 =
      previous == null || previous === 0 ? null : (value - previous) / previous;
    row.집행률 = row.총연구비 ? row.집행액 / row.총연구비 : null;
    previous = value;
  });

  return makeBudgetSection({
    sectionId: "research_budget_trend_growth_v2",
    sectionType: "research_budget_trend_growth",
    title: `${periodHeader}별 연구비/집행액 추이 및 증감률`,
    table,
    rows,
    columns: {
      date: periodHeader,
      metric: columns.metricHeader,
      execution: columns.executionHeader,
      budget: columns.budgetHeader,
      government: columns.governmentHeader,
    },
    chartHint: {
      preferredType: "line",
      categoryField: "기간",
      valueField: "집행액",
      secondaryValueFields: ["총연구비", "정부출연금", "누적금액", "집행률"],
    },
    narrativeHint: {
      focus: "research_budget_trend_growth",
      date: periodHeader,
    },
  });
}

function buildBudgetExecutionRateSection({ table, columns }) {
  if (!columns.executionHeader || !columns.budgetHeader) return null;

  const dimensionHeader =
    columns.programHeader ||
    columns.organizationTypeHeader ||
    columns.organizationHeader ||
    columns.expenseItemHeader ||
    "";
  if (!dimensionHeader) return null;

  const rows = aggregateBy({ table, dimensionHeader, columns })
    .filter((row) => row.집행액 || row.총연구비)
    .sort((a, b) => b.집행액 - a.집행액)
    .slice(0, 30)
    .map((row, index) => ({
      순위: index + 1,
      [dimensionHeader]: row.label,
      총연구비: row.총연구비,
      집행액: row.집행액,
      집행률: row.총연구비 ? row.집행액 / row.총연구비 : null,
      건수: row.건수,
    }));

  return makeBudgetSection({
    sectionId: "budget_execution_rate_v2",
    sectionType: "budget_execution_rate",
    title: `${dimensionHeader}별 연구비 집행률`,
    table,
    rows,
    columns: {
      dimension: dimensionHeader,
      metric: columns.executionHeader,
      budget: columns.budgetHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "집행액",
      secondaryValueFields: ["집행률"],
    },
    narrativeHint: {
      focus: "budget_execution_rate",
      dimension: dimensionHeader,
    },
  });
}

function buildCompositionSection({
  table,
  columns,
  dimensionHeader,
  sectionId,
  sectionType,
  label,
}) {
  if (!dimensionHeader || !columns.metricHeader) return null;

  const aggregated = aggregateBy({ table, dimensionHeader, columns })
    .sort((a, b) => primaryValue(b) - primaryValue(a))
    .slice(0, 30);

  const total = aggregated.reduce((sum, row) => sum + primaryValue(row), 0);
  if (!total) return null;

  const rows = aggregated.map((row, index) => {
    const value = primaryValue(row);
    return {
      순위: index + 1,
      [dimensionHeader]: row.label,
      금액합계: value,
      집행액: row.집행액,
      총연구비: row.총연구비,
      정부출연금: row.정부출연금,
      구성비: value / total,
      건수: row.건수,
    };
  });

  return makeBudgetSection({
    sectionId,
    sectionType,
    title: `${label}별 연구비 구성비`,
    table,
    rows,
    columns: {
      dimension: dimensionHeader,
      metric: columns.metricHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "금액합계",
      secondaryValueFields: ["구성비"],
    },
    narrativeHint: {
      focus: sectionType,
      dimension: dimensionHeader,
    },
  });
}

function buildBudgetCompositionSections({ table, columns }) {
  return [
    {
      header: columns.expenseItemHeader,
      id: "expense_item_budget_composition_v2",
      type: "expense_item_budget_composition",
      label: "비목/항목",
    },
    {
      header: columns.programHeader,
      id: "program_budget_composition_v2",
      type: "program_budget_composition",
      label: "사업",
    },
    {
      header: columns.organizationTypeHeader,
      id: "organization_type_budget_composition_v2",
      type: "organization_type_budget_composition",
      label: "기관유형",
    },
    {
      header: columns.agencyHeader,
      id: "agency_budget_composition_v2",
      type: "agency_budget_composition",
      label: "전문기관",
    },
    {
      header: columns.organizationHeader,
      id: "organization_budget_composition_v2",
      type: "organization_budget_composition",
      label: "연구/수행기관",
    },
  ]
    .map(({ header, id, type, label }) =>
      buildCompositionSection({
        table,
        columns,
        dimensionHeader: header,
        sectionId: id,
        sectionType: type,
        label,
      }),
    )
    .filter(Boolean);
}

function buildCashInKindSection({ table, columns }) {
  if (!columns.cashExecutionHeader || !columns.inKindExecutionHeader)
    return null;

  const dimensionHeader =
    columns.expenseItemHeader ||
    columns.organizationTypeHeader ||
    columns.periodHeader;
  if (!dimensionHeader) return null;

  const rows = aggregateBy({ table, dimensionHeader, columns })
    .filter((row) => row.현금집행액 || row.현물집행액)
    .sort((a, b) => b.현금집행액 + b.현물집행액 - (a.현금집행액 + a.현물집행액))
    .slice(0, 30)
    .map((row, index) => {
      const total = row.현금집행액 + row.현물집행액;
      return {
        순위: index + 1,
        [dimensionHeader]: row.label,
        현금집행액: row.현금집행액,
        현물집행액: row.현물집행액,
        집행액합계: total,
        현물비중: total ? row.현물집행액 / total : null,
        건수: row.건수,
      };
    });

  return makeBudgetSection({
    sectionId: "cash_inkind_execution_v2",
    sectionType: "cash_inkind_execution",
    title: `${dimensionHeader}별 현금/현물 집행 구성`,
    table,
    rows,
    columns: {
      dimension: dimensionHeader,
      cash: columns.cashExecutionHeader,
      inKind: columns.inKindExecutionHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "현금집행액",
      secondaryValueFields: ["현물집행액", "현물비중"],
    },
    narrativeHint: {
      focus: "cash_inkind_execution",
      dimension: dimensionHeader,
    },
  });
}

function buildTopProjectBudgetSection({ table, columns }) {
  const dimensionHeader =
    columns.projectHeader ||
    columns.programHeader ||
    columns.organizationHeader;
  if (!dimensionHeader || !columns.metricHeader) return null;

  const aggregated = aggregateBy({ table, dimensionHeader, columns })
    .filter((row) => primaryValue(row) !== 0)
    .sort((a, b) => primaryValue(b) - primaryValue(a));
  if (!aggregated.length) return null;

  const top = aggregated.slice(0, 10).map((row, index) => ({
    구분: "상위",
    순위: index + 1,
    [dimensionHeader]: row.label,
    금액합계: primaryValue(row),
    집행액: row.집행액,
    총연구비: row.총연구비,
    정부출연금: row.정부출연금,
    건수: row.건수,
  }));
  const bottom = aggregated
    .slice(-10)
    .reverse()
    .map((row, index) => ({
      구분: "하위",
      순위: index + 1,
      [dimensionHeader]: row.label,
      금액합계: primaryValue(row),
      집행액: row.집행액,
      총연구비: row.총연구비,
      정부출연금: row.정부출연금,
      건수: row.건수,
    }));

  return makeBudgetSection({
    sectionId: "top_bottom_research_budget_v2",
    sectionType: "top_bottom_research_budget",
    title: `${dimensionHeader} 기준 연구비 상위/하위`,
    table,
    rows: [...top, ...bottom],
    columns: {
      dimension: dimensionHeader,
      metric: columns.metricHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "금액합계",
    },
    narrativeHint: {
      focus: "top_bottom_research_budget",
      dimension: dimensionHeader,
    },
  });
}

function hasSameSection(sections = [], sectionId = "") {
  if (!sectionId) return false;
  return sections.some((section) => section.sectionId === sectionId);
}

function uniqueSections(sections = []) {
  const seen = new Set();
  return sections.filter((section) => {
    if (!section) return false;
    const key = [
      section.sectionId,
      section.sectionType,
      section.title,
      section.candidate?.columns?.dimension,
      section.candidate?.columns?.date,
      section.candidate?.columns?.metric,
    ].join("|");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildResearchBudgetV2Sections({
  normalizedQueryTables,
  table,
  templateCandidate,
  config,
}) {
  const { selectedTable } = selectExecutionTable({
    normalizedQueryTables,
    table,
    hints: config.hints || RESEARCH_BUDGET_HINTS,
  });
  const columns = resolveBudgetColumns(selectedTable, config);

  const sections = [
    buildBudgetTrendSection({ table: selectedTable, columns }),
    buildBudgetExecutionRateSection({ table: selectedTable, columns }),
    ...buildBudgetCompositionSections({ table: selectedTable, columns }),
    buildCashInKindSection({ table: selectedTable, columns }),
    buildTopProjectBudgetSection({ table: selectedTable, columns }),
  ].filter(Boolean);

  return uniqueSections(sections).map((section) => ({
    ...section,
    candidate: {
      ...section.candidate,
      templateId: templateCandidate.templateId || "research_budget_report",
    },
  }));
}

function executionConfig() {
  return {
    hints: {
      ...RESEARCH_BUDGET_HINTS,
      metric: RESEARCH_BUDGET_HINTS.executionAmount,
      year: RESEARCH_BUDGET_HINTS.year,
      period: RESEARCH_BUDGET_HINTS.period,
      item: RESEARCH_BUDGET_HINTS.expenseItem,
      category: RESEARCH_BUDGET_HINTS.organizationType,
    },
    sectionIds: {
      year: "yearly_execution_trend",
      period: "period_execution_trend",
      topBottom: "top_bottom_execution",
    },
    sectionTypes: {
      year: "yearly_execution_trend",
      period: "period_execution_trend",
      topBottom: "top_bottom_execution",
    },
    dimensions: [
      {
        sectionId: "expense_category_execution",
        sectionType: "expense_category_execution",
        hints: RESEARCH_BUDGET_HINTS.expenseItem,
      },
      {
        sectionId: "organization_type_execution",
        sectionType: "organization_type_execution",
        hints: RESEARCH_BUDGET_HINTS.organizationType,
      },
      {
        sectionId: "agency_execution",
        sectionType: "agency_execution",
        hints: RESEARCH_BUDGET_HINTS.agency,
      },
      {
        sectionId: "organization_execution",
        sectionType: "organization_execution",
        hints: RESEARCH_BUDGET_HINTS.organization,
      },
    ],
    rankingDimensionHints: [
      ...RESEARCH_BUDGET_HINTS.expenseItem,
      ...RESEARCH_BUDGET_HINTS.agency,
      ...RESEARCH_BUDGET_HINTS.organization,
      ...RESEARCH_BUDGET_HINTS.program,
    ],
  };
}

function allocatedBudgetConfig() {
  return {
    hints: {
      ...RESEARCH_BUDGET_HINTS,
      metric: [
        ...RESEARCH_BUDGET_HINTS.budgetAmount,
        ...RESEARCH_BUDGET_HINTS.governmentAmount,
      ],
      year: RESEARCH_BUDGET_HINTS.year,
      period: RESEARCH_BUDGET_HINTS.period,
      item: RESEARCH_BUDGET_HINTS.project,
      category: RESEARCH_BUDGET_HINTS.program,
    },
    sectionIds: {
      year: "yearly_budget_trend",
      period: "period_budget_trend",
      topBottom: "top_bottom_project_budget",
    },
    sectionTypes: {
      year: "yearly_budget_trend",
      period: "period_budget_trend",
      topBottom: "top_bottom_project_budget",
    },
    dimensions: [
      {
        sectionId: "program_budget",
        sectionType: "program_budget",
        hints: RESEARCH_BUDGET_HINTS.program,
      },
      {
        sectionId: "project_budget",
        sectionType: "project_budget",
        hints: RESEARCH_BUDGET_HINTS.project,
      },
      {
        sectionId: "organization_type_budget",
        sectionType: "organization_type_budget",
        hints: RESEARCH_BUDGET_HINTS.organizationType,
      },
      {
        sectionId: "organization_budget",
        sectionType: "organization_budget",
        hints: RESEARCH_BUDGET_HINTS.organization,
      },
      {
        sectionId: "researcher_budget",
        sectionType: "researcher_budget",
        hints: RESEARCH_BUDGET_HINTS.researcher,
      },
    ],
    rankingDimensionHints: [
      ...RESEARCH_BUDGET_HINTS.project,
      ...RESEARCH_BUDGET_HINTS.program,
      ...RESEARCH_BUDGET_HINTS.organization,
    ],
  };
}

function buildResearchBudgetCategoryFallbackSections({
  normalizedQueryTables,
  table,
  templateCandidate,
}) {
  return buildCategorySummaryReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: {
      metricHints: RESEARCH_BUDGET_HINTS.metric,
      dimensions: [
        {
          sectionId: "expense_category_summary",
          sectionType: "expense_category_summary",
          hints: RESEARCH_BUDGET_HINTS.expenseItem,
        },
        {
          sectionId: "program_summary",
          sectionType: "program_summary",
          hints: RESEARCH_BUDGET_HINTS.program,
        },
        {
          sectionId: "organization_summary",
          sectionType: "organization_summary",
          hints: [
            ...RESEARCH_BUDGET_HINTS.organizationType,
            ...RESEARCH_BUDGET_HINTS.organization,
            ...RESEARCH_BUDGET_HINTS.agency,
          ],
        },
      ],
      topBottom: {
        sectionId: "research_budget_top_bottom",
        sectionType: "research_budget_top_bottom",
        dimensionHints: [
          ...RESEARCH_BUDGET_HINTS.project,
          ...RESEARCH_BUDGET_HINTS.expenseItem,
          ...RESEARCH_BUDGET_HINTS.organization,
          ...RESEARCH_BUDGET_HINTS.program,
        ],
      },
    },
  });
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

  const execution = executionConfig();
  const allocated = allocatedBudgetConfig();

  const v2Sections = buildResearchBudgetV2Sections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: { hints: RESEARCH_BUDGET_HINTS },
  });

  const executionSections = buildPeriodMetricReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: execution,
  });

  const allocatedSections = buildPeriodMetricReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: allocated,
  });

  const categoryFallbackSections = buildResearchBudgetCategoryFallbackSections({
    normalizedQueryTables,
    table,
    templateCandidate,
  });

  const sections = uniqueSections([
    ...v2Sections,
    ...executionSections.filter(
      (section) => !hasSameSection(v2Sections, section.sectionId),
    ),
    ...allocatedSections.filter(
      (section) => !hasSameSection(v2Sections, section.sectionId),
    ),
    ...categoryFallbackSections.filter(
      (section) => !hasSameSection(v2Sections, section.sectionId),
    ),
  ]).map((section) => ({
    ...section,
    meta: {
      ...(section.meta || {}),
      researchBudgetReportVersion: RESEARCH_BUDGET_REPORT_VERSION,
    },
  }));

  if (!sections.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  return sections;
}

module.exports = {
  RESEARCH_BUDGET_REPORT_VERSION,
  RESEARCH_BUDGET_HINTS,
  executeResearchBudgetReport,
  buildResearchBudgetV2Sections,
  resolveBudgetColumns,
};
