const {
  findColumnHeader,
  getColumnHeader,
  getColumns,
  headerMatches,
  makeTemplateCandidate,
  makeTemplateSection,
  executeTemplateSections,
  getRows,
  getRowValue,
  createVirtualTable,
  toNumber,
} = require("../businessTemplates/commonTemplateHelpers");

function isYearHeader(header = "") {
  return /^(19|20)\d{2}\s*년?$/.test(String(header || "").trim());
}

function isMonthHeader(header = "") {
  return /^(0?[1-9]|1[0-2])\s*월/.test(String(header || "").trim());
}

function normalizeYear(header = "") {
  const match = String(header || "").match(/(19|20)\d{2}/);
  return match ? match[0] : "";
}

function normalizeMonth(header = "") {
  const match = String(header || "").match(/(0?[1-9]|1[0-2])\s*월/);
  if (!match) return "";
  return String(Number(match[1])).padStart(2, "0");
}

function normalizeYearValue(value = "") {
  const match = String(value ?? "").match(/(19|20)\d{2}/);
  return match ? match[0] : "";
}

function normalizeMonthValue(value = "") {
  const raw = String(value ?? "").trim();

  if (/^(0?[1-9]|1[0-2])$/.test(raw)) {
    return String(Number(raw)).padStart(2, "0");
  }

  const fromHeader = normalizeMonth(raw);
  if (fromHeader) return fromHeader;

  const match = raw.match(/(0?[1-9]|1[0-2])/);
  return match ? String(Number(match[1])).padStart(2, "0") : "";
}

function sameHeader(a = "", b = "") {
  return String(a || "").trim() === String(b || "").trim();
}

function normalizeHeaderLocal(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function findExactColumnHeader(table = {}, headers = []) {
  const targets = headers.map(normalizeHeaderLocal).filter(Boolean);

  const matched = getColumns(table).find((col) => {
    const header = getColumnHeader(col);
    return targets.includes(normalizeHeaderLocal(header));
  });

  return matched ? getColumnHeader(matched) : "";
}

function isExcludedHeader(header = "", excludedHeaders = []) {
  return excludedHeaders
    .filter(Boolean)
    .some((target) => sameHeader(header, target));
}

function isTemporalHeader(header = "") {
  const value = String(header || "").trim();
  return /(연도|년도|년\s*구분|월\s*구분|기준월|연월|년월|기간|일자|날짜|date|year|month|period)/i.test(
    value,
  );
}

function findNonTemporalColumnHeader(
  table = {},
  hints = [],
  excludedHeaders = [],
) {
  const matched = getColumns(table).find((col) => {
    const header = getColumnHeader(col);
    if (!header) return false;
    if (isExcludedHeader(header, excludedHeaders)) return false;
    if (isTemporalHeader(header)) return false;
    return headerMatches(header, hints);
  });

  return matched ? getColumnHeader(matched) : "";
}

function uniqueCandidates(candidates = []) {
  const seen = new Set();

  return candidates.filter((candidate) => {
    const columns = candidate.columns || {};
    const key = [
      candidate.recipeType || "",
      candidate.tableId || "",
      columns.dimension || "",
      columns.date || "",
      columns.metric || "",
      candidate.title || "",
    ].join("|");

    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function findWideYearHeaders(table = {}) {
  return getColumns(table)
    .map((col) => col.header || col.key || "")
    .filter(isYearHeader);
}

function findWideMonthHeaders(table = {}) {
  return getColumns(table)
    .map((col) => col.header || col.key || "")
    .filter(isMonthHeader);
}

function resolveCategoryHeader(table = {}, hints = {}) {
  const categoryHints = [
    ...(hints.item || []),
    ...(hints.category || []),
    "세부사업명",
    "제품명",
    "상품명",
    "품목",
    "업종",
    "구분",
    "분류",
    "카테고리",
    "category",
  ];

  return findColumnHeader(table, categoryHints);
}

function buildYearWideVirtualTable({ table, yearHeaders = [], hints = {} }) {
  const baseCategoryHeader = resolveCategoryHeader(table, hints);
  const groupHeader =
    findColumnHeader(table, [
      ...(hints.group || []),
      "구분",
      "제품류",
      "제품분류",
      "대분류",
    ]) || baseCategoryHeader;

  if (!yearHeaders.length || !baseCategoryHeader) return null;

  const rows = [];

  getRows(table).forEach((row) => {
    yearHeaders.forEach((yearHeader) => {
      const amount = toNumber(getRowValue(row, yearHeader));
      if (amount == null) return;

      rows.push({
        구분: groupHeader ? getRowValue(row, groupHeader) : "",
        항목: getRowValue(row, baseCategoryHeader),
        연도: normalizeYear(yearHeader),
        지표값: amount,
      });
    });
  });

  if (!rows.length) return null;

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_period_year_wide`,
    tableName: `${table.tableName || table.sheetName || "기간지표"}_연도별변환`,
    columns: [
      { header: "구분", type: "category", role: "dimension" },
      { header: "항목", type: "category", role: "dimension" },
      { header: "연도", type: "number", role: "dimension" },
      { header: "지표값", type: "number", role: "metric" },
    ],
    rows,
  });
}

function buildMonthWideVirtualTable({ table, monthHeaders = [], hints = {} }) {
  const yearHeader = findColumnHeader(table, [
    ...(hints.year || []),
    "기준년도",
    "연도",
    "년도",
    "year",
  ]);

  const categoryHeader =
    resolveCategoryHeader(table, hints) ||
    findColumnHeader(table, ["지역", "구", "자치구", "시군구", "region"]);

  if (!monthHeaders.length || !categoryHeader) return null;

  const rows = [];

  getRows(table).forEach((row) => {
    monthHeaders.forEach((monthHeader) => {
      const value = toNumber(getRowValue(row, monthHeader));
      if (value == null) return;

      const year = yearHeader
        ? normalizeYearValue(getRowValue(row, yearHeader))
        : "";
      const month = normalizeMonth(monthHeader);

      rows.push({
        항목: getRowValue(row, categoryHeader),
        연도: year,
        월: month,
        연월: year && month ? `${year}-${month}` : month,
        지표값: value,
      });
    });
  });

  if (!rows.length) return null;

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_period_month_wide`,
    tableName: `${table.tableName || table.sheetName || "기간지표"}_월별변환`,
    columns: [
      { header: "항목", type: "category", role: "dimension" },
      { header: "연도", type: "number", role: "dimension" },
      { header: "월", type: "category", role: "dimension" },
      { header: "연월", type: "category", role: "date" },
      { header: "지표값", type: "number", role: "metric" },
    ],
    rows,
  });
}

function buildLongYearMonthVirtualTable({ table, yearHeader, monthHeader }) {
  if (!yearHeader || !monthHeader) return null;

  const columns = getColumns(table).map((column) => ({ ...column }));
  const hasYearMonth = columns.some(
    (column) => normalizeHeaderLocal(getColumnHeader(column)) === "연월",
  );

  if (!hasYearMonth) {
    columns.push({ header: "연월", type: "category", role: "date" });
  }

  const rows = getRows(table).map((row) => {
    const year = normalizeYearValue(getRowValue(row, yearHeader));
    const month = normalizeMonthValue(getRowValue(row, monthHeader));

    return {
      ...row,
      연월: year && month ? `${year}-${month}` : year || month,
    };
  });

  if (!rows.length) return null;

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_period_year_month`,
    tableName: `${table.tableName || table.sheetName || "기간지표"}_연월변환`,
    columns,
    rows,
  });
}

function selectExecutionTable({
  normalizedQueryTables = [],
  table,
  hints = {},
}) {
  const yearWideVirtualTable = buildYearWideVirtualTable({
    table,
    yearHeaders: findWideYearHeaders(table),
    hints,
  });

  const monthWideVirtualTable = buildMonthWideVirtualTable({
    table,
    monthHeaders: findWideMonthHeaders(table),
    hints,
  });

  const longYearHeader = findColumnHeader(table, [
    ...(hints.year || []),
    "연도 구분",
    "기준년도",
    "진행년도",
    "집행년도",
    "매출연도",
    "연도",
    "년도",
    "year",
  ]);

  const longMonthHeader = findColumnHeader(table, [
    ...(hints.month || []),
    "월 구분",
    "기준월",
    "매출월",
    "월",
    "month",
  ]);

  const longYearMonthVirtualTable = buildLongYearMonthVirtualTable({
    table,
    yearHeader: longYearHeader,
    monthHeader: longMonthHeader,
  });

  const selectedTable =
    monthWideVirtualTable ||
    yearWideVirtualTable ||
    longYearMonthVirtualTable ||
    table;

  if (selectedTable === table) {
    return { selectedTable, executionTables: normalizedQueryTables };
  }

  return {
    selectedTable,
    executionTables: [
      ...normalizedQueryTables.filter(
        (t) => t.tableId !== selectedTable.tableId,
      ),
      selectedTable,
    ],
  };
}

function findMetricHeader(table = {}, hints = []) {
  return (
    findColumnHeader(table, hints, { type: "number" }) ||
    findColumnHeader(table, ["지표값"], { type: "number" })
  );
}

function isPeriodLikeHeader(header = "") {
  return /(연도|년도|월|연월|년월|기간|일자|날짜|date|year|month|period)/i.test(
    String(header || ""),
  );
}

function firstExistingHeader(table = {}, hints = [], excludedHeaders = []) {
  return findNonTemporalColumnHeader(table, hints, excludedHeaders);
}

function resolveRankingDimension({
  table,
  config = {},
  excluded = [],
  periodHeader,
  monthHeader,
  yearHeader,
}) {
  const explicit = firstExistingHeader(
    table,
    config.rankingDimensionHints || [],
    excluded,
  );

  if (explicit) return explicit;

  const businessLabel = firstExistingHeader(
    table,
    [
      "항목명",
      "세부사업명",
      "과제명",
      "사업명",
      "전문기관명",
      "기관명",
      "기관분류",
      "수행기관",
      "연구기관",
      "제품명",
      "상품명",
      "품목",
      "업종",
      "지역",
      "구분",
      "분류",
      "카테고리",
      "category",
      "item",
      "name",
    ],
    excluded,
  );

  if (businessLabel) return businessLabel;

  return periodHeader || monthHeader || yearHeader || "";
}

function periodSortValue(value, header = "") {
  const raw = String(value ?? "").trim();
  if (!raw) return Number.MAX_SAFE_INTEGER;

  const normalized = raw.replace(/\s+/g, "");

  const yearMonthDay = normalized.match(
    /((?:19|20)\d{2})[.\-/년]?(0?[1-9]|1[0-2])?[.\-/월]?(0?[1-9]|[12]\d|3[01])?/,
  );

  if (yearMonthDay) {
    const year = Number(yearMonthDay[1]);
    const month = yearMonthDay[2] ? Number(yearMonthDay[2]) : 0;
    const day = yearMonthDay[3] ? Number(yearMonthDay[3]) : 0;
    return year * 10000 + month * 100 + day;
  }

  if (/월|month/i.test(String(header || ""))) {
    const month = normalizeMonthValue(raw);
    if (month) return Number(month);
  }

  const n = Number(raw.replace(/,/g, ""));
  if (Number.isFinite(n)) return n;

  return raw;
}

function sortRowsByPeriod(rows = [], dimensionHeader = "") {
  if (!Array.isArray(rows) || !dimensionHeader) return rows;

  return [...rows].sort((a, b) => {
    const av = periodSortValue(a?.[dimensionHeader], dimensionHeader);
    const bv = periodSortValue(b?.[dimensionHeader], dimensionHeader);

    if (typeof av === "number" && typeof bv === "number") {
      return av - bv;
    }

    return String(av).localeCompare(String(bv), "ko");
  });
}

function normalizeTemporalSectionRows(section = {}) {
  const sectionType = String(
    section.sectionType ||
      section.candidate?.sectionType ||
      section.candidate?.meta?.sectionType ||
      "",
  );

  if (/top|bottom|상위|하위/i.test(sectionType)) {
    return section;
  }

  const dimensionHeader =
    section.chartHint?.categoryField ||
    section.candidate?.columns?.dimension ||
    section.result?.groupBy?.header ||
    "";

  const shouldSort =
    /year|month|period|trend|연도|년도|월|연월|년월|기간|추이/i.test(
      sectionType,
    ) || isPeriodLikeHeader(dimensionHeader);

  if (!shouldSort || !Array.isArray(section.result?.rows)) {
    return section;
  }

  return {
    ...section,
    result: {
      ...section.result,
      rows: sortRowsByPeriod(section.result.rows, dimensionHeader),
    },
  };
}

function normalizePeriodMetricSections(sections = []) {
  return (sections || []).map(normalizeTemporalSectionRows);
}

function buildPeriodMetricCandidates({ table, config = {} }) {
  const tableId = table.tableId;
  const hints = config.hints || {};
  const metricHeader =
    config.metricHeader || findMetricHeader(table, hints.metric || []);
  const quantityHeader = config.quantityHeader
    ? config.quantityHeader
    : hints.quantity
      ? findColumnHeader(table, hints.quantity, { type: "number" })
      : "";

  const yearHeader = findColumnHeader(table, [
    ...(hints.year || []),
    "연도",
    "년도",
    "진행년도",
    "집행년도",
    "year",
  ]);
  const monthHeader = findColumnHeader(table, [
    ...(hints.month || []),
    "월",
    "기준월",
    "month",
  ]);
  const periodHeader =
    findExactColumnHeader(table, ["연월", "기준년월", "매출년월"]) ||
    findColumnHeader(table, [
      ...(hints.period || []),
      "연월",
      "기간",
      "기준년월",
      "매출년월",
      "period",
    ]);

  const excluded = [
    yearHeader,
    monthHeader,
    periodHeader,
    "연도",
    "월",
    "연월",
  ];
  const candidates = [];

  const pushGroup = ({
    sectionId,
    sectionType,
    title,
    dimension,
    metric,
    chartType = "bar",
  }) => {
    if (!dimension || !metric) return;
    candidates.push(
      makeTemplateCandidate({
        sectionId,
        sectionType,
        recipeType: "group_summary",
        title,
        tableId,
        columns: { dimension, metric },
        chartHint: {
          preferredType: chartType,
          categoryField: dimension,
          valueField: metric,
        },
        narrativeHint: {
          focus: sectionType,
          metric,
          dimension,
        },
      }),
    );
  };

  if (yearHeader && metricHeader) {
    pushGroup({
      sectionId: config.sectionIds?.year || "yearly_metric",
      sectionType: config.sectionTypes?.year || "yearly_metric",
      title: config.titles?.year || `${yearHeader}별 ${metricHeader} 요약`,
      dimension: yearHeader,
      metric: metricHeader,
      chartType: "line",
    });
  }

  if (periodHeader && metricHeader) {
    pushGroup({
      sectionId: config.sectionIds?.period || "period_metric",
      sectionType: config.sectionTypes?.period || "period_metric",
      title: config.titles?.period || `${periodHeader}별 ${metricHeader} 추이`,
      dimension: periodHeader,
      metric: metricHeader,
      chartType: "line",
    });
  }

  if (monthHeader && metricHeader) {
    pushGroup({
      sectionId: config.sectionIds?.month || "monthly_metric",
      sectionType: config.sectionTypes?.month || "monthly_metric",
      title: config.titles?.month || `${monthHeader}별 ${metricHeader} 추이`,
      dimension: monthHeader,
      metric: metricHeader,
      chartType: "line",
    });
  }

  if (monthHeader && quantityHeader) {
    pushGroup({
      sectionId: config.sectionIds?.quantityMonth || "monthly_quantity",
      sectionType: config.sectionTypes?.quantityMonth || "monthly_quantity",
      title:
        config.titles?.quantityMonth ||
        `${monthHeader}별 ${quantityHeader} 추이`,
      dimension: monthHeader,
      metric: quantityHeader,
      chartType: "line",
    });
  }

  for (const item of config.dimensions || []) {
    const dimensionHeader = findNonTemporalColumnHeader(
      table,
      item.hints || [],
      excluded,
    );

    pushGroup({
      sectionId: item.sectionId,
      sectionType: item.sectionType,
      title:
        item.title ||
        (dimensionHeader && metricHeader
          ? `${dimensionHeader}별 ${metricHeader} 요약`
          : ""),
      dimension: dimensionHeader,
      metric: metricHeader,
      chartType: item.chartType || "bar",
    });
  }

  const rankingDimension = resolveRankingDimension({
    table,
    config,
    excluded,
    periodHeader,
    monthHeader,
    yearHeader,
  });

  if (rankingDimension && metricHeader && config.includeTopBottom !== false) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.topBottom || "top_bottom_metric",
        sectionType: config.sectionTypes?.topBottom || "top_bottom_metric",
        recipeType: "top_bottom",
        title: config.titles?.topBottom || `${metricHeader} 상위/하위 항목`,
        tableId,
        columns: {
          dimension: rankingDimension,
          metric: metricHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: rankingDimension,
          valueField: metricHeader,
        },
        narrativeHint: {
          focus: "top_bottom",
          metric: metricHeader,
          dimension: rankingDimension,
        },
      }),
    );
  }

  return uniqueCandidates(candidates);
}

function createAveragePerUnitSection({ table, config = {} }) {
  if (!config.averagePerUnit) return null;

  const hints = config.hints || {};
  const quantityHeader =
    config.quantityHeader ||
    findColumnHeader(table, hints.quantity || [], { type: "number" });
  const metricHeader =
    config.metricHeader || findMetricHeader(table, hints.metric || []);

  if (!quantityHeader || !metricHeader) return null;

  let totalQuantity = 0;
  let totalMetric = 0;

  getRows(table).forEach((row) => {
    const quantity = toNumber(getRowValue(row, quantityHeader));
    const metric = toNumber(getRowValue(row, metricHeader));

    if (quantity != null) totalQuantity += quantity;
    if (metric != null) totalMetric += metric;
  });

  if (!totalQuantity || !totalMetric) return null;

  const title = config.averagePerUnit.title || "평균 단가";

  return makeTemplateSection({
    sectionId: config.averagePerUnit.sectionId || "average_per_unit",
    sectionType: config.averagePerUnit.sectionType || "average_per_unit",
    title,
    candidate: {
      recipeType: "custom_metric",
      title,
      tableId: table.tableId,
      columns: {
        quantity: quantityHeader,
        metric: metricHeader,
      },
    },
    result: {
      ok: true,
      recipeType: "custom_metric",
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns: {
        quantity: quantityHeader,
        metric: metricHeader,
      },
      rows: [
        {
          지표: title,
          수량합계: totalQuantity,
          금액합계: totalMetric,
          값: totalMetric / totalQuantity,
        },
      ],
      rowCount: 1,
    },
    chartHint: { preferredType: "metric_card" },
    narrativeHint: {
      focus: "average_per_unit",
      metric: metricHeader,
      quantity: quantityHeader,
    },
  });
}

function buildPeriodMetricReportSections({
  normalizedQueryTables = [],
  table,
  templateCandidate = {},
  config = {},
}) {
  if (!table?.tableId) return [];

  const { selectedTable, executionTables } = selectExecutionTable({
    normalizedQueryTables,
    table,
    hints: config.hints || {},
  });

  const candidates = buildPeriodMetricCandidates({
    table: selectedTable,
    config,
  });

  const customSections = [
    createAveragePerUnitSection({ table: selectedTable, config }),
  ].filter(Boolean);

  if (!candidates.length && !customSections.length) return [];

  const recipeSections = executeTemplateSections({
    normalizedQueryTables: executionTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });

  return normalizePeriodMetricSections([...recipeSections, ...customSections]);
}

module.exports = {
  buildPeriodMetricReportSections,
  buildPeriodMetricCandidates,
  selectExecutionTable,
  buildYearWideVirtualTable,
  buildMonthWideVirtualTable,
  buildLongYearMonthVirtualTable,
};
