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

const SALES_REPORT_VERSION = "sales_report_v2";

const SALES_HINTS = {
  metric: [
    "순매출액",
    "매출액",
    "매출금액",
    "판매금액",
    "판매액",
    "카드매출액",
    "신용카드매출액",
    "거래액",
    "매출",
    "금액",
    "출하액",
    "구성비율",
    "매출구성비율",
    "매출증가율",
    "증가율",
    "sales",
    "revenue",
    "amount",
    "transaction",
  ],
  quantity: [
    "매출수량",
    "판매수량",
    "수량",
    "판매량",
    "건수",
    "주문수",
    "거래건수",
    "quantity",
    "qty",
    "count",
  ],
  year: [
    "연도 구분",
    "기준년도",
    "매출연도",
    "판매연도",
    "연도",
    "년도",
    "year",
  ],
  month: ["월 구분", "기준월", "매출월", "판매월", "월", "month"],
  period: [
    "연월",
    "기간",
    "기준년월",
    "매출년월",
    "판매년월",
    "일자",
    "날짜",
    "period",
    "date",
  ],
  item: [
    "제품명",
    "상품명",
    "품목",
    "품명",
    "제품류",
    "제품군",
    "브랜드",
    "세부사업명",
    "product",
    "item",
    "sku",
  ],
  category: [
    "제품분류",
    "상품분류",
    "업종",
    "업태",
    "업태별",
    "카테고리",
    "대분류",
    "중분류",
    "소분류",
    "구분",
    "분류",
    "category",
    "type",
  ],
  region: [
    "지역",
    "권역",
    "시도",
    "시군구",
    "자치구",
    "구",
    "점포지역",
    "region",
    "area",
  ],
  channel: [
    "채널",
    "판매채널",
    "유통채널",
    "매장",
    "점포",
    "온라인",
    "오프라인",
    "channel",
    "store",
  ],
  customer: [
    "고객",
    "거래처",
    "회원",
    "고객군",
    "customer",
    "client",
    "account",
  ],
  group: ["구분", "제품류", "제품분류", "대분류", "업종", "지역"],
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

function resolveSalesColumns(table = {}, config = {}) {
  const hints = config.hints || SALES_HINTS;
  const metricHeader =
    config.metricHeader ||
    findFirstColumnHeader(table, [
      hints.metric,
      ["지표값", "금액", "value", "amount"],
    ]);
  const quantityHeader =
    config.quantityHeader || findColumnHeader(table, hints.quantity || []);
  const yearHeader = findColumnHeader(table, hints.year || []);
  const monthHeader = findColumnHeader(table, hints.month || []);
  const periodHeader =
    findColumnHeader(table, ["연월", "기준년월", "매출년월", "판매년월"]) ||
    findColumnHeader(table, hints.period || []) ||
    yearHeader ||
    monthHeader;
  const itemHeader = findColumnHeader(table, hints.item || []);
  const categoryHeader = findColumnHeader(table, hints.category || []);
  const regionHeader = findColumnHeader(table, hints.region || []);
  const channelHeader = findColumnHeader(table, hints.channel || []);
  const customerHeader = findColumnHeader(table, hints.customer || []);

  return {
    metricHeader,
    quantityHeader,
    yearHeader,
    monthHeader,
    periodHeader,
    itemHeader,
    categoryHeader,
    regionHeader,
    channelHeader,
    customerHeader,
  };
}

function normalizePeriodValue(value = "") {
  const raw = String(value ?? "").trim();
  if (!raw) return "";

  const compact = raw.replace(/\s+/g, "");
  const ym = compact.match(/((?:19|20)\d{2})[.\-/년]?(0?[1-9]|1[0-2])?/);
  if (ym) {
    const year = ym[1];
    const month = ym[2] ? String(Number(ym[2])).padStart(2, "0") : "";
    return month ? `${year}-${month}` : year;
  }

  const month = compact.match(/^(0?[1-9]|1[0-2])월?$/);
  if (month) return `${String(Number(month[1])).padStart(2, "0")}월`;

  return raw;
}

function periodSortValue(value = "") {
  const normalized = normalizePeriodValue(value);
  const ym = normalized.match(/^((?:19|20)\d{2})(?:-(0[1-9]|1[0-2]))?$/);
  if (ym) return Number(ym[1]) * 100 + Number(ym[2] || 0);
  const month = normalized.match(/^(0[1-9]|1[0-2])월$/);
  if (month) return Number(month[1]);
  const n = Number(String(normalized).replace(/,/g, ""));
  return Number.isFinite(n) ? n : normalized;
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

function aggregateBy({ table, dimensionHeader, metricHeader }) {
  if (!dimensionHeader || !metricHeader) return [];

  const map = new Map();
  for (const row of getRows(table)) {
    const key = getDimensionValue(row, dimensionHeader);
    const value = toNumber(getRowValue(row, metricHeader));
    if (value == null) continue;
    const current = map.get(key) || { label: key, value: 0, count: 0 };
    current.value += value;
    current.count += 1;
    map.set(key, current);
  }

  return Array.from(map.values());
}

function makeSalesSection({
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
      recipeType: "sales_report_v2_custom",
      sectionType,
      title,
      tableId: table.tableId,
      columns,
      meta: {
        salesReportVersion: SALES_REPORT_VERSION,
        sectionType,
      },
    },
    result: {
      ok: true,
      recipeType: "sales_report_v2_custom",
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
      meta: {
        salesReportVersion: SALES_REPORT_VERSION,
      },
    },
    chartHint,
    narrativeHint: {
      ...narrativeHint,
      salesReportVersion: SALES_REPORT_VERSION,
    },
  });
}

function buildSalesTrendSection({ table, columns }) {
  const { periodHeader, metricHeader, quantityHeader } = columns;
  if (!periodHeader || !metricHeader) return null;

  const periodMap = new Map();
  for (const row of getRows(table)) {
    const period = normalizePeriodValue(getRowValue(row, periodHeader));
    const amount = toNumber(getRowValue(row, metricHeader));
    const quantity = quantityHeader
      ? toNumber(getRowValue(row, quantityHeader))
      : null;
    if (!period || amount == null) continue;

    const current = periodMap.get(period) || {
      기간: period,
      매출합계: 0,
      수량합계: 0,
      건수: 0,
    };
    current.매출합계 += amount;
    if (quantity != null) current.수량합계 += quantity;
    current.건수 += 1;
    periodMap.set(period, current);
  }

  const rows = Array.from(periodMap.values()).sort((a, b) =>
    comparePeriodValues(a.기간, b.기간),
  );

  let previous = null;
  let cumulative = 0;
  rows.forEach((row) => {
    cumulative += row.매출합계;
    row.누적매출 = cumulative;
    row.전기매출 = previous == null ? null : previous;
    row.전기대비증감 = previous == null ? null : row.매출합계 - previous;
    row.전기대비증감률 =
      previous == null || previous === 0
        ? null
        : (row.매출합계 - previous) / previous;
    if (quantityHeader && row.수량합계) {
      row.평균판매금액 = row.매출합계 / row.수량합계;
    }
    previous = row.매출합계;
  });

  return makeSalesSection({
    sectionId: "sales_trend_growth_v2",
    sectionType: "sales_trend_growth",
    title: `${periodHeader}별 ${metricHeader} 추이 및 증감률`,
    table,
    rows,
    columns: {
      date: periodHeader,
      metric: metricHeader,
      quantity: quantityHeader || null,
    },
    chartHint: {
      preferredType: "line",
      categoryField: "기간",
      valueField: "매출합계",
      secondaryValueFields: ["누적매출", "전기대비증감률"],
    },
    narrativeHint: {
      focus: "sales_trend_growth",
      metric: metricHeader,
      date: periodHeader,
    },
  });
}

function buildSalesCompositionSection({ table, columns }) {
  const {
    metricHeader,
    itemHeader,
    categoryHeader,
    regionHeader,
    channelHeader,
  } = columns;
  if (!metricHeader) return null;

  const dimensionHeader =
    itemHeader || categoryHeader || regionHeader || channelHeader || "";
  if (!dimensionHeader) return null;

  const aggregated = aggregateBy({ table, dimensionHeader, metricHeader })
    .sort((a, b) => b.value - a.value)
    .slice(0, 20);

  const total = aggregated.reduce((sum, row) => sum + row.value, 0);
  if (!total) return null;

  const rows = aggregated.map((row, index) => ({
    순위: index + 1,
    [dimensionHeader]: row.label,
    매출합계: row.value,
    구성비: row.value / total,
    건수: row.count,
  }));

  return makeSalesSection({
    sectionId: "sales_composition_v2",
    sectionType: "sales_composition",
    title: `${dimensionHeader}별 ${metricHeader} 구성비`,
    table,
    rows,
    columns: {
      dimension: dimensionHeader,
      metric: metricHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "매출합계",
      secondaryValueFields: ["구성비"],
    },
    narrativeHint: {
      focus: "sales_composition",
      metric: metricHeader,
      dimension: dimensionHeader,
    },
  });
}

function buildSalesRegionChannelSections({ table, columns }) {
  const { metricHeader, regionHeader, channelHeader, customerHeader } = columns;
  if (!metricHeader) return [];

  return [
    {
      header: regionHeader,
      id: "region_sales_v2",
      type: "region_sales",
      label: "지역",
    },
    {
      header: channelHeader,
      id: "channel_sales_v2",
      type: "channel_sales",
      label: "채널",
    },
    {
      header: customerHeader,
      id: "customer_sales_v2",
      type: "customer_sales",
      label: "고객/거래처",
    },
  ]
    .map(({ header, id, type, label }) => {
      if (!header) return null;
      const rows = aggregateBy({ table, dimensionHeader: header, metricHeader })
        .sort((a, b) => b.value - a.value)
        .slice(0, 20)
        .map((row, index) => ({
          순위: index + 1,
          [header]: row.label,
          매출합계: row.value,
          건수: row.count,
        }));

      return makeSalesSection({
        sectionId: id,
        sectionType: type,
        title: `${label}별 ${metricHeader} 요약`,
        table,
        rows,
        columns: {
          dimension: header,
          metric: metricHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: header,
          valueField: "매출합계",
        },
        narrativeHint: {
          focus: type,
          metric: metricHeader,
          dimension: header,
        },
      });
    })
    .filter(Boolean);
}

function buildSalesAverageTicketSection({ table, columns }) {
  const { metricHeader, quantityHeader, periodHeader } = columns;
  if (!metricHeader || !quantityHeader) return null;

  const periodMap = new Map();
  for (const row of getRows(table)) {
    const period = periodHeader
      ? normalizePeriodValue(getRowValue(row, periodHeader))
      : "전체";
    const amount = toNumber(getRowValue(row, metricHeader));
    const quantity = toNumber(getRowValue(row, quantityHeader));
    if (!period || amount == null || quantity == null || !quantity) continue;

    const current = periodMap.get(period) || {
      기간: period,
      매출합계: 0,
      수량합계: 0,
    };
    current.매출합계 += amount;
    current.수량합계 += quantity;
    periodMap.set(period, current);
  }

  const rows = Array.from(periodMap.values())
    .sort((a, b) => comparePeriodValues(a.기간, b.기간))
    .map((row) => ({
      ...row,
      평균판매금액: row.수량합계 ? row.매출합계 / row.수량합계 : null,
    }));

  return makeSalesSection({
    sectionId: "average_sales_amount_v2",
    sectionType: "average_sales_amount",
    title: "평균 판매금액 추이",
    table,
    rows,
    columns: {
      date: periodHeader || null,
      metric: metricHeader,
      quantity: quantityHeader,
    },
    chartHint: {
      preferredType: "line",
      categoryField: "기간",
      valueField: "평균판매금액",
    },
    narrativeHint: {
      focus: "average_sales_amount",
      metric: metricHeader,
      quantity: quantityHeader,
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

function buildSalesV2Sections({
  normalizedQueryTables,
  table,
  templateCandidate,
  config,
}) {
  const { selectedTable } = selectExecutionTable({
    normalizedQueryTables,
    table,
    hints: config.hints || SALES_HINTS,
  });
  const columns = resolveSalesColumns(selectedTable, config);

  const sections = [
    buildSalesTrendSection({ table: selectedTable, columns }),
    buildSalesCompositionSection({ table: selectedTable, columns }),
    ...buildSalesRegionChannelSections({ table: selectedTable, columns }),
    buildSalesAverageTicketSection({ table: selectedTable, columns }),
  ].filter(Boolean);

  return uniqueSections(sections).map((section) => ({
    ...section,
    candidate: {
      ...section.candidate,
      templateId: templateCandidate.templateId || "sales_report",
    },
  }));
}

function buildSalesReportConfig() {
  return {
    hints: SALES_HINTS,
    sectionIds: {
      year: "yearly_sales",
      period: "period_sales",
      month: "monthly_sales",
      quantityMonth: "monthly_quantity",
      topBottom: "top_bottom_sales",
    },
    sectionTypes: {
      year: "yearly_sales",
      period: "period_sales",
      month: "monthly_sales",
      quantityMonth: "monthly_quantity",
      topBottom: "top_bottom_sales",
    },
    titles: {
      year: "연도별 매출 추이",
      period: "기간별 매출 추이",
      month: "월별 매출 추이",
      quantityMonth: "월별 판매수량 추이",
      topBottom: "매출 상위/하위 항목",
    },
    dimensions: [
      {
        sectionId: "product_sales",
        sectionType: "product_sales",
        title: "제품/상품별 매출 요약",
        hints: SALES_HINTS.item,
      },
      {
        sectionId: "category_sales",
        sectionType: "category_sales",
        title: "카테고리별 매출 요약",
        hints: SALES_HINTS.category,
      },
      {
        sectionId: "region_sales",
        sectionType: "region_sales",
        title: "지역별 매출 요약",
        hints: SALES_HINTS.region,
      },
      {
        sectionId: "channel_sales",
        sectionType: "channel_sales",
        title: "채널/점포별 매출 요약",
        hints: SALES_HINTS.channel,
      },
      {
        sectionId: "customer_sales",
        sectionType: "customer_sales",
        title: "고객/거래처별 매출 요약",
        hints: SALES_HINTS.customer,
      },
    ],
    rankingDimensionHints: [
      "연월",
      "제품명",
      "상품명",
      "품목",
      "제품류",
      "제품분류",
      "카테고리",
      "업종",
      "지역",
      "자치구",
      "채널",
      "매장",
      "점포",
      "거래처",
    ],
    averagePerUnit: {
      sectionId: "average_sales_amount",
      sectionType: "average_sales_amount",
      title: "평균 판매금액",
    },
  };
}

function executeSalesReport({
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

  const config = buildSalesReportConfig();

  const baseSections = buildPeriodMetricReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config,
  });

  const v2Sections = buildSalesV2Sections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config,
  });

  const sections = uniqueSections([
    ...v2Sections,
    ...baseSections.filter(
      (section) => !hasSameSection(v2Sections, section.sectionId),
    ),
  ]).map((section) => ({
    ...section,
    meta: {
      ...(section.meta || {}),
      salesReportVersion: SALES_REPORT_VERSION,
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
  SALES_REPORT_VERSION,
  executeSalesReport,
  buildSalesReportConfig,
  buildSalesV2Sections,
};
