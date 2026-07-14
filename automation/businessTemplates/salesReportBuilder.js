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
const SALES_LEGACY_FILTER_VERSION =
  "sales_legacy_filter_virtual_period_binding_v1";
const SALES_WIDE_TREND_RESTORATION_VERSION = "sales_wide_trend_restoration_v1";
const SALES_RATE_SCALE_VERSION = "growth_rate_scale_hotfix_v1";

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

function getColumnHeader(column = {}) {
  return (
    column.header ||
    column.originalHeader ||
    column.name ||
    column.key ||
    column.accessor ||
    ""
  );
}

function getColumns(table = {}) {
  return Array.isArray(table.columns) ? table.columns : [];
}

function getHeaderStats(table = {}, header = "") {
  const rows = getRows(table);
  let numericCount = 0;
  let nonEmptyCount = 0;
  for (const row of rows) {
    const value = getRowValue(row, header);
    if (value !== null && value !== undefined && String(value).trim() !== "") {
      nonEmptyCount += 1;
      if (toNumber(value) != null) numericCount += 1;
    }
  }
  return {
    rowCount: rows.length,
    numericCount,
    nonEmptyCount,
    numericRatio: nonEmptyCount ? numericCount / nonEmptyCount : 0,
  };
}

function isQuantityLikeHeader(header = "") {
  const n = normalizeText(header);
  return /매출수량|판매수량|수량|판매량|거래건수|주문수|건수|quantity|qty|count/.test(
    n,
  );
}

function isTimeLikeHeader(header = "") {
  const n = normalizeText(header);
  return /연도구분|기준년도|매출연도|판매연도|연도|년도|월구분|기준월|매출월|판매월|월|연월|기간|date|year|month/.test(
    n,
  );
}

function scoreHeaderByHints(header = "", hints = []) {
  const h = normalizeText(header);
  if (!h) return 0;
  let score = 0;
  for (const hint of hints || []) {
    const normalizedHint = normalizeText(hint);
    if (!normalizedHint) continue;
    if (h === normalizedHint) score = Math.max(score, 120);
    else if (h.includes(normalizedHint)) score = Math.max(score, 80);
    else if (normalizedHint.includes(h) && h.length >= 2)
      score = Math.max(score, 35);
  }
  return score;
}

function findBestHeader(table = {}, hints = [], options = {}) {
  const {
    exclude = () => false,
    requireNumeric = false,
    bonus = () => 0,
  } = options;
  let best = { header: "", score: -Infinity };
  for (const column of getColumns(table)) {
    const header = getColumnHeader(column);
    if (!header || exclude(header, column)) continue;
    const stats = getHeaderStats(table, header);
    if (requireNumeric && stats.numericCount === 0) continue;

    let score = scoreHeaderByHints(header, hints);
    if (score <= 0) continue;
    score += Math.min(30, Math.round(stats.numericRatio * 30));
    if (column.role === "metric" || column.inferredRole === "metric")
      score += 15;
    if (column.type === "number" || column.dominantType === "number")
      score += 10;
    score += bonus(header, column, stats) || 0;

    if (score > best.score) best = { header, score };
  }
  return best.header;
}

function findBestDimensionHeader(table = {}, hints = [], options = {}) {
  const { exclude = () => false, bonus = () => 0 } = options;
  let best = { header: "", score: -Infinity };
  for (const column of getColumns(table)) {
    const header = getColumnHeader(column);
    if (!header || exclude(header, column)) continue;
    const h = normalizeText(header);
    let score = 0;
    for (const hint of hints || []) {
      const normalizedHint = normalizeText(hint);
      if (!normalizedHint) continue;
      if (normalizedHint.length <= 1 && h !== normalizedHint) continue;
      if (h === normalizedHint) score = Math.max(score, 100);
      else if (h.includes(normalizedHint)) score = Math.max(score, 55);
      else if (normalizedHint.includes(h) && h.length >= 2)
        score = Math.max(score, 25);
    }
    if (score <= 0) continue;
    if (isTimeLikeHeader(header)) score -= 35;
    if (isQuantityLikeHeader(header)) score -= 60;
    score += bonus(header, column) || 0;
    if (score > best.score) best = { header, score };
  }
  return best.header;
}

function findBestTimeHeader(table = {}, hints = []) {
  let best = { header: "", score: -Infinity };
  for (const column of getColumns(table)) {
    const header = getColumnHeader(column);
    const h = normalizeText(header);
    if (!h) continue;
    let score = 0;
    for (const hint of hints || []) {
      const normalizedHint = normalizeText(hint);
      if (!normalizedHint) continue;
      if (h === normalizedHint) score = Math.max(score, 120);
      else if (h.length >= 2 && h.includes(normalizedHint))
        score = Math.max(score, 75);
      else if (h.length >= 2 && normalizedHint.includes(h))
        score = Math.max(score, 20);
    }
    if (score <= 0) continue;
    if (/연도|년도|year/.test(h)) score += 30;
    if (/월|month/.test(h)) score += 30;
    if (/기간|연월|date/.test(h)) score += 25;
    if (h.length <= 1) score -= 100;
    if (score > best.score) best = { header, score };
  }
  return best.header;
}

function findSalesAmountHeader(table = {}, hints = SALES_HINTS) {
  const amountHints = [
    "순매출액",
    "매출액",
    "매출금액",
    "판매금액",
    "판매액",
    "카드매출액",
    "신용카드매출액",
    "거래액",
    "출하액",
    "금액",
    "지표값",
    "revenue",
    "salesamount",
    "amount",
    "transactionamount",
  ];

  return findBestHeader(table, amountHints, {
    requireNumeric: true,
    exclude: (header) =>
      isQuantityLikeHeader(header) || isTimeLikeHeader(header),
    bonus: (header) => {
      const h = normalizeText(header);
      if (
        /순매출액|매출액|매출금액|판매금액|판매액|카드매출액|신용카드매출액|거래액|출하액/.test(
          h,
        )
      )
        return 80;
      if (h === "지표값") return 40;
      if (h === "금액") return 10;
      return 0;
    },
  });
}

function findSalesQuantityHeader(
  table = {},
  hints = SALES_HINTS,
  metricHeader = "",
) {
  return findBestHeader(table, hints.quantity || [], {
    requireNumeric: true,
    exclude: (header) => sameHeader(header, metricHeader),
    bonus: (header) => (isQuantityLikeHeader(header) ? 60 : 0),
  });
}

function scoreSalesTable(table = {}, config = {}) {
  const hints = config.hints || SALES_HINTS;
  const metricHeader = findSalesAmountHeader(table, hints);
  const quantityHeader = findSalesQuantityHeader(table, hints, metricHeader);
  const periodHeader =
    findBestTimeHeader(table, [
      "연월",
      "기간",
      "기준년월",
      "월",
      "연도",
      "년도",
      "year",
      "month",
    ]) || "";
  const dimensionHeader =
    findBestDimensionHeader(table, [
      ...(hints.item || []),
      ...(hints.category || []),
      ...(hints.region || []),
      ...(hints.channel || []),
      ...(hints.customer || []),
    ]) || "";
  const tableId = String(table.tableId || "");

  let score = 0;
  if (metricHeader) score += 70;
  if (quantityHeader) score += 15;
  if (periodHeader) score += 20;
  if (dimensionHeader) score += 10;
  if (table.isVirtual || /WIDE_LONG|CROSS_LONG/.test(tableId))
    score += metricHeader ? 25 : 0;
  score += Math.min(10, Math.floor((getRows(table).length || 0) / 20));
  if (table.isPrimary) score += 5;
  return score;
}

function selectSalesExecutionTable({
  normalizedQueryTables = [],
  table = {},
  config = {},
}) {
  const candidates = [table, ...normalizedQueryTables].filter(Boolean);
  let best = { table, score: scoreSalesTable(table, config) };
  for (const candidate of candidates) {
    const score = scoreSalesTable(candidate, config);
    if (score > best.score) best = { table: candidate, score };
  }
  return best.table || table;
}

function getAllCandidateHeaders(table = {}) {
  const headers = [];
  const seen = new Set();
  const add = (header) => {
    const value = String(header || "").trim();
    if (!value || seen.has(value)) return;
    seen.add(value);
    headers.push(value);
  };

  for (const column of getColumns(table)) add(getColumnHeader(column));
  for (const row of getRows(table).slice(0, 50)) {
    if (row && typeof row === "object" && !Array.isArray(row)) {
      Object.keys(row).forEach(add);
    }
  }
  return headers;
}

function findBestTimeHeaderAny(table = {}, hints = []) {
  const headers = getAllCandidateHeaders(table);
  let best = { header: "", score: -Infinity };
  for (const header of headers) {
    const h = normalizeText(header);
    if (!h) continue;
    let score = 0;
    for (const hint of hints || []) {
      const normalizedHint = normalizeText(hint);
      if (!normalizedHint) continue;
      if (h === normalizedHint) score = Math.max(score, 130);
      else if (h.length >= 2 && h.includes(normalizedHint))
        score = Math.max(score, 85);
      else if (h.length >= 2 && normalizedHint.includes(h))
        score = Math.max(score, 25);
    }
    if (/^기간$|^연월$|기준년월|매출년월|판매년월|period|date/.test(h))
      score += 45;
    if (/^연도$|^년도$|기준년도|매출연도|판매연도|year/.test(h)) score += 35;
    if (/^월$|기준월|매출월|판매월|month/.test(h)) score += 30;
    if (h.length <= 1 && h !== "월") score -= 100;
    if (score <= 0) continue;
    if (score > best.score) best = { header, score };
  }
  return best.header;
}

function resolveSalesColumnsForTable(table = {}, config = {}) {
  const hints = config.hints || SALES_HINTS;
  const metricHeader =
    config.metricHeader ||
    findSalesAmountHeader(table, hints) ||
    findFirstColumnHeader(table, [["지표값", "금액", "value", "amount"]]);
  const quantityHeader =
    config.quantityHeader ||
    findSalesQuantityHeader(table, hints, metricHeader);
  const yearHeader =
    findBestTimeHeader(table, hints.year || []) ||
    findBestTimeHeaderAny(table, hints.year || []);
  const monthHeader =
    findBestTimeHeader(table, hints.month || []) ||
    findBestTimeHeaderAny(table, hints.month || []);
  const periodHeader =
    findBestTimeHeader(table, [
      "연월",
      "기준년월",
      "매출년월",
      "판매년월",
      "기간",
      "period",
      "date",
    ]) ||
    findBestTimeHeaderAny(table, [
      "연월",
      "기준년월",
      "매출년월",
      "판매년월",
      "기간",
      "period",
      "date",
    ]) ||
    (yearHeader && monthHeader ? "__YEAR_MONTH__" : "") ||
    findColumnHeader(table, hints.period || []) ||
    yearHeader ||
    monthHeader;
  const itemHeader = findBestDimensionHeader(table, hints.item || []);
  const categoryHeader = findBestDimensionHeader(table, hints.category || []);
  const regionHeader = findBestDimensionHeader(table, hints.region || [], {
    exclude: (header) =>
      /구분|분류|연도|년도|월/.test(normalizeText(header)) &&
      normalizeText(header) !== "구",
  });
  const channelHeader = findBestDimensionHeader(table, hints.channel || []);
  const customerHeader = findBestDimensionHeader(table, hints.customer || []);

  return {
    metricHeader,
    quantityHeader: sameHeader(quantityHeader, metricHeader)
      ? ""
      : quantityHeader,
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

function countValidSalesTrendRows(table = {}, columns = {}) {
  const metricHeader = columns.metricHeader;
  if (!metricHeader || !columns.periodHeader) return 0;
  let count = 0;
  for (const row of getRows(table)) {
    const period = getSalesPeriod(row, columns);
    const amount = toNumber(getRowValue(row, metricHeader));
    if (period && amount != null) count += 1;
  }
  return count;
}

function scoreSalesTrendTable(table = {}, config = {}) {
  if (!table?.tableId) return -Infinity;
  const columns = resolveSalesColumnsForTable(table, config);
  if (!columns.metricHeader || !columns.periodHeader) return -Infinity;
  const validRows = countValidSalesTrendRows(table, columns);
  if (!validRows) return -Infinity;

  let score = validRows;
  const tableId = String(table.tableId || "");
  if (isVirtualSalesTable(table)) score += 500;
  if (/WIDE_LONG/i.test(tableId)) score += 120;
  if (/CROSS_LONG/i.test(tableId)) score += 80;
  if (columns.periodHeader === "기간") score += 80;
  if (columns.periodHeader === "__YEAR_MONTH__") score += 70;
  if (columns.yearHeader && columns.monthHeader) score += 35;
  if (columns.metricHeader === "지표값") score += 20;
  if (table.isPrimary) score += 5;
  return score;
}

function selectSalesTrendTable({
  normalizedQueryTables = [],
  table = {},
  config = {},
}) {
  const candidates = [table, ...normalizedQueryTables].filter(Boolean);
  let best = { table: null, score: -Infinity };
  for (const candidate of candidates) {
    const score = scoreSalesTrendTable(candidate, config);
    if (score > best.score) best = { table: candidate, score };
  }
  return best.table || table;
}

function findFirstColumnHeader(table = {}, hintGroups = []) {
  for (const hints of hintGroups) {
    const matched = findColumnHeader(table, hints || []);
    if (matched) return matched;
  }
  return "";
}

function resolveSalesColumns(table = {}, config = {}) {
  return resolveSalesColumnsForTable(table, config);
}

function normalizePeriodValue(value = "") {
  const raw = String(value ?? "").trim();
  if (!raw) return "";

  const compact = raw.replace(/\s+/g, "");
  const ym = compact.match(/((?:19|20)\d{2})[.\-/년]?(1[0-2]|0?[1-9])?/);
  if (ym) {
    const year = ym[1];
    const month = ym[2] ? String(Number(ym[2])).padStart(2, "0") : "";
    return month ? `${year}-${month}` : year;
  }

  const month = compact.match(/^(1[0-2]|0?[1-9])월?$/);
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

function getSalesPeriod(row = {}, columns = {}) {
  const { periodHeader, yearHeader, monthHeader } = columns;
  if (periodHeader === "__YEAR_MONTH__" && yearHeader && monthHeader) {
    const year =
      String(getRowValue(row, yearHeader) ?? "").match(/(19|20)\d{2}/)?.[0] ||
      "";
    const monthRaw = String(getRowValue(row, monthHeader) ?? "").trim();
    const month = monthRaw.match(/^(1[0-2]|0?[1-9])(?:월)?$/)?.[1] || "";
    if (year && month)
      return `${year}-${String(Number(month)).padStart(2, "0")}`;
    if (year) return year;
  }
  return normalizePeriodValue(getRowValue(row, periodHeader));
}

function getSalesPeriodLabel(columns = {}) {
  return columns.periodHeader === "__YEAR_MONTH__"
    ? "연월"
    : columns.periodHeader;
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
        rateValueScale: "ratio",
        growthRateScaleVersion: SALES_RATE_SCALE_VERSION,
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
        rateValueScale: "ratio",
        growthRateScaleVersion: SALES_RATE_SCALE_VERSION,
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
    const period = getSalesPeriod(row, columns);
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
    title: `${getSalesPeriodLabel(columns)}별 ${metricHeader} 추이 및 증감률`,
    table,
    rows,
    columns: {
      date: getSalesPeriodLabel(columns),
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
    const period = periodHeader ? getSalesPeriod(row, columns) : "전체";
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
      date: getSalesPeriodLabel(columns) || null,
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

function getSalesRankingDimension(columns = {}) {
  return (
    columns.itemHeader ||
    columns.categoryHeader ||
    columns.regionHeader ||
    columns.channelHeader ||
    columns.customerHeader ||
    (columns.periodHeader ? "__PERIOD__" : "")
  );
}

function buildSalesTopBottomSection({ table, columns }) {
  const { metricHeader } = columns;
  if (!metricHeader) return null;

  const rankingDimension = getSalesRankingDimension(columns);
  if (!rankingDimension) return null;

  let aggregated = [];
  let dimensionHeader = rankingDimension;

  if (rankingDimension === "__PERIOD__") {
    dimensionHeader = getSalesPeriodLabel(columns) || "기간";
    const periodMap = new Map();
    for (const row of getRows(table)) {
      const period = getSalesPeriod(row, columns);
      const amount = toNumber(getRowValue(row, metricHeader));
      if (!period || amount == null) continue;
      const current = periodMap.get(period) || {
        label: period,
        value: 0,
        count: 0,
      };
      current.value += amount;
      current.count += 1;
      periodMap.set(period, current);
    }
    aggregated = Array.from(periodMap.values());
  } else {
    aggregated = aggregateBy({ table, dimensionHeader, metricHeader });
  }

  aggregated = aggregated
    .filter((row) => row && row.value != null && row.value !== 0)
    .sort((a, b) => b.value - a.value);
  if (!aggregated.length) return null;

  const top = aggregated.slice(0, 10).map((row, index) => ({
    구분: "상위",
    순위: index + 1,
    [dimensionHeader]: row.label,
    매출합계: row.value,
    건수: row.count,
  }));
  const bottom = aggregated
    .slice(-10)
    .reverse()
    .map((row, index) => ({
      구분: "하위",
      순위: index + 1,
      [dimensionHeader]: row.label,
      매출합계: row.value,
      건수: row.count,
    }));

  return makeSalesSection({
    sectionId: "top_bottom_sales_v2",
    sectionType: "top_bottom_sales",
    title: `${dimensionHeader} 기준 ${metricHeader} 상위/하위`,
    table,
    rows: [...top, ...bottom],
    columns: {
      dimension: dimensionHeader,
      metric: metricHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "매출합계",
    },
    narrativeHint: {
      focus: "top_bottom_sales",
      metric: metricHeader,
      dimension: dimensionHeader,
    },
  });
}

function getSectionRows(section = {}) {
  return section?.result?.rows || section?.rows || [];
}

function hasMeaningfulNumericValue(row = {}) {
  return Object.entries(row || {}).some(([key, value]) => {
    if (["구분", "순위", "기간", "건수", "행수"].includes(key)) return false;
    const number = toNumber(value);
    return number != null && number !== 0;
  });
}

function isEmptyOrNoDataSection(section = {}) {
  const rows = getSectionRows(section);
  if (!Array.isArray(rows) || rows.length === 0) return true;
  const title = String(section.title || section?.result?.title || "");
  if (/데이터\s*없음|no\s*data/i.test(title)) return true;
  return !rows.some((row) => hasMeaningfulNumericValue(row));
}

function isVirtualSalesTable(table = {}) {
  const id = String(table.tableId || table.sourceTableId || "");
  return Boolean(
    table.isVirtual ||
    /#?(WIDE_LONG|CROSS_LONG)$/i.test(id) ||
    /WIDE_LONG|CROSS_LONG/i.test(id),
  );
}

function shouldKeepSalesLegacySection(section = {}, context = {}) {
  if (!section) return false;
  const { v2Sections = [], selectedTable = {} } = context;
  const sectionId = String(
    section.sectionId || section?.candidate?.sectionId || "",
  );
  const sectionType = String(
    section.sectionType || section?.candidate?.sectionType || "",
  );
  const title = String(section.title || section?.result?.title || "");
  const hasV2Average = v2Sections.some(
    (item) => item.sectionId === "average_sales_amount_v2",
  );
  const hasV2TopBottom = v2Sections.some(
    (item) => item.sectionId === "top_bottom_sales_v2",
  );
  const hasV2Trend = v2Sections.some(
    (item) => item.sectionId === "sales_trend_growth_v2",
  );
  const isVirtual = isVirtualSalesTable(selectedTable);

  if (isEmptyOrNoDataSection(section)) return false;
  if (
    hasV2Average &&
    /average_sales_amount|평균\s*판매금액/.test(
      `${sectionId} ${sectionType} ${title}`,
    )
  ) {
    return false;
  }
  if (
    hasV2TopBottom &&
    /top_bottom_sales|매출\s*상위|상위\s*하위/.test(
      `${sectionId} ${sectionType} ${title}`,
    )
  ) {
    return false;
  }
  if (
    isVirtual &&
    hasV2Trend &&
    /yearly_sales|period_sales|monthly_sales|monthly_quantity/.test(sectionId)
  ) {
    return false;
  }
  if (
    isVirtual &&
    hasV2Trend &&
    /(연도별|월별|기간별).*매출\s*추이|월별\s*판매수량/.test(title)
  ) {
    return false;
  }
  return true;
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
  const { selectedTable: structuralSelectedTable } = selectExecutionTable({
    normalizedQueryTables,
    table,
    hints: config.hints || SALES_HINTS,
  });
  const selectedTable = selectSalesExecutionTable({
    normalizedQueryTables,
    table: structuralSelectedTable || table,
    config,
  });
  const columns = resolveSalesColumns(selectedTable, config);

  const trendTable = selectSalesTrendTable({
    normalizedQueryTables,
    table: selectedTable,
    config,
  });
  const trendColumns = resolveSalesColumns(trendTable, config);

  const trendSection = buildSalesTrendSection({
    table: trendTable,
    columns: trendColumns,
  });
  const averageSection = buildSalesAverageTicketSection({
    table: trendSection ? trendTable : selectedTable,
    columns: trendSection ? trendColumns : columns,
  });

  const sections = [
    trendSection,
    buildSalesCompositionSection({ table: selectedTable, columns }),
    ...buildSalesRegionChannelSections({ table: selectedTable, columns }),
    buildSalesTopBottomSection({ table: selectedTable, columns }),
    averageSection,
  ].filter(Boolean);

  return uniqueSections(sections).map((section) => ({
    ...section,
    candidate: {
      ...section.candidate,
      templateId: templateCandidate.templateId || "sales_report",
      meta: {
        ...(section.candidate?.meta || {}),
        salesWideTrendRestorationVersion: SALES_WIDE_TREND_RESTORATION_VERSION,
      },
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

  const selectedBusinessTable = selectSalesExecutionTable({
    normalizedQueryTables,
    table,
    config,
  });

  const baseSections = buildPeriodMetricReportSections({
    normalizedQueryTables,
    table: selectedBusinessTable,
    templateCandidate,
    config,
  });

  const v2Sections = buildSalesV2Sections({
    normalizedQueryTables,
    table: selectedBusinessTable,
    templateCandidate,
    config,
  });

  const filteredBaseSections = baseSections.filter(
    (section) =>
      !hasSameSection(v2Sections, section.sectionId) &&
      shouldKeepSalesLegacySection(section, {
        v2Sections,
        selectedTable: selectedBusinessTable,
      }),
  );

  const sections = uniqueSections([...v2Sections, ...filteredBaseSections]).map(
    (section) => ({
      ...section,
      meta: {
        ...(section.meta || {}),
        salesReportVersion: SALES_REPORT_VERSION,
        salesLegacyFilterVersion: SALES_LEGACY_FILTER_VERSION,
        salesWideTrendRestorationVersion: SALES_WIDE_TREND_RESTORATION_VERSION,
      },
    }),
  );

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
  resolveSalesColumns,
  findSalesAmountHeader,
  selectSalesExecutionTable,
  selectSalesTrendTable,
  buildSalesTopBottomSection,
  shouldKeepSalesLegacySection,
  normalizePeriodValue,
  getSalesPeriod,
};
