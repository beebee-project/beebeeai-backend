const {
  findTableForTemplate,
  findColumnHeader,
  findNumericHeaders,
  makeTemplateCandidate,
  executeTemplateSections,
  getColumns,
  getRows,
  getRowValue,
  createVirtualTable,
  toNumber,
  headerMatches,
  getColumnHeader,
} = require("./commonTemplateHelpers");

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

function isExcludedHeader(header = "", excludedHeaders = []) {
  return excludedHeaders
    .filter(Boolean)
    .some((target) => sameHeader(header, target));
}

function isTemporalSalesHeader(header = "") {
  const value = String(header || "").trim();

  return /(연도|년도|년\s*구분|월\s*구분|기준월|매출월|연월|년월|기간|일자|날짜|date|year|month|period)/i.test(
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
    if (isTemporalSalesHeader(header)) return false;

    return headerMatches(header, hints);
  });

  return matched ? getColumnHeader(matched) : "";
}

function findSalesQuantityHeader(table = {}) {
  return findColumnHeader(
    table,
    ["매출수량", "판매수량", "수량", "quantity", "qty"],
    { type: "number" },
  );
}

function findSalesAmountHeader(table = {}) {
  return findColumnHeader(
    table,
    [
      "순매출액",
      "매출액",
      "판매금액",
      "매출금액",
      "카드매출액",
      "sales",
      "revenue",
    ],
    { type: "number" },
  );
}

function uniqueTemplateCandidates(candidates = []) {
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

function buildYearWideVirtualTable({ table, yearHeaders = [] }) {
  const baseCategoryHeader =
    findColumnHeader(table, [
      "세부사업명",
      "제품명",
      "상품명",
      "품목",
      "업종",
      "구분",
    ]) || findColumnHeader(table, ["분류", "카테고리", "category"]);

  const groupHeader =
    findColumnHeader(table, ["구분", "제품류", "제품분류", "대분류"]) ||
    baseCategoryHeader;

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
        매출액: amount,
      });
    });
  });

  if (!rows.length) return null;

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_sales_year_wide`,
    tableName: `${table.tableName || table.sheetName || "매출"}_연도별변환`,
    columns: [
      { header: "구분", type: "category", role: "dimension" },
      { header: "항목", type: "category", role: "dimension" },
      { header: "연도", type: "number", role: "dimension" },
      { header: "매출액", type: "number", role: "metric" },
    ],
    rows,
  });
}

function buildMonthWideVirtualTable({ table, monthHeaders = [] }) {
  const yearHeader = findColumnHeader(table, [
    "기준년도",
    "연도",
    "년도",
    "year",
  ]);

  const regionHeader = findColumnHeader(table, [
    "구",
    "자치구",
    "지역",
    "시군구",
    "행정구",
    "region",
  ]);

  const categoryHeader = findColumnHeader(table, [
    "업종",
    "업태",
    "분류",
    "카테고리",
    "category",
  ]);

  if (!monthHeaders.length) return null;

  const rows = [];

  getRows(table).forEach((row) => {
    monthHeaders.forEach((monthHeader) => {
      const amount = toNumber(getRowValue(row, monthHeader));
      if (amount == null) return;

      const year = yearHeader ? getRowValue(row, yearHeader) : "";
      const month = normalizeMonth(monthHeader);
      const period = year && month ? `${year}-${month}` : month;

      rows.push({
        연도: year,
        월: month,
        연월: period,
        지역: regionHeader ? getRowValue(row, regionHeader) : "",
        업종: categoryHeader ? getRowValue(row, categoryHeader) : "",
        매출액: amount,
      });
    });
  });

  if (!rows.length) return null;

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_sales_month_wide`,
    tableName: `${table.tableName || table.sheetName || "매출"}_월별변환`,
    columns: [
      { header: "연도", type: "number", role: "dimension" },
      { header: "월", type: "category", role: "dimension" },
      { header: "연월", type: "category", role: "dimension" },
      { header: "지역", type: "category", role: "dimension" },
      { header: "업종", type: "category", role: "dimension" },
      { header: "매출액", type: "number", role: "metric" },
    ],
    rows,
  });
}

function buildLongYearMonthVirtualTable({ table, yearHeader, monthHeader }) {
  const quantityHeader = findSalesQuantityHeader(table);
  const amountHeader = findSalesAmountHeader(table);

  if (!yearHeader || !monthHeader || (!quantityHeader && !amountHeader)) {
    return null;
  }

  const rows = [];

  getRows(table).forEach((row) => {
    const year = normalizeYearValue(getRowValue(row, yearHeader));
    const month = normalizeMonthValue(getRowValue(row, monthHeader));

    if (!year || !month) return;

    const item = {
      연도: Number(year),
      월: month,
      연월: `${year}-${month}`,
    };

    if (quantityHeader) {
      const quantity = toNumber(getRowValue(row, quantityHeader));
      if (quantity != null) item.매출수량 = quantity;
    }

    if (amountHeader) {
      const amount = toNumber(getRowValue(row, amountHeader));
      if (amount != null) item.순매출액 = amount;
    }

    rows.push(item);
  });

  if (!rows.length) return null;

  const columns = [
    { header: "연도", type: "number", role: "dimension" },
    { header: "월", type: "category", role: "dimension" },
    { header: "연월", type: "category", role: "dimension" },
  ];

  if (quantityHeader) {
    columns.push({ header: "매출수량", type: "number", role: "metric" });
  }

  if (amountHeader) {
    columns.push({ header: "순매출액", type: "number", role: "metric" });
  }

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_sales_year_month`,
    tableName: `${table.tableName || table.sheetName || "매출"}_연월변환`,
    columns,
    rows,
  });
}

function createAverageSalesSection({ table }) {
  const rows = getRows(table);

  const quantityHeader = findSalesQuantityHeader(table);
  const amountHeader = findSalesAmountHeader(table);

  if (!quantityHeader || !amountHeader) return null;

  let totalQuantity = 0;
  let totalAmount = 0;

  rows.forEach((row) => {
    const quantity = toNumber(getRowValue(row, quantityHeader));
    const amount = toNumber(getRowValue(row, amountHeader));

    if (quantity != null) totalQuantity += quantity;
    if (amount != null) totalAmount += amount;
  });

  if (!totalQuantity || !totalAmount) return null;

  return {
    sectionId: "average_sales_amount",
    title: "평균 판매금액",
    candidate: {
      recipeType: "custom_metric",
      title: "평균 판매금액",
      tableId: table.tableId,
      columns: {
        quantity: quantityHeader,
        metric: amountHeader,
      },
    },
    result: {
      ok: true,
      recipeType: "custom_metric",
      title: "평균 판매금액",
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns: {
        quantity: quantityHeader,
        metric: amountHeader,
      },
      rows: [
        {
          지표: "평균 판매금액",
          매출수량합계: totalQuantity,
          매출액합계: totalAmount,
          값: totalAmount / totalQuantity,
        },
      ],
      rowCount: 1,
    },
  };
}

function buildLongSalesCandidates({ table }) {
  const tableId = table.tableId;

  const yearHeader = findColumnHeader(table, [
    "연도 구분",
    "기준년도",
    "매출연도",
    "연도",
    "년도",
    "year",
  ]);

  const monthHeader = findColumnHeader(table, [
    "월 구분",
    "기준월",
    "매출월",
    "월",
    "month",
  ]);

  const periodHeader = findColumnHeader(table, [
    "연월",
    "기간",
    "기준년월",
    "매출년월",
    "period",
  ]);

  const excludedDimensionHeaders = [
    yearHeader,
    monthHeader,
    periodHeader,
    "연도",
    "월",
    "연월",
  ];

  const productHeader = findNonTemporalColumnHeader(
    table,
    ["제품명", "상품명", "품목", "세부사업명", "제품류", "product", "item"],
    excludedDimensionHeaders,
  );

  const categoryHeader = findNonTemporalColumnHeader(
    table,
    ["제품분류", "상품분류", "업종", "업태", "구분", "분류", "category"],
    excludedDimensionHeaders,
  );

  const regionHeader = findNonTemporalColumnHeader(
    table,
    ["지역", "구", "자치구", "시군구", "region"],
    excludedDimensionHeaders,
  );

  const quantityHeader = findSalesQuantityHeader(table);
  const amountHeader = findSalesAmountHeader(table);

  const candidates = [];

  if (yearHeader && amountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${yearHeader}별 ${amountHeader} 요약`,
        tableId,
        columns: {
          dimension: yearHeader,
          metric: amountHeader,
        },
        meta: {
          sectionType: "yearly_sales",
        },
      }),
    );
  }

  if (periodHeader && amountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${periodHeader}별 ${amountHeader} 추이`,
        tableId,
        columns: {
          dimension: periodHeader,
          metric: amountHeader,
        },
        meta: {
          sectionType: "period_sales",
        },
      }),
    );
  }

  if (monthHeader && amountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${monthHeader}별 ${amountHeader} 추이`,
        tableId,
        columns: {
          dimension: monthHeader,
          metric: amountHeader,
        },
        meta: {
          sectionType: "monthly_sales",
        },
      }),
    );
  }

  if (monthHeader && quantityHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${monthHeader}별 ${quantityHeader} 추이`,
        tableId,
        columns: {
          dimension: monthHeader,
          metric: quantityHeader,
        },
        meta: {
          sectionType: "monthly_quantity",
        },
      }),
    );
  }

  if (productHeader && amountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${productHeader}별 ${amountHeader} 요약`,
        tableId,
        columns: {
          dimension: productHeader,
          metric: amountHeader,
        },
        meta: {
          sectionType: "product_sales",
        },
      }),
    );
  }

  if (categoryHeader && amountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${categoryHeader}별 ${amountHeader} 요약`,
        tableId,
        columns: {
          dimension: categoryHeader,
          metric: amountHeader,
        },
        meta: {
          sectionType: "category_sales",
        },
      }),
    );
  }

  if (regionHeader && amountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "group_summary",
        title: `${regionHeader}별 ${amountHeader} 요약`,
        tableId,
        columns: {
          dimension: regionHeader,
          metric: amountHeader,
        },
        meta: {
          sectionType: "region_sales",
        },
      }),
    );
  }

  const rankingDimension =
    periodHeader ||
    findColumnHeader(table, [
      "연월",
      "기간",
      "기준년월",
      "매출년월",
      "period",
    ]) ||
    productHeader ||
    categoryHeader ||
    regionHeader ||
    monthHeader ||
    yearHeader;

  if (rankingDimension && amountHeader) {
    candidates.push(
      makeTemplateCandidate({
        recipeType: "top_bottom",
        title: `${amountHeader} 상위/하위 항목`,
        tableId,
        columns: {
          dimension: rankingDimension,
          metric: amountHeader,
        },
        meta: {
          sectionType: "top_bottom_sales",
        },
      }),
    );
  }

  return uniqueTemplateCandidates(candidates);
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

  const yearWideHeaders = findWideYearHeaders(table);
  const monthWideHeaders = findWideMonthHeaders(table);

  let executionTables = normalizedQueryTables;
  let selectedTable = table;

  const yearWideVirtualTable = buildYearWideVirtualTable({
    table,
    yearHeaders: yearWideHeaders,
  });

  const monthWideVirtualTable = buildMonthWideVirtualTable({
    table,
    monthHeaders: monthWideHeaders,
  });

  const longYearHeader = findColumnHeader(table, [
    "연도 구분",
    "기준년도",
    "매출연도",
    "연도",
    "년도",
    "year",
  ]);

  const longMonthHeader = findColumnHeader(table, [
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

  if (monthWideVirtualTable) {
    selectedTable = monthWideVirtualTable;
    executionTables = [
      ...normalizedQueryTables.filter(
        (t) => t.tableId !== selectedTable.tableId,
      ),
      selectedTable,
    ];
  } else if (yearWideVirtualTable) {
    selectedTable = yearWideVirtualTable;
    executionTables = [
      ...normalizedQueryTables.filter(
        (t) => t.tableId !== selectedTable.tableId,
      ),
      selectedTable,
    ];
  } else if (longYearMonthVirtualTable) {
    selectedTable = longYearMonthVirtualTable;
    executionTables = [
      ...normalizedQueryTables.filter(
        (t) => t.tableId !== selectedTable.tableId,
      ),
      selectedTable,
    ];
  }

  const candidates = buildLongSalesCandidates({ table: selectedTable });
  const customSections = [
    createAverageSalesSection({ table: selectedTable }),
  ].filter(Boolean);

  if (!candidates.length && !customSections.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const recipeSections = executeTemplateSections({
    normalizedQueryTables: executionTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });

  return [...recipeSections, ...customSections];
}

module.exports = {
  executeSalesReport,
};
