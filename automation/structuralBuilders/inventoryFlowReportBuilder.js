const {
  findColumnHeader,
  getColumns,
  getColumnHeader,
  makeTemplateCandidate,
  makeTemplateSection,
  executeTemplateSections,
  getRows,
  getRowValue,
  toNumber,
} = require("../businessTemplates/commonTemplateHelpers");

const INVENTORY_FLOW_REPORT_VERSION = "inventory_flow_report_builder_v2";

const DEFAULT_FLOW_TYPE_HINTS = [
  "입출고구분",
  "입출고",
  "거래구분",
  "이동구분",
  "처리구분",
  "구분",
  "유형",
  "분류",
  "flow",
  "transaction type",
  "movement type",
  "type",
];

const DEFAULT_INBOUND_QTY_HINTS = [
  "입고수량",
  "입고량",
  "입하수량",
  "입하량",
  "반입수량",
  "반입량",
  "수령수량",
  "구매수량",
  "생산수량",
  "inbound quantity",
  "received quantity",
  "in qty",
  "receipt qty",
];

const DEFAULT_OUTBOUND_QTY_HINTS = [
  "출고수량",
  "출고량",
  "출하수량",
  "출하량",
  "반출수량",
  "반출량",
  "사용수량",
  "판매수량",
  "불출수량",
  "outbound quantity",
  "shipped quantity",
  "out qty",
  "issue qty",
];

const DEFAULT_STOCK_QTY_HINTS = [
  "현재재고",
  "재고수량",
  "재고량",
  "기말재고",
  "기초재고",
  "잔여수량",
  "잔량",
  "보유수량",
  "가용재고",
  "stock",
  "inventory",
  "on hand",
  "balance quantity",
];

const DEFAULT_QUANTITY_HINTS = [
  "수량",
  "사용량",
  "이동수량",
  "거래수량",
  "처리수량",
  "조정수량",
  "quantity",
  "qty",
  "amount used",
];

const DEFAULT_AMOUNT_HINTS = [
  "재고금액",
  "입고금액",
  "출고금액",
  "구매금액",
  "취득금액",
  "평가금액",
  "합계금액",
  "총액",
  "금액",
  "inventory value",
  "extended cost",
  "total value",
  "value",
  "amount",
  "cost",
];

const DEFAULT_UNIT_PRICE_HINTS = [
  "단가",
  "개당가격",
  "개당금액",
  "개당원가",
  "unit price",
  "unit cost",
  "price per unit",
  "cost per unit",
  "price",
];

const DEFAULT_CATEGORY_HINTS = [
  "품목",
  "품목명",
  "상품",
  "제품",
  "자재",
  "소모품",
  "물품",
  "장비",
  "비품",
  "자산",
  "카테고리",
  "분류",
  "유형",
  "item",
  "product",
  "material",
  "asset",
  "equipment",
  "sku",
];

const DEFAULT_LOCATION_HINTS = [
  "창고",
  "보관위치",
  "위치",
  "센터",
  "지점",
  "매장",
  "부서",
  "사용부서",
  "location",
  "warehouse",
  "store",
  "branch",
  "department",
];

const DEFAULT_DATE_HINTS = [
  "일자",
  "날짜",
  "월",
  "연월",
  "기준월",
  "입고일",
  "출고일",
  "이동일",
  "처리일",
  "구매일",
  "취득일",
  "대여일",
  "반납일",
  "date",
  "month",
  "period",
];

const DEFAULT_STATUS_HINTS = [
  "상태",
  "처리상태",
  "재고상태",
  "장비상태",
  "대여상태",
  "보유상태",
  "검수상태",
  "status",
  "state",
];

const INBOUND_KEYWORDS = [
  "입고",
  "입하",
  "반입",
  "수령",
  "구매",
  "생산",
  "증가",
  "보충",
  "반납",
  "inbound",
  "in",
  "receive",
  "receipt",
  "purchase",
  "return",
];

const OUTBOUND_KEYWORDS = [
  "출고",
  "출하",
  "반출",
  "사용",
  "판매",
  "불출",
  "감소",
  "소진",
  "폐기",
  "대여",
  "outbound",
  "out",
  "ship",
  "issue",
  "use",
  "sale",
  "dispose",
  "rental",
];

const ADJUSTMENT_KEYWORDS = [
  "조정",
  "이동",
  "전환",
  "실사",
  "adjust",
  "move",
  "transfer",
];

function normalizeText(value = "") {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()\[\]{}]/g, "")
    .trim();
}

function includesAny(text = "", hints = []) {
  const normalized = normalizeText(text);
  return hints.some((hint) => {
    const target = normalizeText(hint);
    return target && normalized.includes(target);
  });
}

function safeRate(numerator = 0, denominator = 0) {
  const n = Number(numerator || 0);
  const d = Number(denominator || 0);
  if (!Number.isFinite(n) || !Number.isFinite(d) || d === 0) return null;
  return n / d;
}

function makePercent(value) {
  return value == null ? null : value * 100;
}

function sumNumbers(values = []) {
  return values.reduce((sum, value) => {
    const n = Number(value || 0);
    return Number.isFinite(n) ? sum + n : sum;
  }, 0);
}

function firstNonEmpty(values = []) {
  return values.find((value) => String(value ?? "").trim()) || "";
}

function normalizePeriod(value = "") {
  const raw = String(value ?? "").trim();
  if (!raw) return "미입력";

  const ymd = raw.match(/(20\d{2}|19\d{2})[.\-/년\s]*(0?[1-9]|1[0-2])?/);
  if (ymd) {
    const year = ymd[1];
    const month = ymd[2] ? String(ymd[2]).padStart(2, "0") : "";
    return month ? `${year}-${month}` : year;
  }

  return raw;
}

function classifyFlowType(value = "") {
  const text = normalizeText(value);
  if (!text) return "unknown";
  if (includesAny(text, INBOUND_KEYWORDS)) return "inbound";
  if (includesAny(text, OUTBOUND_KEYWORDS)) return "outbound";
  if (includesAny(text, ADJUSTMENT_KEYWORDS)) return "adjustment";
  return "other";
}

function numericValuesForHeader(table = {}, header = "") {
  if (!header) return [];
  return getRows(table)
    .map((row) => toNumber(getRowValue(row, header)))
    .filter((value) => value != null && Number.isFinite(Number(value)));
}

function findNumericHeaderByHints(
  table = {},
  hints = [],
  { excludeHints = [] } = {},
) {
  const columns = getColumns(table);

  const matched = columns.find((column) => {
    const header = getColumnHeader(column);

    // 힌트가 실제로 일치한 열만 허용한다.
    // 숫자형이라는 이유만으로 다른 열을 fallback하지 않는다.
    if (!includesAny(header, hints)) {
      return false;
    }

    if (includesAny(header, excludeHints)) {
      return false;
    }

    const values = numericValuesForHeader(table, header);

    return values.length > 0;
  });

  return matched ? getColumnHeader(matched) : "";
}

function findInventoryFlowHeaders(table = {}, config = {}) {
  const hints = config.hints || {};

  const flowTypeHeader = findColumnHeader(table, [
    ...(hints.flowType || []),
    ...DEFAULT_FLOW_TYPE_HINTS,
  ]);

  const inboundQuantityHeader = findNumericHeaderByHints(table, [
    ...(hints.inboundQuantity || []),
    ...DEFAULT_INBOUND_QTY_HINTS,
  ]);

  const outboundQuantityHeader = findNumericHeaderByHints(table, [
    ...(hints.outboundQuantity || []),
    ...DEFAULT_OUTBOUND_QTY_HINTS,
  ]);

  const stockQuantityHeader = findNumericHeaderByHints(table, [
    ...(hints.stockQuantity || []),
    ...DEFAULT_STOCK_QTY_HINTS,
  ]);

  const quantityHeader =
    inboundQuantityHeader ||
    outboundQuantityHeader ||
    stockQuantityHeader ||
    findNumericHeaderByHints(table, [
      ...(hints.quantity || []),
      ...DEFAULT_QUANTITY_HINTS,
    ]);

  const amountHeader = findNumericHeaderByHints(
    table,
    [...(hints.amount || []), ...(hints.value || []), ...DEFAULT_AMOUNT_HINTS],
    {
      // "단가"처럼 행 단위 가격인 열은 합계 금액으로 사용하지 않는다.
      excludeHints: DEFAULT_UNIT_PRICE_HINTS,
    },
  );

  const unitPriceHeader = findNumericHeaderByHints(table, [
    ...(hints.unitPrice || []),
    ...DEFAULT_UNIT_PRICE_HINTS,
  ]);

  // 기존 호출부와의 호환을 위해 valueHeader는 직접 금액 열만 가리킨다.
  const valueHeader = amountHeader;

  const categoryHeader = findColumnHeader(table, [
    ...(hints.category || []),
    ...DEFAULT_CATEGORY_HINTS,
  ]);

  const locationHeader = findColumnHeader(table, [
    ...(hints.location || []),
    ...DEFAULT_LOCATION_HINTS,
  ]);

  const dateHeader = findColumnHeader(table, [
    ...(hints.date || []),
    ...DEFAULT_DATE_HINTS,
  ]);

  const statusHeader = findColumnHeader(table, [
    ...(hints.status || []),
    ...DEFAULT_STATUS_HINTS,
  ]);

  return {
    flowTypeHeader,
    inboundQuantityHeader,
    outboundQuantityHeader,
    stockQuantityHeader,
    quantityHeader,
    amountHeader,
    unitPriceHeader,
    valueHeader,
    categoryHeader,
    locationHeader,
    dateHeader,
    statusHeader,
  };
}

function hasInventoryFlowEvidence(headers = {}) {
  return Boolean(
    headers.inboundQuantityHeader ||
    headers.outboundQuantityHeader ||
    headers.stockQuantityHeader ||
    (headers.flowTypeHeader && headers.quantityHeader) ||
    (headers.categoryHeader &&
      headers.quantityHeader &&
      headers.locationHeader),
  );
}

function rowFlowValues(row = {}, headers = {}) {
  const flowClass = classifyFlowType(getRowValue(row, headers.flowTypeHeader));
  const directInbound = toNumber(
    getRowValue(row, headers.inboundQuantityHeader),
  );
  const directOutbound = toNumber(
    getRowValue(row, headers.outboundQuantityHeader),
  );
  const stockQty = toNumber(getRowValue(row, headers.stockQuantityHeader));
  const genericQty = toNumber(getRowValue(row, headers.quantityHeader));
  const amount = toNumber(
    getRowValue(row, headers.amountHeader || headers.valueHeader),
  );
  const unitPrice = toNumber(getRowValue(row, headers.unitPriceHeader));

  let inboundQty = directInbound || 0;
  let outboundQty = directOutbound || 0;
  let adjustmentQty = 0;

  if (!directInbound && !directOutbound && genericQty != null) {
    if (flowClass === "inbound") inboundQty += Math.abs(genericQty);
    else if (flowClass === "outbound") outboundQty += Math.abs(genericQty);
    else if (flowClass === "adjustment") adjustmentQty += genericQty;
  }

  const netQty = inboundQty - outboundQty + adjustmentQty;
  const derivedStockValue =
    amount == null && stockQty != null && unitPrice != null
      ? stockQty * unitPrice
      : null;
  const stockValue = amount != null ? amount : derivedStockValue;

  return {
    flowClass,
    inboundQty,
    outboundQty,
    adjustmentQty,
    stockQty: stockQty || 0,
    genericQty: genericQty || 0,
    netQty,
    amount: amount || 0,
    unitPrice: unitPrice || 0,
    stockValue: stockValue || 0,
    stockValueSource:
      amount != null
        ? "direct_amount"
        : derivedStockValue != null
          ? "stock_quantity_x_unit_price"
          : "unavailable",
    // 기존 집계 코드와의 호환 필드. 의미는 직접 금액 또는 파생 재고평가금액이다.
    value: stockValue || 0,
  };
}

function makeCustomMetricSection({
  sectionId,
  sectionType,
  title,
  table,
  rows,
  columns = {},
  chartHint = {},
  narrativeHint = {},
  meta = {},
}) {
  return makeTemplateSection({
    sectionId,
    sectionType,
    title,
    candidate: {
      recipeType: "custom_metric",
      title,
      tableId: table.tableId,
      columns,
      meta: {
        ...meta,
        inventoryFlowReportVersion: INVENTORY_FLOW_REPORT_VERSION,
      },
    },
    result: {
      ok: true,
      recipeType: "custom_metric",
      resultType: sectionType,
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
      meta: {
        ...meta,
        inventoryFlowReportVersion: INVENTORY_FLOW_REPORT_VERSION,
      },
    },
    chartHint,
    narrativeHint,
  });
}

function buildInventoryOverviewSection({ table, headers, config = {} }) {
  if (!table?.tableId || !hasInventoryFlowEvidence(headers)) return null;

  const rows = getRows(table);
  if (!rows.length) return null;

  const flowValues = rows.map((row) => rowFlowValues(row, headers));
  const inboundQty = sumNumbers(flowValues.map((item) => item.inboundQty));
  const outboundQty = sumNumbers(flowValues.map((item) => item.outboundQty));
  const adjustmentQty = sumNumbers(
    flowValues.map((item) => item.adjustmentQty),
  );
  const stockQty = sumNumbers(flowValues.map((item) => item.stockQty));
  const totalStockValue = sumNumbers(flowValues.map((item) => item.stockValue));
  const stockValueSource = firstNonEmpty(
    flowValues
      .map((item) => item.stockValueSource)
      .filter((value) => value && value !== "unavailable"),
  );
  const netQty = inboundQty - outboundQty + adjustmentQty;
  const estimatedStock = stockQty || netQty;

  const resultRows = [
    { 지표: "전체 행 수", 값: rows.length, 보조값: null, 비율Percent: null },
    {
      지표: config.labels?.inbound || "입고 수량",
      값: inboundQty,
      보조값: null,
      비율Percent: makePercent(safeRate(inboundQty, inboundQty + outboundQty)),
    },
    {
      지표: config.labels?.outbound || "출고 수량",
      값: outboundQty,
      보조값: null,
      비율Percent: makePercent(safeRate(outboundQty, inboundQty + outboundQty)),
    },
    { 지표: "순증감 수량", 값: netQty, 보조값: null, 비율Percent: null },
  ];

  if (stockQty) {
    resultRows.push({
      지표: "재고 수량 합계",
      값: stockQty,
      보조값: null,
      비율Percent: null,
    });
  } else {
    resultRows.push({
      지표: "추정 재고 증감",
      값: estimatedStock,
      보조값: null,
      비율Percent: null,
    });
  }

  if (totalStockValue) {
    resultRows.push({
      지표:
        stockValueSource === "stock_quantity_x_unit_price"
          ? "총 재고평가금액"
          : `${headers.amountHeader || "금액"} 합계`,
      값: totalStockValue,
      보조값: null,
      비율Percent: null,
    });
  }

  return makeCustomMetricSection({
    sectionId: config.sectionIds?.overview || "inventory_flow_overview",
    sectionType: "inventory_flow_overview",
    title: config.titles?.overview || "재고·입출고 흐름 요약",
    table,
    rows: resultRows,
    columns: {
      inboundQuantity: headers.inboundQuantityHeader,
      outboundQuantity: headers.outboundQuantityHeader,
      stockQuantity: headers.stockQuantityHeader,
      quantity: headers.quantityHeader,
      amount: headers.amountHeader,
      unitPrice: headers.unitPriceHeader,
      stockValue:
        stockValueSource === "stock_quantity_x_unit_price"
          ? "현재재고 × 단가"
          : headers.amountHeader,
      value: headers.valueHeader,
    },
    chartHint: {
      preferredType: "metric_card",
      valueField: "값",
      ratioField: "비율Percent",
    },
    narrativeHint: {
      focus: "inventory_flow_overview",
    },
  });
}

function buildFlowTypeBreakdownSection({ table, headers, config = {} }) {
  if (!table?.tableId || !hasInventoryFlowEvidence(headers)) return null;

  const map = new Map();
  getRows(table).forEach((row) => {
    const values = rowFlowValues(row, headers);
    const label =
      values.flowClass === "inbound"
        ? "입고"
        : values.flowClass === "outbound"
          ? "출고"
          : values.flowClass === "adjustment"
            ? "조정·이동"
            : values.flowClass === "unknown"
              ? "미분류"
              : "기타";

    if (!map.has(label)) {
      map.set(label, {
        구분: label,
        건수: 0,
        입고수량: 0,
        출고수량: 0,
        조정수량: 0,
        순증감수량: 0,
        금액: 0,
      });
    }

    const item = map.get(label);
    item.건수 += 1;
    item.입고수량 += values.inboundQty;
    item.출고수량 += values.outboundQty;
    item.조정수량 += values.adjustmentQty;
    item.순증감수량 += values.netQty;
    item.금액 += values.value;
  });

  const totalCount = sumNumbers(
    Array.from(map.values()).map((item) => item.건수),
  );
  const resultRows = Array.from(map.values())
    .map((item) => ({
      ...item,
      구성비: safeRate(item.건수, totalCount),
      구성비Percent: makePercent(safeRate(item.건수, totalCount)),
    }))
    .sort((a, b) => Number(b.건수 || 0) - Number(a.건수 || 0));

  if (!resultRows.length) return null;

  return makeCustomMetricSection({
    sectionId:
      config.sectionIds?.flowBreakdown || "inventory_flow_type_breakdown",
    sectionType: "inventory_flow_type_breakdown",
    title: config.titles?.flowBreakdown || "입출고 구분별 흐름",
    table,
    rows: resultRows,
    columns: {
      flowType: headers.flowTypeHeader,
      inboundQuantity: "입고수량",
      outboundQuantity: "출고수량",
      netQuantity: "순증감수량",
      value: "금액",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: "구분",
      valueField: "순증감수량",
    },
    narrativeHint: {
      focus: "inventory_flow_type_breakdown",
    },
  });
}

function buildFlowByPeriodSection({ table, headers, config = {} }) {
  const { dateHeader } = headers || {};
  if (!table?.tableId || !dateHeader || !hasInventoryFlowEvidence(headers))
    return null;

  const map = new Map();
  getRows(table).forEach((row) => {
    const period = normalizePeriod(getRowValue(row, dateHeader));
    const values = rowFlowValues(row, headers);

    if (!map.has(period)) {
      map.set(period, {
        기간: period,
        건수: 0,
        입고수량: 0,
        출고수량: 0,
        조정수량: 0,
        순증감수량: 0,
        재고수량: 0,
        금액: 0,
      });
    }

    const item = map.get(period);
    item.건수 += 1;
    item.입고수량 += values.inboundQty;
    item.출고수량 += values.outboundQty;
    item.조정수량 += values.adjustmentQty;
    item.순증감수량 += values.netQty;
    item.재고수량 += values.stockQty;
    item.금액 += values.value;
  });

  const resultRows = Array.from(map.values()).sort((a, b) =>
    String(a.기간).localeCompare(String(b.기간)),
  );
  if (!resultRows.length) return null;

  return makeCustomMetricSection({
    sectionId: config.sectionIds?.periodFlow || "inventory_flow_by_period",
    sectionType: "inventory_flow_by_period",
    title: config.titles?.periodFlow || `${dateHeader}별 입출고 흐름`,
    table,
    rows: resultRows,
    columns: {
      date: dateHeader,
      inboundQuantity: "입고수량",
      outboundQuantity: "출고수량",
      netQuantity: "순증감수량",
      stockQuantity: "재고수량",
      value: "금액",
    },
    chartHint: {
      preferredType: "line",
      categoryField: "기간",
      valueField: "순증감수량",
    },
    narrativeHint: {
      focus: "inventory_flow_by_period",
      date: dateHeader,
    },
  });
}

function buildInventoryByDimensionSection({
  table,
  headers,
  dimensionHeader = "",
  title = "",
  sectionId = "",
}) {
  if (!table?.tableId || !dimensionHeader || !hasInventoryFlowEvidence(headers))
    return null;

  const map = new Map();
  getRows(table).forEach((row) => {
    const dimension =
      String(getRowValue(row, dimensionHeader) ?? "").trim() || "미입력";
    const values = rowFlowValues(row, headers);

    if (!map.has(dimension)) {
      map.set(dimension, {
        [dimensionHeader]: dimension,
        건수: 0,
        입고수량: 0,
        출고수량: 0,
        조정수량: 0,
        순증감수량: 0,
        재고수량: 0,
        금액: 0,
      });
    }

    const item = map.get(dimension);
    item.건수 += 1;
    item.입고수량 += values.inboundQty;
    item.출고수량 += values.outboundQty;
    item.조정수량 += values.adjustmentQty;
    item.순증감수량 += values.netQty;
    item.재고수량 += values.stockQty;
    item.금액 += values.value;
  });

  const resultRows = Array.from(map.values())
    .map((item) => ({
      ...item,
      출고율: makePercent(
        safeRate(item.출고수량, item.입고수량 + item.재고수량),
      ),
    }))
    .sort(
      (a, b) =>
        Number((b.재고수량 || 0) + (b.순증감수량 || 0)) -
        Number((a.재고수량 || 0) + (a.순증감수량 || 0)),
    );

  if (!resultRows.length) return null;

  return makeCustomMetricSection({
    sectionId: sectionId || `inventory_flow_by_${dimensionHeader}`,
    sectionType: "inventory_flow_by_dimension",
    title: title || `${dimensionHeader}별 재고·입출고 요약`,
    table,
    rows: resultRows,
    columns: {
      dimension: dimensionHeader,
      inboundQuantity: "입고수량",
      outboundQuantity: "출고수량",
      stockQuantity: "재고수량",
      netQuantity: "순증감수량",
      value: "금액",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: headers.stockQuantityHeader ? "재고수량" : "순증감수량",
    },
    narrativeHint: {
      focus: "inventory_flow_by_dimension",
      dimension: dimensionHeader,
    },
  });
}

function buildInventoryFlowCandidates({ table, headers, config = {} }) {
  if (!table?.tableId) return [];

  const {
    dateHeader,
    categoryHeader,
    locationHeader,
    statusHeader,
    flowTypeHeader,
    quantityHeader,
    stockQuantityHeader,
    amountHeader,
    valueHeader,
  } = headers || {};

  // 단가는 합산 가능한 대표 metric이 아니다.
  // 재고수량·수량·직접 금액 열만 일반 recipe 후보로 사용한다.
  const metricHeader =
    stockQuantityHeader || quantityHeader || amountHeader || valueHeader;

  const candidates = [];
  const tableId = table.tableId;

  if (dateHeader && metricHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.timeQuantity || "inventory_time_quantity",
        sectionType: "inventory_time_quantity",
        recipeType: "time_sum",
        title:
          config.titles?.timeQuantity || `${dateHeader}별 ${metricHeader} 추이`,
        tableId,
        columns: {
          date: dateHeader,
          metric: metricHeader,
        },
        chartHint: {
          preferredType: "line",
          categoryField: dateHeader,
          valueField: metricHeader,
        },
        narrativeHint: {
          focus: "inventory_time_sum",
          date: dateHeader,
          metric: metricHeader,
        },
      }),
    );
  }

  const dimensions = [
    categoryHeader,
    locationHeader,
    statusHeader,
    flowTypeHeader,
  ].filter((value, index, arr) => value && arr.indexOf(value) === index);

  for (const dimensionHeader of dimensions) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: `inventory_count_by_${dimensionHeader}`,
        sectionType: "inventory_dimension_count",
        recipeType: "category_count",
        title: `${dimensionHeader}별 건수`,
        tableId,
        columns: {
          dimension: dimensionHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: dimensionHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "inventory_dimension_count",
          dimension: dimensionHeader,
        },
      }),
    );

    if (metricHeader) {
      candidates.push(
        makeTemplateCandidate({
          sectionId: `inventory_metric_by_${dimensionHeader}`,
          sectionType: "inventory_dimension_metric",
          recipeType: "group_sum",
          title: `${dimensionHeader}별 ${metricHeader} 합계`,
          tableId,
          columns: {
            dimension: dimensionHeader,
            metric: metricHeader,
          },
          chartHint: {
            preferredType: "bar",
            categoryField: dimensionHeader,
            valueField: metricHeader,
          },
          narrativeHint: {
            focus: "inventory_dimension_metric",
            dimension: dimensionHeader,
            metric: metricHeader,
          },
        }),
      );
    }
  }

  if (categoryHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId:
          config.sectionIds?.composition || "inventory_category_composition",
        sectionType: "inventory_category_composition",
        recipeType: "composition_ratio",
        title: `${categoryHeader} 구성비`,
        tableId,
        columns: {
          dimension: categoryHeader,
          metric: metricHeader,
        },
        chartHint: {
          preferredType: "donut",
          categoryField: categoryHeader,
          valueField: metricHeader,
        },
        narrativeHint: {
          focus: "inventory_composition",
          dimension: categoryHeader,
        },
      }),
    );
  }

  if ((categoryHeader || locationHeader) && metricHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.topBottom || "inventory_top_bottom",
        sectionType: "inventory_top_bottom",
        recipeType: "top_bottom",
        title: `${metricHeader} 상위·하위 항목`,
        tableId,
        columns: {
          dimension: categoryHeader || locationHeader,
          metric: metricHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: categoryHeader || locationHeader,
          valueField: metricHeader,
        },
        narrativeHint: {
          focus: "top_bottom",
          metric: metricHeader,
        },
      }),
    );
  }

  return candidates.filter((candidate) => {
    if (!candidate.columns) return true;
    return Object.values(candidate.columns).every(Boolean);
  });
}

function buildInventoryFlowReportSections({
  normalizedQueryTables = [],
  table,
  templateCandidate = {},
  config = {},
}) {
  if (!table?.tableId) return [];

  const headers = findInventoryFlowHeaders(table, config);

  if (!hasInventoryFlowEvidence(headers)) {
    const fallbackCandidates = Array.isArray(templateCandidate.candidates)
      ? templateCandidate.candidates
      : [];

    if (!fallbackCandidates.length) return [];

    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const customSections = [
    buildInventoryOverviewSection({ table, headers, config }),
    buildFlowTypeBreakdownSection({ table, headers, config }),
    buildFlowByPeriodSection({ table, headers, config }),
    buildInventoryByDimensionSection({
      table,
      headers,
      dimensionHeader: headers.categoryHeader,
      sectionId: "inventory_flow_by_category",
      title: headers.categoryHeader
        ? `${headers.categoryHeader}별 재고·입출고 요약`
        : "품목별 재고·입출고 요약",
    }),
    buildInventoryByDimensionSection({
      table,
      headers,
      dimensionHeader: headers.locationHeader,
      sectionId: "inventory_flow_by_location",
      title: headers.locationHeader
        ? `${headers.locationHeader}별 재고·입출고 요약`
        : "위치별 재고·입출고 요약",
    }),
  ].filter(Boolean);

  const candidates = buildInventoryFlowCandidates({ table, headers, config });
  const recipeSections = executeTemplateSections({
    normalizedQueryTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });

  return [...customSections, ...recipeSections];
}

module.exports = {
  INVENTORY_FLOW_REPORT_VERSION,
  classifyFlowType,
  findInventoryFlowHeaders,
  buildInventoryFlowCandidates,
  buildInventoryFlowReportSections,
};
