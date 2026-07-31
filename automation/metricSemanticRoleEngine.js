const METRIC_SEMANTIC_ROLE_ENGINE_VERSION = "metric_semantic_role_engine_v1";
const AGGREGATION_CONTRACT_RESOLVER_VERSION =
  "aggregation_contract_resolver_v1";

const ROLE = Object.freeze({
  MONEY_FLOW: "money_flow",
  QUANTITY_FLOW: "quantity_flow",
  STOCK_SNAPSHOT: "stock_snapshot",
  UNIT_RATE: "unit_rate",
  DURATION: "duration",
  PERCENTAGE_RATE: "percentage_rate",
  COUNT: "count",
  GENERIC_MEASURE: "generic_measure",
});

const OPERATION = Object.freeze({
  SUM: "sum",
  AVERAGE: "average",
  LATEST: "latest",
});

const PERCENTAGE_PATTERN =
  /%|퍼센트|백분율|비율|비중|구성비|점유율|증감률|달성률|진행률|수익률|가동률|rate|ratio|share|percent|percentage/i;
const DURATION_PATTERN =
  /일수|소요일|소요기간|대여기간|연체기간|회수기간|처리기간|리드타임|lead\s*time|duration|elapsed|days?|hours?|months?/i;
const UNIT_RATE_PATTERN =
  /단가|평균단가|단위당|개당|건당|인당|시간당|가격|평균가격|평균비용|평균금액|평점|점수|만족도|지수|unit\s*price|price\s*per|per\s*unit|average\s*price|score|index/i;
const STOCK_SNAPSHOT_PATTERN =
  /현재재고|기초재고|기말재고|안전재고|재고수량|재고금액|잔여수량|잔량|현재수량|보유재고|재고잔액|재고잔량|on\s*hand|onhand|opening\s*stock|closing\s*stock|safety\s*stock|inventory\s*(?:quantity|amount|value|balance)?|stock\s*(?:quantity|amount|value|balance)?|remaining\s*(?:quantity|balance)?/i;
const COUNT_PATTERN =
  /건수|개수|횟수|인원수|명수|항목수|레코드수|row\s*count|record\s*count|count$/i;
const MONEY_PATTERN =
  /금액|비용|매출|매출액|매입|지출|예산|지원금|수입|수익|손익|원가|구매액|구매금액|취득가|취득금액|사용금액|물류비|출장비|교통비|숙박비|일비|에너지비용|투자비|amount|cost|revenue|expense|budget|sales|income|profit|fee/i;
const QUANTITY_PATTERN =
  /수량|사용량|입고량|입고수량|출고량|출고수량|이동량|이동수량|대여량|대여수량|생산량|판매량|소비량|절감량|전력사용량|가스사용량|quantity|volume|usage|consumption|units?/i;
const TOTALITY_PATTERN =
  /^(?:총|전체|누적)|(?:합계|총계)$|\b(?:total|grand\s*total|cumulative)\b/i;

function normalizeText(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKey(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "");
}

function normalizeDeclaredAggregation(value = "") {
  const text = normalizeText(value).toLowerCase();
  if (!text) return "";
  if (
    ["average", "avg", "mean", "평균"].includes(text) ||
    /평균|비율|지수|점수/.test(text)
  ) {
    return OPERATION.AVERAGE;
  }
  if (
    ["sum", "total", "합계", "합산"].includes(text) ||
    /합계|합산|총계/.test(text)
  ) {
    return OPERATION.SUM;
  }
  if (
    ["latest", "last", "snapshot", "최신", "최근", "스냅샷"].includes(text) ||
    /최신|최근|스냅샷|기말/.test(text)
  ) {
    return OPERATION.LATEST;
  }
  return "";
}

function declaredSemanticRole(column = {}) {
  const raw = normalizeKey(
    column.metricSemanticRole ||
      column.semanticMetricRole ||
      column.semanticRole ||
      column.meta?.metricSemanticRole ||
      "",
  );
  const aliases = new Map([
    ["moneyflow", ROLE.MONEY_FLOW],
    ["amountflow", ROLE.MONEY_FLOW],
    ["quantityflow", ROLE.QUANTITY_FLOW],
    ["flows", ROLE.QUANTITY_FLOW],
    ["stocksnapshot", ROLE.STOCK_SNAPSHOT],
    ["snapshot", ROLE.STOCK_SNAPSHOT],
    ["unitrate", ROLE.UNIT_RATE],
    ["rate", ROLE.UNIT_RATE],
    ["duration", ROLE.DURATION],
    ["percentagerate", ROLE.PERCENTAGE_RATE],
    ["percentage", ROLE.PERCENTAGE_RATE],
    ["count", ROLE.COUNT],
    ["genericmeasure", ROLE.GENERIC_MEASURE],
  ]);
  return aliases.get(raw) || "";
}

function classifyMetricRole({ metricLabel = "", unit = "", column = {} } = {}) {
  const explicitRole = declaredSemanticRole(column);
  const evidence = [
    metricLabel,
    unit,
    column.header,
    column.originalHeader,
    column.name,
    column.label,
    column.semanticType,
    column.role,
    column.type,
  ]
    .map(normalizeText)
    .filter(Boolean)
    .join(" ");

  const totality = TOTALITY_PATTERN.test(normalizeText(metricLabel))
    ? "total"
    : "detail";

  if (explicitRole) {
    return {
      role: explicitRole,
      totality,
      confidence: 1,
      source: "declared_column_role",
      evidence: normalizeText(metricLabel),
    };
  }

  const tests = [
    [ROLE.PERCENTAGE_RATE, PERCENTAGE_PATTERN, 0.98],
    [ROLE.DURATION, DURATION_PATTERN, 0.96],
    [ROLE.UNIT_RATE, UNIT_RATE_PATTERN, 0.95],
    [ROLE.STOCK_SNAPSHOT, STOCK_SNAPSHOT_PATTERN, 0.98],
    [ROLE.COUNT, COUNT_PATTERN, 0.94],
    [ROLE.MONEY_FLOW, MONEY_PATTERN, 0.9],
    [ROLE.QUANTITY_FLOW, QUANTITY_PATTERN, 0.9],
  ];

  for (const [role, pattern, confidence] of tests) {
    if (pattern.test(evidence)) {
      return {
        role,
        totality,
        confidence,
        source: "header_unit_pattern",
        evidence: normalizeText(metricLabel),
      };
    }
  }

  return {
    role: ROLE.GENERIC_MEASURE,
    totality,
    confidence: 0.5,
    source: "numeric_fallback",
    evidence: normalizeText(metricLabel),
  };
}

function roleDefaultOperation({
  role = ROLE.GENERIC_MEASURE,
  totality = "detail",
  hasTemporalAxis = false,
  fallbackOperation = OPERATION.SUM,
} = {}) {
  if (role === ROLE.STOCK_SNAPSHOT) {
    return hasTemporalAxis ? OPERATION.LATEST : OPERATION.SUM;
  }
  if (role === ROLE.UNIT_RATE || role === ROLE.PERCENTAGE_RATE) {
    return OPERATION.AVERAGE;
  }
  if (role === ROLE.DURATION) {
    return totality === "total" ? OPERATION.SUM : OPERATION.AVERAGE;
  }
  if (
    role === ROLE.MONEY_FLOW ||
    role === ROLE.QUANTITY_FLOW ||
    role === ROLE.COUNT
  ) {
    return OPERATION.SUM;
  }
  return normalizeDeclaredAggregation(fallbackOperation) || OPERATION.SUM;
}

function declaredOperationIsUnsafe({
  role = ROLE.GENERIC_MEASURE,
  totality = "detail",
  hasTemporalAxis = false,
  declaredOperation = "",
} = {}) {
  if (!declaredOperation) return false;
  if (
    role === ROLE.STOCK_SNAPSHOT &&
    hasTemporalAxis &&
    declaredOperation === OPERATION.SUM
  ) {
    return true;
  }
  if (
    (role === ROLE.UNIT_RATE || role === ROLE.PERCENTAGE_RATE) &&
    declaredOperation === OPERATION.SUM
  ) {
    return true;
  }
  if (
    role === ROLE.DURATION &&
    totality !== "total" &&
    declaredOperation === OPERATION.SUM
  ) {
    return true;
  }
  return false;
}

function resolveAggregationContract({
  metricLabel = "",
  unit = "",
  column = {},
  declaredAggregation = "",
  hasTemporalAxis = false,
  fallbackAggregation = OPERATION.SUM,
} = {}) {
  const classification = classifyMetricRole({
    metricLabel,
    unit,
    column,
  });
  const declaredOperation = normalizeDeclaredAggregation(
    declaredAggregation ||
      column.aggregation ||
      column.metricKind ||
      column.meta?.aggregation ||
      "",
  );
  const defaultOperation = roleDefaultOperation({
    role: classification.role,
    totality: classification.totality,
    hasTemporalAxis,
    fallbackOperation: fallbackAggregation,
  });
  const unsafeDeclaredAggregationOverridden = declaredOperationIsUnsafe({
    role: classification.role,
    totality: classification.totality,
    hasTemporalAxis,
    declaredOperation,
  });
  const operation =
    declaredOperation && !unsafeDeclaredAggregationOverridden
      ? declaredOperation
      : defaultOperation;

  return {
    ...classification,
    operation,
    additive: operation === OPERATION.SUM,
    hasTemporalAxis: Boolean(hasTemporalAxis),
    declaredOperation,
    defaultOperation,
    unsafeDeclaredAggregationOverridden,
    resolverVersion: AGGREGATION_CONTRACT_RESOLVER_VERSION,
    roleEngineVersion: METRIC_SEMANTIC_ROLE_ENGINE_VERSION,
  };
}

module.exports = {
  AGGREGATION_CONTRACT_RESOLVER_VERSION,
  METRIC_SEMANTIC_ROLE_ENGINE_VERSION,
  OPERATION,
  ROLE,
  classifyMetricRole,
  normalizeDeclaredAggregation,
  resolveAggregationContract,
  roleDefaultOperation,
};
