const crypto = require("crypto");
const {
  METRIC_ID_CONTRACT_VERSION,
  applySectionMetricIds,
  collectSectionMetricIds,
  normalizeSectionMetricIds,
  uniqueMetricIds,
} = require("./metricIdContract");

const SEMANTIC_OUTPUT_PLANNER_VERSION =
  "semantic_output_planner_common_v2_2_generic_section_cleanup";
const SEMANTIC_OUTPUT_CONTRACT_VERSION = "semantic_output_contract_v1";

const SUMMARY_LABEL_PATTERN =
  /^(?:계|합계|소계|총계|전체|전국|세계|total|subtotal|grand\s*total)$/i;
const SUMMARY_SUFFIX_PATTERN = /(?:합계|소계|총계)\s*$/;
const IDENTIFIER_HEADER_PATTERN =
  /(?:^|[_\s])(id|code)(?:$|[_\s])|번호|코드|순번|연번/i;
const EXCLUDED_DIMENSION_HEADER_PATTERN =
  /^(?:단위|출처|주석|비고|설명|테스트\s*목적|source|note|unit)$/i;
const METRIC_IDENTITY_HEADER_PATTERN =
  /^(?:지표명|지표|항목|측정항목|세부항목|metric|measure|indicator)$/i;
const METRIC_VALUE_HEADER_PATTERN =
  /^(?:지표값|값|수치|metric\s*value|measure\s*value|value)$/i;
const UNIT_HEADER_PATTERN = /^(?:단위|측정단위|unit|measure\s*unit)$/i;
const AGGREGATION_HEADER_PATTERN =
  /^(?:집계유형|집계방식|aggregation|aggregate)$/i;
const PERIOD_HEADER_PATTERN =
  /^(?:기간|기준기간|년월|연월|날짜|일자|period|date|month|quarter)$/i;
const YEAR_HEADER_PATTERN = /^(?:연도|년도|year)$/i;
const GENERIC_METRIC_LABEL_PATTERN =
  /^(?:지표값|값|수치|실적|metric|measure|value)$/i;
const BUSINESS_TIME_HEADER_PATTERN =
  /기간|년월|연월|월|일자|날짜|date|month|quarter/i;
const BUSINESS_SUM_HEADER_PATTERN =
  /금액|매출|매출액|비용|지출|예산|지원금|수량|건수|인원|횟수|amount|cost|revenue|budget|count|quantity/i;
const BUSINESS_AVERAGE_HEADER_PATTERN =
  /점수|만족도|진행률|비율|평균|평점|내용연수|수명|score|rate|ratio|average/i;

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

function cloneValue(value) {
  return JSON.parse(JSON.stringify(value));
}

function tableRows(table = {}) {
  return Array.isArray(table.rows) ? table.rows : [];
}

function tableColumns(table = {}) {
  return Array.isArray(table.columns) ? table.columns : [];
}

function tableLabel(table = {}, index = 0) {
  return (
    normalizeText(
      table.tableName || table.sheetName || table.title || table.tableId || "",
    ) || `표 ${index + 1}`
  );
}

function semanticContextText(table = {}, context = {}) {
  return [
    tableLabel(table),
    table.fileName,
    table.sourceFileName,
    table.description,
    context.templateId,
    context.templateTitle,
    context.templateDescription,
    context.message,
  ]
    .map(normalizeText)
    .filter(Boolean)
    .join(" ");
}

function isVirtualSemanticTable(table = {}) {
  return Boolean(
    table.isVirtual === true ||
    table.virtual === true ||
    table.transformation?.type ||
    table.sourceTableId,
  );
}

function semanticSourceTableId(table = {}) {
  return normalizeText(
    table.sourceTableId ||
      table.transformation?.sourceTableId ||
      table.meta?.sourceTableId ||
      "",
  );
}

function semanticTableId(table = {}) {
  return normalizeText(table.tableId || table.id || "");
}

function isSemanticAnalysisEligible(table = {}) {
  if (table.tableUsage?.analysisEligible === false) return false;
  if (table.analysisEligible === false) return false;
  return tableRows(table).length > 0 && tableColumns(table).length > 0;
}

function semanticDimensionHeaderSet(table = {}) {
  const canonical = canonicalLongContract(table);
  if (canonical) {
    return new Set(
      canonical.dimensions.map((entry) => normalizeKey(entry.header)),
    );
  }

  const physical = physicalWideContract(table);
  if (physical) {
    return new Set(
      physical.dimensions.map((entry) => normalizeKey(entry.header)),
    );
  }

  return new Set();
}

function selectPreferredSemanticTables(tables = [], options = {}) {
  const eligible = (Array.isArray(tables) ? tables : []).filter(
    isSemanticAnalysisEligible,
  );
  const virtual = eligible.filter(isVirtualSemanticTable);
  if (!virtual.length) return eligible;

  const virtualBySource = new Map();
  for (const table of virtual) {
    const sourceId = semanticSourceTableId(table);
    if (!sourceId) continue;
    if (!virtualBySource.has(sourceId)) {
      virtualBySource.set(sourceId, []);
    }
    virtualBySource.get(sourceId).push(table);
  }

  const physicalPreferredSources = new Set();

  if (options.preferDimensionCompletePhysical === true) {
    for (const table of eligible) {
      if (isVirtualSemanticTable(table)) continue;
      const tableId = semanticTableId(table);
      const representedBy = virtualBySource.get(tableId) || [];
      if (!tableId || !representedBy.length) continue;

      const physicalDimensions = semanticDimensionHeaderSet(table);
      const virtualDimensions = new Set(
        representedBy.flatMap((item) => [...semanticDimensionHeaderSet(item)]),
      );
      const losesDimension = [...physicalDimensions].some(
        (header) => !virtualDimensions.has(header),
      );

      if (losesDimension) {
        physicalPreferredSources.add(tableId);
      }
    }
  }

  return eligible.filter((table) => {
    if (isVirtualSemanticTable(table)) {
      const sourceId = semanticSourceTableId(table);
      return !sourceId || !physicalPreferredSources.has(sourceId);
    }

    const id = semanticTableId(table);
    if (!id || !virtualBySource.has(id)) return true;
    return physicalPreferredSources.has(id);
  });
}

function inferContextualMetricLabel({
  metricLabel = "",
  unit = "",
  table = {},
  context = {},
} = {}) {
  const label = normalizeText(metricLabel);
  if (label && !GENERIC_METRIC_LABEL_PATTERN.test(label)) return label;

  const evidence = semanticContextText(table, context);
  const moneyLike = /원|만원|억원|천원|금액|amount|revenue|cost/i.test(
    `${unit} ${evidence}`,
  );

  if (/매출|판매|sales|revenue/i.test(evidence)) return "매출액";
  if (/월간\s*지출|지출\s*리포트|expense/i.test(evidence)) {
    return "지출금액";
  }
  if (/회의비|meeting\s*expense/i.test(evidence)) return "사용금액";
  if (/거래처|업체|vendor|거래\s*실적/i.test(evidence) && moneyLike) {
    return "거래금액";
  }
  if (/지원사업|신청|grant|application/i.test(evidence) && moneyLike) {
    return "신청금액";
  }
  if (/자산|취득|asset/i.test(evidence) && moneyLike) return "취득금액";

  return label || "지표값";
}

function columnHeader(column = {}, index = 0) {
  return normalizeText(
    column.header ||
      column.originalHeader ||
      column.name ||
      column.label ||
      `열${index + 1}`,
  );
}

function rowValue(row, column = {}, index = 0) {
  if (Array.isArray(row)) return row[index];
  if (!row || typeof row !== "object") return undefined;

  const keys = [
    column.key,
    column.canonicalKey,
    column.accessor,
    column.name,
    column.header,
    column.originalHeader,
    column.label,
    column.id,
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return row[key];
    }
  }

  const normalizedTargets = new Set(keys.map(normalizeKey));
  for (const [key, value] of Object.entries(row)) {
    if (normalizedTargets.has(normalizeKey(key))) return value;
  }

  return Object.values(row)[index];
}

function numericValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const source = normalizeText(value);
  if (!source || source === "-") return null;

  const normalized = source.replace(/,/g, "").replace(/%$/g, "").trim();

  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) {
    return null;
  }

  const result = Number(normalized);
  return Number.isFinite(result) ? result : null;
}

function semanticType(column = {}) {
  return normalizeText(
    column.semanticType ||
      column.semantic?.type ||
      column.meta?.semanticType ||
      "",
  ).toLowerCase();
}

function columnRole(column = {}) {
  return normalizeText(column.role || column.meta?.role || "").toLowerCase();
}

function extractHeaderUnit(header = "") {
  const matches = [
    ...normalizeText(header).matchAll(
      /(?:\(|\[|（)\s*([^()[\]（）]{1,32})\s*(?:\)|\]|）)/g,
    ),
  ];
  return matches.length ? normalizeText(matches[matches.length - 1][1]) : "";
}

function stripHeaderUnit(header = "") {
  return normalizeText(header)
    .replace(/\([^)]*\)/g, " ")
    .replace(/\[[^\]]*\]/g, " ")
    .replace(/（[^）]*）/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function strictPeriodValue(value) {
  const source = normalizeText(value);
  if (!source) return "";

  if (/^(?:19|20|21)\d{2}$/.test(source)) return source;

  let yearOnly = source.match(/^((?:19|20|21)\d{2})\s*년$/);
  if (yearOnly) return yearOnly[1];

  let monthOnly = source.match(/^(0?[1-9]|1[0-2])\s*월$/);
  if (monthOnly) {
    return `${String(Number(monthOnly[1])).padStart(2, "0")}월`;
  }

  let match = source.match(/^((?:19|20|21)\d{2})[-./]\s*(0?[1-9]|1[0-2])$/);
  if (match) {
    return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
  }

  match = source.match(/^((?:19|20|21)\d{2})\s*년\s*(0?[1-9]|1[0-2])\s*월$/);
  if (match) {
    return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
  }

  match = source.match(/^((?:19|20|21)\d{2})\s*(?:Q([1-4])|([1-4])\s*분기)$/i);
  if (match) return `${match[1]}-Q${match[2] || match[3]}`;

  match = source.match(
    /^((?:19|20|21)\d{2})[-./]\s*(0?[1-9]|1[0-2])[-./]\s*(0?[1-9]|[12]\d|3[01])$/,
  );
  if (match) {
    return [
      match[1],
      String(Number(match[2])).padStart(2, "0"),
      String(Number(match[3])).padStart(2, "0"),
    ].join("-");
  }

  return "";
}

function canonicalPeriodValue(periodValue = "", yearValue = "") {
  const period = strictPeriodValue(periodValue);
  const year = strictPeriodValue(yearValue);

  if (/^\d{2}월$/.test(period) && /^\d{4}$/.test(year)) {
    return `${year}-${period.slice(0, 2)}`;
  }

  return period || year;
}

function parseTemporalMeasureHeader(header = "") {
  const raw = normalizeText(header);
  if (!raw) return { period: "", metricLabel: "", unit: "" };

  const unit = extractHeaderUnit(raw);
  const withoutUnit = stripHeaderUnit(raw);
  const patterns = [
    {
      regex:
        /(?:^|[_\s])((?:19|20|21)\d{2})\s*년\s*(0?[1-9]|1[0-2])\s*월(?:[_\s]|$)/,
      period: (match) =>
        `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`,
    },
    {
      regex: /(?:^|[_\s])(0?[1-9]|1[0-2])\s*월(?:[_\s]|$)/,
      period: (match) => `${String(Number(match[1])).padStart(2, "0")}월`,
    },
    {
      regex: /(?:^|[_\s])((?:19|20|21)\d{2})[./-](0?[1-9]|1[0-2])(?:[_\s]|$)/,
      period: (match) =>
        `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`,
    },
    {
      regex:
        /(?:^|[_\s])((?:19|20|21)\d{2})\s*(?:Q([1-4])|([1-4])\s*분기)(?:[_\s]|$)/i,
      period: (match) => `${match[1]}-Q${match[2] || match[3]}`,
    },
    {
      regex: /(?:^|[_\s])((?:19|20|21)\d{2})(?:년)?(?:[_\s]|$)/,
      period: (match) => match[1],
    },
  ];

  for (const spec of patterns) {
    const match = withoutUnit.match(spec.regex);
    if (!match) continue;

    const metricLabel = normalizeText(
      `${withoutUnit.slice(0, match.index)} ${withoutUnit.slice(
        Number(match.index || 0) + match[0].length,
      )}`,
    )
      .replace(/^[|/_\-–—:]+|[|/_\-–—:]+$/g, "")
      .trim();

    return {
      period: spec.period(match),
      metricLabel,
      unit,
    };
  }

  return {
    period: "",
    metricLabel: withoutUnit,
    unit,
  };
}

function isSummaryLabel(value = "") {
  const text = normalizeText(value);
  return SUMMARY_LABEL_PATTERN.test(text) || SUMMARY_SUFFIX_PATTERN.test(text);
}

function summaryMetricBase(value = "") {
  return normalizeText(value)
    .replace(/(?:[_/|>:\-]\s*|\s+)(?:계|합계|소계|총계)\s*$/i, "")
    .trim();
}
function officialTotalHasDetailSiblings(value = "", allLabels = []) {
  const text = normalizeText(value);
  if (!/\s+합계\s*$/i.test(text)) return false;
  if (/[_/|>:\-]\s*합계\s*$/i.test(text)) return false;
  const base = summaryMetricBase(text);
  if (!base) return false;
  return (
    (allLabels || []).filter((candidate) => {
      const label = normalizeText(candidate);
      return (
        label &&
        label !== text &&
        label.startsWith(`${base} `) &&
        !/(?:계|합계|소계|총계)\s*$/i.test(label)
      );
    }).length >= 2
  );
}
function isSummaryMetricLabel(value = "", allLabels = []) {
  const text = normalizeText(value);
  if (!text) return false;
  if (SUMMARY_LABEL_PATTERN.test(text)) return true;
  if (/(?:[_/|>:\-]\s*|\s+)(?:소계|총계)\s*$/i.test(text)) return true;
  if (/(?:[_/|>:\-]\s*)(?:계|합계)\s*$/i.test(text)) return true;
  if (/\s+합계\s*$/i.test(text))
    return !officialTotalHasDetailSiblings(text, allLabels);
  return false;
}

function isIdentifierColumn(column = {}, header = "") {
  const role = columnRole(column);
  const semantic = semanticType(column);
  return (
    ["id", "identifier", "code"].includes(role) ||
    ["id", "identifier", "code"].includes(semantic) ||
    IDENTIFIER_HEADER_PATTERN.test(normalizeText(header))
  );
}

function isMetricIdentityColumn(column = {}, header = "") {
  return (
    semanticType(column) === "metricidentity" ||
    columnRole(column) === "metricidentity" ||
    METRIC_IDENTITY_HEADER_PATTERN.test(normalizeText(header))
  );
}

function isMetricValueColumn(column = {}, header = "") {
  return (
    semanticType(column) === "measure" ||
    columnRole(column) === "metric" ||
    METRIC_VALUE_HEADER_PATTERN.test(normalizeText(header))
  );
}

function isUnitColumn(column = {}, header = "") {
  return (
    semanticType(column) === "unit" ||
    UNIT_HEADER_PATTERN.test(normalizeText(header))
  );
}

function isAggregationColumn(column = {}, header = "") {
  return (
    semanticType(column) === "aggregation" ||
    AGGREGATION_HEADER_PATTERN.test(normalizeText(header))
  );
}

function isPeriodColumn(column = {}, header = "") {
  return (
    ["period", "date", "datetime", "month", "quarter"].includes(
      semanticType(column),
    ) || PERIOD_HEADER_PATTERN.test(normalizeText(header))
  );
}

function isYearColumn(column = {}, header = "") {
  return (
    semanticType(column) === "year" ||
    YEAR_HEADER_PATTERN.test(normalizeText(header))
  );
}

function normalizeAggregation(value = "") {
  const text = normalizeText(value).toLowerCase();
  if (
    ["average", "avg", "mean", "평균"].includes(text) ||
    /평균|비율|지수|점수|시간/.test(text)
  ) {
    return "average";
  }
  if (
    ["sum", "total", "합계", "합산"].includes(text) ||
    /합계|합산/.test(text)
  ) {
    return "sum";
  }
  return "";
}

function inferAggregation({ metricLabel = "", unit = "", column = {} } = {}) {
  const declared = normalizeAggregation(
    column.aggregation || column.metricKind || column.meta?.aggregation || "",
  );
  if (declared) return declared;

  const evidence = [metricLabel, unit, semanticType(column), columnRole(column)]
    .map(normalizeText)
    .join(" ");

  if (BUSINESS_AVERAGE_HEADER_PATTERN.test(evidence)) return "average";
  if (BUSINESS_SUM_HEADER_PATTERN.test(evidence)) return "sum";

  return /%|퍼센트|백분율|비율|비중|구성비|점유율|증감률|달성률|지수|평균|평점|점수|시간|기록|속도|초|분초|cm|명\/천명|rate|ratio|share|percent|index|average|avg|score|duration|time/i.test(
    evidence,
  )
    ? "average"
    : "sum";
}

function columnProfile(table = {}, column = {}, index = 0) {
  const values = tableRows(table)
    .map((row) => rowValue(row, column, index))
    .filter((value) => normalizeText(value));

  const numericCount = values.filter(
    (value) => numericValue(value) != null,
  ).length;

  return {
    nonBlankCount: values.length,
    numericCount,
    numericRatio: values.length ? numericCount / values.length : 0,
    distinctCount: new Set(values.map(normalizeText)).size,
  };
}

function indexedColumns(table = {}) {
  return tableColumns(table).map((column, index) => ({
    column,
    index,
    header: columnHeader(column, index),
    profile: columnProfile(table, column, index),
  }));
}

function canonicalLongContract(table = {}) {
  const columns = indexedColumns(table);
  const metricIdentity = columns.find((entry) =>
    isMetricIdentityColumn(entry.column, entry.header),
  );
  const metricValue = columns.find((entry) =>
    isMetricValueColumn(entry.column, entry.header),
  );

  if (!metricIdentity || !metricValue) return null;

  const explicitMetricIdentity =
    semanticType(metricIdentity.column) === "metricidentity" ||
    columnRole(metricIdentity.column) === "metricidentity";
  const canonicalMetricValueHeader = METRIC_VALUE_HEADER_PATTERN.test(
    metricValue.header,
  );

  if (!explicitMetricIdentity && !canonicalMetricValueHeader) {
    return null;
  }

  const unit = columns.find((entry) =>
    isUnitColumn(entry.column, entry.header),
  );
  const aggregation = columns.find((entry) =>
    isAggregationColumn(entry.column, entry.header),
  );
  const period = columns.find((entry) =>
    isPeriodColumn(entry.column, entry.header),
  );
  const year = columns.find((entry) =>
    isYearColumn(entry.column, entry.header),
  );

  const protectedIndexes = new Set(
    [metricIdentity, metricValue, unit, aggregation, period, year]
      .filter(Boolean)
      .map((entry) => entry.index),
  );

  const dimensions = columns.filter((entry) => {
    if (protectedIndexes.has(entry.index)) return false;
    if (isIdentifierColumn(entry.column, entry.header)) return false;
    if (EXCLUDED_DIMENSION_HEADER_PATTERN.test(entry.header)) return false;
    return entry.profile.nonBlankCount > 0;
  });

  return {
    type: "canonical_long",
    columns,
    metricIdentity,
    metricValue,
    unit,
    aggregation,
    period,
    year,
    dimensions,
    protectedHeaders: [
      metricIdentity.header,
      metricValue.header,
      unit?.header,
      aggregation?.header,
      period?.header,
      year?.header,
    ].filter(Boolean),
  };
}

function dimensionValuesForRow(row, dimensions = []) {
  const values = {};
  for (const entry of dimensions) {
    const value = normalizeText(rowValue(row, entry.column, entry.index));
    if (value) values[entry.header] = value;
  }
  return values;
}

function shouldSkipDimensionRow(dimensionValues = {}) {
  return Object.values(dimensionValues).some(isSummaryLabel);
}

function canonicalLongSeries(
  table = {},
  tableIndex = 0,
  contract = null,
  context = {},
) {
  const rows = tableRows(table);
  const distinctMetricLabels = new Set();

  for (const row of rows) {
    const label = normalizeText(
      rowValue(
        row,
        contract.metricIdentity.column,
        contract.metricIdentity.index,
      ),
    );
    if (label) distinctMetricLabels.add(label);
  }

  const metricLabels = [...distinctMetricLabels];
  const hasDetailMetric = metricLabels.some(
    (label) => !isSummaryMetricLabel(label, metricLabels),
  );
  const seriesMap = new Map();

  for (const row of rows) {
    let metricLabel = normalizeText(
      rowValue(
        row,
        contract.metricIdentity.column,
        contract.metricIdentity.index,
      ),
    );
    if (!metricLabel) continue;
    if (hasDetailMetric && isSummaryMetricLabel(metricLabel, metricLabels))
      continue;

    const value = numericValue(
      rowValue(row, contract.metricValue.column, contract.metricValue.index),
    );
    if (value == null) continue;

    const dimensions = dimensionValuesForRow(row, contract.dimensions);
    if (shouldSkipDimensionRow(dimensions)) continue;

    const unit = contract.unit
      ? normalizeText(rowValue(row, contract.unit.column, contract.unit.index))
      : normalizeText(
          contract.metricValue.column.unit ||
            contract.metricValue.column.measureUnit ||
            "",
        );

    metricLabel = inferContextualMetricLabel({
      metricLabel,
      unit,
      table,
      context,
    });

    const declaredAggregation = contract.aggregation
      ? normalizeAggregation(
          rowValue(
            row,
            contract.aggregation.column,
            contract.aggregation.index,
          ),
        )
      : "";

    const operation =
      declaredAggregation ||
      inferAggregation({
        metricLabel,
        unit,
        column: contract.metricValue.column,
      });

    const period = canonicalPeriodValue(
      contract.period
        ? rowValue(row, contract.period.column, contract.period.index)
        : "",
      contract.year
        ? rowValue(row, contract.year.column, contract.year.index)
        : "",
    );

    const key = [normalizeKey(metricLabel), normalizeKey(unit), operation].join(
      "::",
    );

    if (!seriesMap.has(key)) {
      seriesMap.set(key, {
        key,
        tableIndex,
        metricLabel,
        unit,
        operation,
        sourceContract: "canonical_long",
        valueHeader: contract.metricValue.header,
        protectedHeaders: contract.protectedHeaders,
        dimensionHeaders: contract.dimensions.map((entry) => entry.header),
        records: [],
      });
    }

    seriesMap.get(key).records.push({
      value,
      period,
      dimensions,
    });
  }

  return [...seriesMap.values()].filter((series) => series.records.length);
}

function physicalWideContract(table = {}) {
  const columns = indexedColumns(table);
  const period = columns.find((entry) =>
    isPeriodColumn(entry.column, entry.header),
  );
  const year = columns.find((entry) =>
    isYearColumn(entry.column, entry.header),
  );

  const dimensions = columns.filter((entry) => {
    if (isIdentifierColumn(entry.column, entry.header)) return false;
    if (isUnitColumn(entry.column, entry.header)) return false;
    if (isAggregationColumn(entry.column, entry.header)) return false;
    if (isPeriodColumn(entry.column, entry.header)) return false;
    if (isYearColumn(entry.column, entry.header)) return false;
    if (EXCLUDED_DIMENSION_HEADER_PATTERN.test(entry.header)) return false;
    return entry.profile.nonBlankCount > 0 && entry.profile.numericRatio < 0.5;
  });

  const dimensionIndexes = new Set(dimensions.map((entry) => entry.index));

  const measures = columns.filter((entry) => {
    if (dimensionIndexes.has(entry.index)) return false;
    if (isIdentifierColumn(entry.column, entry.header)) return false;
    if (isUnitColumn(entry.column, entry.header)) return false;
    if (isAggregationColumn(entry.column, entry.header)) return false;
    if (isPeriodColumn(entry.column, entry.header)) return false;
    if (isYearColumn(entry.column, entry.header)) return false;

    const declaredNumeric =
      normalizeText(entry.column.type).toLowerCase() === "number" ||
      columnRole(entry.column) === "metric" ||
      semanticType(entry.column) === "measure";

    return (
      entry.profile.numericCount > 0 &&
      (declaredNumeric || entry.profile.numericRatio >= 0.5)
    );
  });

  if (!measures.length) return null;

  return {
    type: "physical_wide",
    columns,
    dimensions,
    measures,
    period,
    year,
    protectedHeaders: columns
      .filter(
        (entry) =>
          isIdentifierColumn(entry.column, entry.header) ||
          isPeriodColumn(entry.column, entry.header) ||
          isYearColumn(entry.column, entry.header) ||
          isUnitColumn(entry.column, entry.header) ||
          isAggregationColumn(entry.column, entry.header),
      )
      .map((entry) => entry.header),
  };
}

function fallbackMetricLabel(table = {}, header = "") {
  const stripped = stripHeaderUnit(header);
  if (
    stripped &&
    !/^(?:기간|연도|년도|년월|월|분기|값|지표값)$/i.test(stripped)
  ) {
    return stripped;
  }

  const explicit = normalizeText(table.metricLabel || table.measureName || "");
  if (explicit) return explicit;

  return "지표값";
}

function physicalWideSeries(
  table = {},
  tableIndex = 0,
  contract = null,
  context = {},
) {
  const rows = tableRows(table);
  const measureLabels = contract.measures.map((entry) => {
    const temporal = parseTemporalMeasureHeader(entry.header);
    return (
      temporal.metricLabel ||
      (temporal.period ? "지표값" : fallbackMetricLabel(table, entry.header))
    );
  });
  const hasDetailMetric = measureLabels.some(
    (label) => !isSummaryMetricLabel(label, measureLabels),
  );
  const seriesMap = new Map();

  contract.measures.forEach((measure, measureIndex) => {
    const temporal = parseTemporalMeasureHeader(measure.header);
    let metricLabel =
      temporal.metricLabel ||
      (temporal.period ? "지표값" : fallbackMetricLabel(table, measure.header));

    if (!metricLabel) return;
    if (hasDetailMetric && isSummaryMetricLabel(metricLabel, measureLabels))
      return;

    const unit =
      normalizeText(
        measure.column.unit ||
          measure.column.measureUnit ||
          measure.column.meta?.unit ||
          "",
      ) || temporal.unit;

    metricLabel = inferContextualMetricLabel({
      metricLabel,
      unit,
      table,
      context,
    });

    const operation = inferAggregation({
      metricLabel,
      unit,
      column: measure.column,
    });

    const key = [normalizeKey(metricLabel), normalizeKey(unit), operation].join(
      "::",
    );

    if (!seriesMap.has(key)) {
      seriesMap.set(key, {
        key,
        tableIndex,
        metricLabel,
        unit,
        operation,
        sourceContract: "physical_wide",
        valueHeader: measure.header,
        protectedHeaders: contract.protectedHeaders,
        dimensionHeaders: contract.dimensions.map((entry) => entry.header),
        records: [],
      });
    }

    for (const row of rows) {
      const value = numericValue(rowValue(row, measure.column, measure.index));
      if (value == null) continue;

      const dimensions = dimensionValuesForRow(row, contract.dimensions);
      if (shouldSkipDimensionRow(dimensions)) continue;

      const rowPeriod =
        temporal.period ||
        strictPeriodValue(
          contract.period
            ? rowValue(row, contract.period.column, contract.period.index)
            : contract.year
              ? rowValue(row, contract.year.column, contract.year.index)
              : "",
        );

      seriesMap.get(key).records.push({
        value,
        period: rowPeriod,
        dimensions,
        measureIndex,
      });
    }
  });

  return [...seriesMap.values()].filter((series) => series.records.length);
}

function buildSemanticSeries(table = {}, tableIndex = 0, context = {}) {
  const canonical = canonicalLongContract(table);
  if (canonical) {
    return {
      contract: canonical,
      series: canonicalLongSeries(table, tableIndex, canonical, context),
    };
  }

  const physical = physicalWideContract(table);
  if (!physical) {
    return { contract: null, series: [] };
  }

  return {
    contract: physical,
    series: physicalWideSeries(table, tableIndex, physical, context),
  };
}

function canPlanSemanticOutput(tables = [], context = {}) {
  return (
    Array.isArray(tables) &&
    tables.some(
      (table, index) =>
        buildSemanticSeries(table, index, context).series.length > 0,
    )
  );
}

function numberStats(values = []) {
  const numbers = values.filter(
    (value) => typeof value === "number" && Number.isFinite(value),
  );
  if (!numbers.length) {
    return {
      count: 0,
      sum: 0,
      average: null,
      min: null,
      max: null,
    };
  }

  const sum = numbers.reduce((total, value) => total + value, 0);
  return {
    count: numbers.length,
    sum,
    average: sum / numbers.length,
    min: Math.min(...numbers),
    max: Math.max(...numbers),
  };
}

function safeId(value = "") {
  return (
    normalizeText(value)
      .toLowerCase()
      .replace(/[^0-9a-zA-Z가-힣]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 80) || "item"
  );
}

function metricId(tableIndex, metricLabel, unit = "", suffix = "") {
  return [
    "semantic",
    `table_${tableIndex + 1}`,
    safeId(metricLabel),
    safeId(unit || "value"),
    safeId(suffix),
  ]
    .filter(Boolean)
    .join(".");
}

function groupRows({
  records = [],
  groupHeader = "그룹",
  operation = "sum",
  metricLabel = "지표값",
  unit = "",
  groupValue,
} = {}) {
  const grouped = new Map();

  for (const record of records) {
    const key = normalizeText(groupValue(record));
    if (!key) continue;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(record.value);
  }

  return [...grouped.entries()]
    .map(([group, values]) => {
      const stats = numberStats(values);
      return {
        [groupHeader]: group,
        operation,
        metric: metricLabel,
        value: operation === "average" ? stats.average : stats.sum,
        rowCount: stats.count,
        단위: unit,
        평균: stats.average,
        최솟값: stats.min,
        최댓값: stats.max,
      };
    })
    .sort((left, right) =>
      String(left[groupHeader]).localeCompare(
        String(right[groupHeader]),
        "ko",
        { numeric: true },
      ),
    );
}

function sectionHash(section = {}) {
  return crypto
    .createHash("sha256")
    .update(
      JSON.stringify({
        type: section.sectionType,
        operation: section.result?.operation,
        metric: section.result?.metric,
        groupBy: section.result?.groupBy,
        rows: section.result?.rows,
      }),
    )
    .digest("hex");
}

function dedupeSections(sections = []) {
  const seen = new Set();
  const titles = new Map();
  const output = [];

  for (const section of sections) {
    const hash = sectionHash(section);
    if (seen.has(hash)) continue;
    seen.add(hash);

    const baseTitle =
      normalizeText(section.title || section.sectionId) || "분석";
    const count = (titles.get(baseTitle) || 0) + 1;
    titles.set(baseTitle, count);

    output.push({
      ...section,
      title: count === 1 ? baseTitle : `${baseTitle} ${count}`,
    });
  }

  return output;
}

function overviewSection(table = {}, tableIndex = 0, contract = null) {
  const label = tableLabel(table, tableIndex);
  const id = metricId(tableIndex, "overview", "", "overview");

  return {
    sectionId: `semantic_table_${tableIndex + 1}_overview`,
    title: `${label} 개요`,
    sectionType: "semantic_table_overview",
    metricIds: [id],
    result: {
      ok: true,
      resultType: "pivot",
      operation: "semanticTableOverview",
      rows: [
        { 지표: "테이블명", 값: label },
        { 지표: "행 수", 값: tableRows(table).length },
        { 지표: "열 수", 값: tableColumns(table).length },
        {
          지표: "Planner 계약",
          값: contract?.type || "unplanned",
        },
        {
          지표: "계산값 열",
          값:
            contract?.type === "canonical_long"
              ? contract.metricValue.header
              : "숫자 measure 열",
        },
      ],
      meta: {
        metricIds: [id],
        complete: true,
        plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
      },
    },
  };
}

function dimensionSemanticPriority({ header = "", distinctCount = 0 } = {}) {
  const text = normalizeText(header);
  let score = 0;

  if (
    /상태|등급|결과|구분|분류|유형|단계|채널|지역|부서|업종|직급|소속|담당|category|status|grade|result|type/i.test(
      text,
    )
  ) {
    score += 120;
  }

  if (
    /명$|이름|업체|거래처|기관|사업|프로젝트|과제|강사|강좌|행사|회의|자산|고객|서비스|품목|항목|entity|vendor|project|customer/i.test(
      text,
    )
  ) {
    score += 80;
  }

  if (
    BUSINESS_TIME_HEADER_PATTERN.test(text) ||
    /연도|년도|기준년도|일$|일자$|날짜$/i.test(text)
  ) {
    score -= 100;
  }

  if (distinctCount >= 2 && distinctCount <= 12) score += 35;
  else if (distinctCount <= 50) score += 20;
  else if (distinctCount > 100) score -= 20;

  return score;
}

function seriesSections(table = {}, tableIndex = 0, series = {}, options = {}) {
  const label = tableLabel(table, tableIndex);
  const displayMetric = series.unit
    ? `${series.metricLabel} (${series.unit})`
    : series.metricLabel;
  const titlePrefix = options.compactTitles ? "" : `${label} · `;
  const maxDimensionsPerSeries = Math.max(
    0,
    Number(options.maxDimensionsPerSeries ?? 3),
  );
  const baseId = metricId(
    tableIndex,
    series.metricLabel,
    series.unit,
    series.operation,
  );
  const stats = numberStats(series.records.map((record) => record.value));
  const additive = series.operation === "sum";
  const sections = [
    {
      sectionId: `${baseId}.summary`,
      title: `${titlePrefix}${displayMetric} 통계`,
      sectionType: additive
        ? "semantic_additive_summary"
        : "semantic_non_additive_summary",
      metricIds: [`${baseId}.summary`],
      result: {
        ok: true,
        resultType: "pivot",
        operation: additive
          ? "semanticAggregateSummary"
          : "semanticAverageRange",
        metric: {
          header: series.metricLabel,
          unit: series.unit,
        },
        rows: additive
          ? [
              { 지표: "유효값 수", 값: stats.count, 단위: "건" },
              { 지표: "합계", 값: stats.sum, 단위: series.unit },
              { 지표: "평균", 값: stats.average, 단위: series.unit },
              { 지표: "최솟값", 값: stats.min, 단위: series.unit },
              { 지표: "최댓값", 값: stats.max, 단위: series.unit },
            ]
          : [
              { 지표: "유효값 수", 값: stats.count, 단위: "건" },
              { 지표: "평균", 값: stats.average, 단위: series.unit },
              { 지표: "최솟값", 값: stats.min, 단위: series.unit },
              { 지표: "최댓값", 값: stats.max, 단위: series.unit },
            ],
        meta: {
          metricIds: [`${baseId}.summary`],
          complete: true,
          additive,
          unit: series.unit,
          sourceContract: series.sourceContract,
          valueHeader: series.valueHeader,
          protectedHeaders: series.protectedHeaders,
          plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
        },
      },
    },
  ];

  const dimensionHeaders = series.dimensionHeaders
    .map((header) => {
      const distinctCount = new Set(
        series.records
          .map((record) => normalizeText(record.dimensions?.[header]))
          .filter(Boolean),
      ).size;

      return {
        header,
        distinctCount,
        semanticPriority: dimensionSemanticPriority({
          header,
          distinctCount,
        }),
      };
    })
    .filter((entry) => entry.distinctCount >= 2 && entry.distinctCount <= 200)
    .sort(
      (left, right) =>
        right.semanticPriority - left.semanticPriority ||
        right.distinctCount - left.distinctCount ||
        left.header.localeCompare(right.header, "ko"),
    )
    .slice(0, maxDimensionsPerSeries);

  for (const dimension of dimensionHeaders) {
    const rows = groupRows({
      records: series.records,
      groupHeader: dimension.header,
      operation: series.operation,
      metricLabel: series.metricLabel,
      unit: series.unit,
      groupValue: (record) => record.dimensions?.[dimension.header],
    });
    if (!rows.length) continue;

    sections.push({
      sectionId: `${baseId}.by_${safeId(dimension.header)}`,
      title: `${titlePrefix}${dimension.header}별 ${displayMetric}`,
      sectionType: additive ? "semantic_group_sum" : "semantic_group_average",
      metricIds: [`${baseId}.by_${safeId(dimension.header)}`],
      result: {
        ok: true,
        resultType: "grouped",
        operation: series.operation,
        groupBy: { header: dimension.header },
        metric: {
          header: series.metricLabel,
          unit: series.unit,
        },
        rowCount: series.records.length,
        rows,
        meta: {
          complete: true,
          additive,
          unit: series.unit,
          sourceContract: series.sourceContract,
          valueHeader: series.valueHeader,
          plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
        },
      },
    });
  }

  const periodRows = groupRows({
    records: series.records.filter((record) => record.period),
    groupHeader: "기간",
    operation: series.operation,
    metricLabel: series.metricLabel,
    unit: series.unit,
    groupValue: (record) => record.period,
  });

  if (periodRows.length) {
    sections.push({
      sectionId: `${baseId}.by_period`,
      title: `${titlePrefix}기간별 ${displayMetric}`,
      sectionType: additive ? "semantic_period_sum" : "semantic_period_average",
      metricIds: [`${baseId}.by_period`],
      result: {
        ok: true,
        resultType: "grouped",
        operation: series.operation,
        groupBy: { header: "기간" },
        metric: {
          header: series.metricLabel,
          unit: series.unit,
        },
        rowCount: series.records.filter((record) => record.period).length,
        rows: periodRows,
        meta: {
          complete: true,
          additive,
          explicitPeriodOnly: true,
          unit: series.unit,
          sourceContract: series.sourceContract,
          valueHeader: series.valueHeader,
          protectedHeaders: series.protectedHeaders,
          plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
        },
      },
    });
  }

  return sections;
}

function semanticScalarText(value, output = [], depth = 0) {
  if (depth > 5 || value == null) return output;
  if (["string", "number", "boolean"].includes(typeof value)) {
    const text = normalizeText(value);
    if (text) output.push(text);
    return output;
  }
  if (Array.isArray(value)) {
    value
      .slice(0, 100)
      .forEach((item) => semanticScalarText(item, output, depth + 1));
    return output;
  }
  if (typeof value === "object") {
    Object.entries(value)
      .slice(0, 100)
      .forEach(([key, item]) => {
        output.push(normalizeText(key));
        semanticScalarText(item, output, depth + 1);
      });
  }
  return output;
}

function semanticSectionText(section = {}) {
  return semanticScalarText({
    sectionId: section.sectionId,
    title: section.title,
    sectionType: section.sectionType,
    candidate: section.candidate,
    operation: section.result?.operation,
    metric: section.result?.metric,
    metrics: section.result?.metrics,
    groupBy: section.result?.groupBy,
    rowHeaders: Array.isArray(section.result?.rows)
      ? Array.from(
          new Set(
            section.result.rows
              .slice(0, 30)
              .flatMap((row) =>
                row && typeof row === "object" ? Object.keys(row) : [],
              ),
          ),
        )
      : [],
  })
    .map(normalizeKey)
    .filter(Boolean)
    .join(" ");
}

function semanticOperationFamily(section = {}) {
  const text = normalizeText(
    `${section.sectionType || ""} ${section.result?.operation || ""}`,
  ).toLowerCase();
  if (/composition|ratio|비율|구성비/.test(text)) return "ratio";
  if (/top|bottom|rank|상위|하위|순위/.test(text)) return "rank";
  if (/count|건수|응답\s*수/.test(text)) return "count";
  if (/average|mean|평균/.test(text)) return "average";
  if (/sum|aggregate|합계|합산/.test(text)) return "sum";
  if (/summary|overview|통계|요약/.test(text)) return "summary";
  return "other";
}

function semanticGroupAliases(header = "") {
  const normalized = normalizeText(header);
  if (!normalized) return [];
  if (normalized === "기간" || BUSINESS_TIME_HEADER_PATTERN.test(normalized)) {
    return [
      "기간",
      "월",
      "연월",
      "년월",
      "일자",
      "날짜",
      "기준월",
      "응답일",
      "평가일",
      "거래일",
      "취득일",
    ];
  }
  return [normalized];
}

function sectionContainsToken(sectionText = "", token = "") {
  const normalized = normalizeKey(token);
  return Boolean(normalized && sectionText.includes(normalized));
}

function explicitSectionMetricHeader(section = {}) {
  return normalizeText(
    section.result?.metric?.header ||
      section.metric?.header ||
      section.metricHeader ||
      section.candidate?.metricHeader ||
      section.candidate?.metric ||
      "",
  );
}

function explicitSectionGroupHeader(section = {}) {
  return normalizeText(
    section.result?.groupBy?.header ||
      section.groupBy?.header ||
      section.groupHeader ||
      section.candidate?.groupHeader ||
      section.candidate?.groupBy ||
      "",
  );
}

function metricHeadersSemanticallyMatch(
  existingMetric = "",
  plannedMetric = "",
) {
  const existing = normalizeText(existingMetric);
  const planned = normalizeText(plannedMetric);

  if (!existing || !planned) return true;

  const existingGeneric = GENERIC_METRIC_LABEL_PATTERN.test(existing);
  const plannedGeneric = GENERIC_METRIC_LABEL_PATTERN.test(planned);

  if (existingGeneric !== plannedGeneric) return false;

  return normalizeKey(existing) === normalizeKey(planned);
}

function groupHeadersSemanticallyMatch(existingGroup = "", plannedGroup = "") {
  const existing = normalizeText(existingGroup);
  const planned = normalizeText(plannedGroup);

  if (!existing || !planned) return true;

  const aliases = new Set(semanticGroupAliases(planned).map(normalizeKey));
  return aliases.has(normalizeKey(existing));
}

function existingSectionCoversPlanned(existing = {}, planned = {}) {
  const metric = normalizeText(planned.result?.metric?.header || "");
  const group = normalizeText(planned.result?.groupBy?.header || "");
  const text = semanticSectionText(existing);
  const existingMetric = explicitSectionMetricHeader(existing);
  const existingGroup = explicitSectionGroupHeader(existing);

  if (!metric || !sectionContainsToken(text, metric)) return false;

  if (!metricHeadersSemanticallyMatch(existingMetric, metric)) {
    return false;
  }

  if (group && !groupHeadersSemanticallyMatch(existingGroup, group)) {
    return false;
  }

  if (group) {
    const groupCovered = semanticGroupAliases(group).some((alias) =>
      sectionContainsToken(text, alias),
    );
    if (!groupCovered) return false;
  }

  const plannedFamily = semanticOperationFamily(planned);
  const existingFamily = semanticOperationFamily(existing);

  if (["ratio", "rank", "count"].includes(existingFamily)) return false;
  if (plannedFamily === "average") {
    return (
      ["average", "summary"].includes(existingFamily) ||
      /평균|점수\s*요약/.test(normalizeText(existing.title))
    );
  }
  if (plannedFamily === "sum") {
    return (
      ["sum", "summary"].includes(existingFamily) ||
      /합계|금액\s*요약|수량\s*요약/.test(normalizeText(existing.title))
    );
  }
  if (!group) {
    return (
      ["summary", "average", "sum"].includes(existingFamily) ||
      /요약|통계/.test(normalizeText(existing.title))
    );
  }
  return true;
}

function annotateExistingSection(existing = {}, planned = {}) {
  const metricIds = uniqueMetricIds([
    collectSectionMetricIds(existing),
    collectSectionMetricIds(planned),
  ]);
  const annotated = applySectionMetricIds(existing, metricIds);
  annotated.result = {
    ...(annotated.result || {}),
    meta: {
      ...(annotated.result?.meta || {}),
      semanticCoverage: {
        plannerSectionId: planned.sectionId,
        plannerSectionType: planned.sectionType,
        matchedExistingSection: true,
      },
    },
  };
  return annotated;
}

const GENERIC_SECTION_CLEANUP_VERSION = "semantic_generic_section_cleanup_v1";

function plannedSpecificMetricHeaders(plannedSections = []) {
  return Array.from(
    new Set(
      plannedSections
        .map((section) => normalizeText(section.result?.metric?.header || ""))
        .filter(
          (header) => header && !GENERIC_METRIC_LABEL_PATTERN.test(header),
        ),
    ),
  );
}

function sectionHasGenericMetricIdentity(section = {}) {
  const explicitMetric = explicitSectionMetricHeader(section);

  if (explicitMetric) {
    return GENERIC_METRIC_LABEL_PATTERN.test(explicitMetric);
  }

  const title = normalizeText(section.title);
  return /지표값|metric\s*value|measure\s*value/i.test(title);
}

function plannedMetricHasReplacementCoverage({
  plannedSections = [],
  metricHeader = "",
} = {}) {
  const metricKey = normalizeKey(metricHeader);
  const relevant = plannedSections.filter(
    (section) =>
      normalizeKey(section.result?.metric?.header || "") === metricKey,
  );

  if (!relevant.length) return false;

  const hasSummary = relevant.some((section) => {
    const group = normalizeText(section.result?.groupBy?.header || "");
    const family = semanticOperationFamily(section);
    return !group && ["summary", "sum", "average"].includes(family);
  });

  const hasGroupedOrPeriod = relevant.some((section) =>
    normalizeText(section.result?.groupBy?.header || ""),
  );

  return hasSummary && hasGroupedOrPeriod;
}

function pruneSupersededGenericSections({
  sections = [],
  plannedSections = [],
  baseSectionCount = 0,
} = {}) {
  const specificMetrics = plannedSpecificMetricHeaders(plannedSections);

  if (specificMetrics.length !== 1) {
    return {
      sections,
      prunedSectionCount: 0,
      prunedSectionIds: [],
      cleanupApplied: false,
      reason: "SPECIFIC_METRIC_COUNT_NOT_ONE",
    };
  }

  const [specificMetric] = specificMetrics;

  if (
    !plannedMetricHasReplacementCoverage({
      plannedSections,
      metricHeader: specificMetric,
    })
  ) {
    return {
      sections,
      prunedSectionCount: 0,
      prunedSectionIds: [],
      cleanupApplied: false,
      reason: "SPECIFIC_METRIC_COVERAGE_INCOMPLETE",
    };
  }

  const prunedSectionIds = [];

  const retained = sections.filter((section, index) => {
    if (index >= baseSectionCount) return true;

    if (collectSectionMetricIds(section).length > 0) {
      return true;
    }

    if (!sectionHasGenericMetricIdentity(section)) {
      return true;
    }

    prunedSectionIds.push(
      normalizeText(
        section.sectionId || section.title || `base_section_${index + 1}`,
      ),
    );
    return false;
  });

  return {
    sections: retained,
    prunedSectionCount: prunedSectionIds.length,
    prunedSectionIds,
    cleanupApplied: prunedSectionIds.length > 0,
    reason:
      prunedSectionIds.length > 0 ? "" : "NO_UNCONTRACTED_GENERIC_SECTION",
    specificMetric,
  };
}

function augmentBusinessTemplateResult({
  executionResult = {},
  tables = [],
  templateCandidate = {},
  context = {},
  options = {},
} = {}) {
  const inputSnapshot = JSON.stringify(tables);
  const baseSections = Array.isArray(executionResult.sections)
    ? cloneValue(executionResult.sections)
    : [];
  const preferredTables = selectPreferredSemanticTables(tables, {
    preferDimensionCompletePhysical: true,
  });
  const plannerContext = {
    ...cloneValue(context),
    templateId:
      executionResult.templateId || templateCandidate.templateId || "",
    templateTitle: executionResult.title || templateCandidate.title || "",
    templateDescription:
      executionResult.description || templateCandidate.description || "",
  };

  const plan = buildSemanticOutputPlan({
    tables: preferredTables,
    templateId:
      executionResult.templateId ||
      templateCandidate.templateId ||
      "business_template",
    title: executionResult.title || templateCandidate.title || "업무 템플릿",
    context: plannerContext,
    options: {
      includeOverview: false,
      includeDiagnostics: false,
      compactTitles: true,
      maxDimensionsPerSeries: options.maxDimensionsPerSeries ?? 8,
      maxSections: options.maxPlannedSections ?? 120,
      preferVirtualTables: false,
    },
  });

  const plannedSections = (plan.sections || []).filter(
    (section) =>
      !String(section.sectionType || "").startsWith(
        "semantic_table_overview",
      ) && !String(section.sectionType || "").startsWith("semanticSource"),
  );

  const merged = [...baseSections];
  const renderedPlannerIds = [];
  let matchedExistingSectionCount = 0;
  let addedSectionCount = 0;
  const maxAddedSections = Math.max(0, Number(options.maxAddedSections ?? 64));

  for (const planned of plannedSections) {
    const existingIndex = merged.findIndex((section) =>
      existingSectionCoversPlanned(section, planned),
    );

    if (existingIndex >= 0) {
      merged[existingIndex] = annotateExistingSection(
        merged[existingIndex],
        planned,
      );
      renderedPlannerIds.push(...collectSectionMetricIds(planned));
      matchedExistingSectionCount += 1;
      continue;
    }

    if (addedSectionCount >= maxAddedSections) continue;

    const added = applySectionMetricIds(planned);
    added.result = {
      ...(added.result || {}),
      meta: {
        ...(added.result?.meta || {}),
        semanticCoverage: {
          plannerSectionId: planned.sectionId,
          plannerSectionType: planned.sectionType,
          matchedExistingSection: false,
          addedByBusinessAugmentation: true,
        },
      },
    };
    merged.push(added);
    renderedPlannerIds.push(...collectSectionMetricIds(planned));
    addedSectionCount += 1;
  }

  if (JSON.stringify(tables) !== inputSnapshot) {
    throw new Error(
      "업무 템플릿 Semantic Output Planner가 입력 테이블을 변경했습니다.",
    );
  }

  const genericCleanup = pruneSupersededGenericSections({
    sections: merged,
    plannedSections,
    baseSectionCount: baseSections.length,
  });

  const normalizedSections = normalizeSectionMetricIds(genericCleanup.sections);
  const expectedMetricIds = uniqueMetricIds(renderedPlannerIds);

  return {
    ...cloneValue(executionResult),
    sections: normalizedSections,
    contractSummaryCoverage: {
      version: `${SEMANTIC_OUTPUT_CONTRACT_VERSION}_business_coverage_v1`,
      contractCatalogVersion: `${SEMANTIC_OUTPUT_CONTRACT_VERSION}_business_coverage_catalog_v1`,
      expectedMetricIds,
    },
    executionMeta: {
      ...(executionResult.executionMeta || {}),
      semanticBusinessAugmentation: true,
      semanticOutputPlannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
      preferredTableCount: preferredTables.length,
      plannedSectionCount: plannedSections.length,
      matchedExistingSectionCount,
      addedSectionCount,
      expectedMetricIdCount: expectedMetricIds.length,
      maxAddedSections,
      genericSectionCleanupVersion: GENERIC_SECTION_CLEANUP_VERSION,
      genericSectionCleanupApplied: genericCleanup.cleanupApplied,
      prunedGenericSectionCount: genericCleanup.prunedSectionCount,
      prunedGenericSectionIds: genericCleanup.prunedSectionIds,
      genericSectionCleanupReason: genericCleanup.reason,
      genericSectionCleanupMetric: genericCleanup.specificMetric || "",
      sourceTablesPreserved: true,
    },
  };
}

function diagnosticSections(tables = [], plans = []) {
  const summaryRows = [];
  const diagnostics = [];
  const previews = [];

  tables.forEach((table, tableIndex) => {
    const label = tableLabel(table, tableIndex);
    const summary = Array.isArray(table.summaryRows) ? table.summaryRows : [];
    const excluded = Array.isArray(table.excludedRows)
      ? table.excludedRows
      : [];
    const plan = plans[tableIndex];

    summary.forEach((row, index) => {
      summaryRows.push({
        테이블: label,
        순번: index + 1,
        내용: JSON.stringify(row),
      });
    });

    diagnostics.push({
      테이블: label,
      행수: tableRows(table).length,
      열수: tableColumns(table).length,
      계약: plan?.contract?.type || "unplanned",
      series수: plan?.series?.length || 0,
      요약행수: summary.length,
      제외행수: excluded.length,
      계산값열:
        plan?.contract?.type === "canonical_long"
          ? plan.contract.metricValue.header
          : "숫자 measure 열",
    });

    tableRows(table)
      .slice(0, 10)
      .forEach((row, index) => {
        previews.push({
          테이블: label,
          행번호: index + 1,
          내용: JSON.stringify(row),
        });
      });
  });

  const make = (id, title, operation, rows) => ({
    sectionId: id,
    title,
    sectionType: operation,
    metricIds: [id],
    result: {
      ok: true,
      resultType: "pivot",
      operation,
      rows,
      meta: {
        metricIds: [id],
        complete: true,
        plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
      },
    },
  });

  return [
    make(
      "semantic.diagnostics.summary_rows",
      "요약행",
      "semanticSourceSummaryRows",
      summaryRows.length
        ? summaryRows
        : [{ 테이블: "-", 순번: 0, 내용: "분리된 요약행 없음" }],
    ),
    make(
      "semantic.diagnostics.tables",
      "진단정보",
      "semanticSourceDiagnostics",
      diagnostics,
    ),
    make(
      "semantic.diagnostics.preview",
      "실행결과_미리보기",
      "semanticExecutionPreview",
      previews,
    ),
  ];
}

function buildSemanticOutputPlan({
  tables = [],
  templateId = "generic_structured_summary",
  title = "범용 구조화 통계 요약",
  patchVersion = SEMANTIC_OUTPUT_PLANNER_VERSION,
  context = {},
  options = {},
} = {}) {
  const inputTables = Array.isArray(tables) ? tables : [];
  const inputSnapshot = JSON.stringify(inputTables);
  const sourceTables = options.preferVirtualTables
    ? selectPreferredSemanticTables(inputTables)
    : inputTables;
  const plans = sourceTables.map((table, index) =>
    buildSemanticSeries(table, index, context),
  );
  const sections = [];

  sourceTables.forEach((table, tableIndex) => {
    const plan = plans[tableIndex];
    if (options.includeOverview !== false) {
      sections.push(overviewSection(table, tableIndex, plan.contract));
    }
    plan.series.forEach((series) => {
      sections.push(...seriesSections(table, tableIndex, series, options));
    });
  });

  if (options.includeDiagnostics !== false) {
    sections.push(...diagnosticSections(sourceTables, plans));
  }

  const maxSections = Number(options.maxSections);
  const limitedSections =
    Number.isFinite(maxSections) && maxSections >= 0
      ? sections.slice(0, maxSections)
      : sections;

  if (JSON.stringify(inputTables) !== inputSnapshot) {
    throw new Error(
      "Semantic Output Planner가 입력 query table을 변경했습니다.",
    );
  }

  const dedupedSections = normalizeSectionMetricIds(
    dedupeSections(limitedSections),
  );
  const expectedMetricIds = Array.from(
    new Set(
      dedupedSections.flatMap((section) =>
        Array.isArray(section.metricIds) ? section.metricIds : [],
      ),
    ),
  );

  const canonicalLongTableCount = plans.filter(
    (plan) => plan.contract?.type === "canonical_long",
  ).length;
  const physicalWideTableCount = plans.filter(
    (plan) => plan.contract?.type === "physical_wide",
  ).length;

  return {
    ok: true,
    resultType: "businessTemplate",
    operation: "genericStructuredSummary",
    templateId,
    title,
    sections: dedupedSections,
    contractSummaryCoverage: {
      version: SEMANTIC_OUTPUT_CONTRACT_VERSION,
      contractCatalogVersion: `${SEMANTIC_OUTPUT_CONTRACT_VERSION}_catalog`,
      expectedMetricIds,
    },
    executionMeta: {
      sourceDataOnly: false,
      genericSummary: true,
      patchVersion,
      semanticOutputPlannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
      valueColumnOnlyForCanonicalLong: true,
      protectedTemporalColumns: true,
      metricIdentitySeparated: true,
      unitSeparated: true,
      aggregationRespected: true,
      summaryMetricSuppression: true,
      duplicateSuppression: true,
      sourceTablesPreserved: true,
      metricIdsPropagated: true,
      metricIdContractVersion: METRIC_ID_CONTRACT_VERSION,
      canonicalLongTableCount,
      physicalWideTableCount,
      plannedTableCount: plans.filter((plan) => plan.series.length > 0).length,
      context: cloneValue(context),
      compactTitles: options.compactTitles === true,
      maxDimensionsPerSeries: Number(options.maxDimensionsPerSeries ?? 3),
      preferVirtualTables: options.preferVirtualTables === true,
    },
  };
}

module.exports = {
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  SEMANTIC_OUTPUT_CONTRACT_VERSION,
  buildSemanticOutputPlan,
  buildSemanticSeries,
  augmentBusinessTemplateResult,
  canPlanSemanticOutput,
  selectPreferredSemanticTables,
  existingSectionCoversPlanned,
  canonicalLongContract,
  physicalWideContract,
  dimensionSemanticPriority,
  pruneSupersededGenericSections,
  strictPeriodValue,
};
