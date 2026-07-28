"use strict";

const crypto = require("crypto");
const {
  METRIC_ID_CONTRACT_VERSION,
  applySectionMetricIds,
  collectSectionMetricIds,
  normalizeSectionMetricIds,
  uniqueMetricIds,
} = require("./metricIdContract");
const {
  AGGREGATION_CONTRACT_RESOLVER_VERSION,
  METRIC_SEMANTIC_ROLE_ENGINE_VERSION,
  OPERATION: SEMANTIC_AGGREGATION_OPERATION,
  ROLE: SEMANTIC_METRIC_ROLE,
  classifyMetricRole,
  resolveAggregationContract,
} = require("./metricSemanticRoleEngine");
const {
  FLOW_DIRECTION_SECTION_REPAIR_VERSION,
  FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
  applyFlowDirectionSemantics,
  buildDirectionRows,
  buildEntityFlowRows,
  buildLocationLedgerRows,
  buildPeriodFlowRows,
  buildSystemFlowSummary,
  canonicalFlowDirection,
  resolveFlowDirectionEvidence,
} = require("./flowDirectionSemanticEngine");
const {
  DERIVED_TOTAL_RELATION_VERSION,
  METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
  REPRESENTATIVE_METRIC_PRIORITY_VERSION,
  applyMetricRelationshipPriorities,
  prioritizeBusinessSections,
} = require("./metricRelationshipPriorityEngine");
const {
  DISTINCT_ENTITY_SECTION_VERSION,
  DURATION_SUMMARY_CONTRACT_VERSION,
  MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
  SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
  applySemanticSectionBudget,
  buildDistinctEntitySection,
  median,
  sectionPolicyForSeries,
} = require("./semanticSectionBudgetEngine");
const {
  FINAL_OUTPUT_QUALITY_GATE_VERSION,
  OUTPUT_COMPLETENESS_CONTRACT_VERSION,
  SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
  applyFinalOutputQualityGate,
} = require("./finalOutputQualityGateEngine");

const SEMANTIC_OUTPUT_PLANNER_LEGACY_VERSION =
  "semantic_output_planner_common_v2_9_snapshot_latest_by_entity";
const SEMANTIC_OUTPUT_PLANNER_FLOW_DIRECTION_VERSION =
  "semantic_output_planner_common_v2_10_flow_direction_semantics";
const SEMANTIC_OUTPUT_PLANNER_PREVIOUS_VERSION =
  "semantic_output_planner_common_v2_12_mandatory_summary_coverage_floor";
const SEMANTIC_OUTPUT_PLANNER_VERSION =
  "semantic_output_planner_common_v2_13_final_output_quality_gate";
const SEMANTIC_OUTPUT_CONTRACT_VERSION =
  "semantic_output_contract_v1";
const SEMANTIC_CONTRACT_PRECEDENCE_VERSION =
  "semantic_contract_precedence_v1";
const MIXED_SECTION_ROW_PRECEDENCE_VERSION =
  "semantic_mixed_section_row_precedence_v2_contract_exclusion";
const GENERAL_STOCK_SNAPSHOT_ALIAS_VERSION =
  "general_stock_snapshot_alias_v2_pre_finalization_row_shape";
const CONTRACT_KPI_SNAPSHOT_BRIDGE_VERSION =
  "contract_kpi_snapshot_bridge_v2_aggregation_aware";
const ACTUAL_FLOW_EVIDENCE_GATE_VERSION =
  "actual_flow_evidence_restore_gate_v1";
const SNAPSHOT_ENTITY_RESOLVER_VERSION =
  "snapshot_latest_by_entity_resolver_v1";

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
const UNIT_HEADER_PATTERN =
  /^(?:단위|측정단위|unit|measure\s*unit)$/i;
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
const SNAPSHOT_ENTITY_IDENTITY_HEADER_PATTERN =
  /(?:^|[_\s])(?:id|code)(?:$|[_\s])|(?:품목|소모품|제품|상품|자산|장비|시설|고객|거래처|업체|기관|사업|프로젝트|과제|서비스|항목)(?:명|이름)$|^(?:이름|명칭|entity|item|product|asset|equipment|facility|customer|vendor|project)$/i;
const SNAPSHOT_ENTITY_SECONDARY_HEADER_PATTERN =
  /창고명|보관위치|지점명|매장명|사업장명|부서명|담당부서|사용부서|warehouse|location|branch|department/i;
const SNAPSHOT_ENTITY_EXCLUDED_HEADER_PATTERN =
  /상태|구분|분류|유형|결과|등급|채널|지역|기간|연월|월|일자|날짜|연도|년도|단위|비고|설명|status|category|type|result|grade|channel|region|period|date|month|year|unit|note/i;

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
      table.tableName ||
        table.sheetName ||
        table.title ||
        table.tableId ||
        "",
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

function selectPreferredSemanticTables(
  tables = [],
  options = {},
) {
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
        representedBy.flatMap((item) =>
          [...semanticDimensionHeaderSet(item)],
        ),
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
      return (
        !sourceId ||
        !physicalPreferredSources.has(sourceId)
      );
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

  const normalized = source
    .replace(/,/g, "")
    .replace(/%$/g, "")
    .trim();

  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) {
    return null;
  }

  const result = Number(normalized);
  return Number.isFinite(result) ? result : null;
}


const ACTUAL_FLOW_INBOUND_HEADER_PATTERN =
  /^(?:총|누적|기간)?(?:입고|입하|반입|수취|inbound|receipt|received)(?:수량|량|quantity|qty|count)?$/i;
const ACTUAL_FLOW_OUTBOUND_HEADER_PATTERN =
  /^(?:총|누적|기간)?(?:출고|출하|반출|outbound|shipment|shipped)(?:수량|량|quantity|qty|count)?$/i;
const ACTUAL_FLOW_DIRECTION_HEADER_PATTERN =
  /^(?:(?:입출고|이동|수불|재고이동|창고이동|거래|물류)(?:구분|유형|방향)|(?:flow|movement|transaction)(?:type|direction)|direction)$/i;
const ACTUAL_FLOW_QUANTITY_HEADER_PATTERN =
  /(?:수량|quantity|qty)$/i;

function columnNumericObservationCount(table = {}, column = {}, index = 0) {
  return tableRows(table).reduce((count, row) => {
    return numericValue(rowValue(row, column, index)) == null
      ? count
      : count + 1;
  }, 0);
}

function canonicalActualFlowDirection(value = "") {
  const key = normalizeKey(value);
  if (!key) return "";
  if (/^(?:입고|입하|반입|수취|inbound|receipt|received)$/.test(key)) {
    return "inbound";
  }
  if (/^(?:출고|출하|반출|outbound|shipment|shipped)$/.test(key)) {
    return "outbound";
  }
  if (/^(?:이동|재고이동|창고이동|transfer|movement)$/.test(key)) {
    return "transfer";
  }
  return "";
}

function tableActualFlowEvidence(table = {}, tableIndex = 0) {
  const columns = tableColumns(table);
  const rows = tableRows(table);
  const columnFacts = columns.map((column, index) => {
    const header = columnHeader(column, index);
    const key = normalizeKey(header);
    return {
      column,
      index,
      header,
      key,
      numericObservationCount: columnNumericObservationCount(
        table,
        column,
        index,
      ),
    };
  });

  const inbound = columnFacts.find(
    (fact) =>
      ACTUAL_FLOW_INBOUND_HEADER_PATTERN.test(fact.key) &&
      fact.numericObservationCount > 0,
  );
  const outbound = columnFacts.find(
    (fact) =>
      ACTUAL_FLOW_OUTBOUND_HEADER_PATTERN.test(fact.key) &&
      fact.numericObservationCount > 0,
  );

  if (inbound && outbound) {
    return {
      pass: true,
      mode: "explicit_inbound_outbound_columns",
      tableIndex,
      tableLabel: tableLabel(table, tableIndex),
      inboundHeader: inbound.header,
      outboundHeader: outbound.header,
      directionHeader: "",
      quantityHeader: "",
      directionClasses: [],
      recognizedDirectionRowCount: 0,
      rowCount: rows.length,
    };
  }

  const directionCandidates = columnFacts.filter((fact) =>
    ACTUAL_FLOW_DIRECTION_HEADER_PATTERN.test(fact.key),
  );
  const quantityCandidates = columnFacts.filter(
    (fact) =>
      ACTUAL_FLOW_QUANTITY_HEADER_PATTERN.test(fact.key) &&
      !ACTUAL_FLOW_INBOUND_HEADER_PATTERN.test(fact.key) &&
      !ACTUAL_FLOW_OUTBOUND_HEADER_PATTERN.test(fact.key) &&
      fact.numericObservationCount > 0,
  );

  for (const direction of directionCandidates) {
    const classes = new Set();
    let recognizedDirectionRowCount = 0;
    for (const row of rows) {
      const canonical = canonicalActualFlowDirection(
        rowValue(row, direction.column, direction.index),
      );
      if (!canonical) continue;
      recognizedDirectionRowCount += 1;
      classes.add(canonical);
    }

    const hasExternalDirection =
      classes.has("inbound") || classes.has("outbound");
    if (
      hasExternalDirection &&
      recognizedDirectionRowCount > 0 &&
      quantityCandidates.length > 0
    ) {
      const quantity = quantityCandidates[0];
      return {
        pass: true,
        mode: "direction_column_with_quantity",
        tableIndex,
        tableLabel: tableLabel(table, tableIndex),
        inboundHeader: "",
        outboundHeader: "",
        directionHeader: direction.header,
        quantityHeader: quantity.header,
        directionClasses: Array.from(classes).sort(),
        recognizedDirectionRowCount,
        rowCount: rows.length,
      };
    }
  }

  return {
    pass: false,
    mode: "no_actual_flow_evidence",
    tableIndex,
    tableLabel: tableLabel(table, tableIndex),
    inboundHeader: inbound?.header || "",
    outboundHeader: outbound?.header || "",
    directionHeader: directionCandidates[0]?.header || "",
    quantityHeader: quantityCandidates[0]?.header || "",
    directionClasses: [],
    recognizedDirectionRowCount: 0,
    rowCount: rows.length,
  };
}

function detectActualFlowEvidence(tables = []) {
  const input = Array.isArray(tables) ? tables : [];
  const tableEvidence = input
    .filter(isSemanticAnalysisEligible)
    .map((table, index) => tableActualFlowEvidence(table, index));
  const matched = tableEvidence.find((item) => item.pass === true);

  return {
    version: ACTUAL_FLOW_EVIDENCE_GATE_VERSION,
    evaluated: true,
    pass: Boolean(matched),
    mode: matched?.mode || "no_actual_flow_evidence",
    matchedTableIndex: matched?.tableIndex ?? -1,
    matchedTableLabel: matched?.tableLabel || "",
    tableEvidence,
  };
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
  return matches.length
    ? normalizeText(matches[matches.length - 1][1])
    : "";
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

  let match = source.match(
    /^((?:19|20|21)\d{2})[-./]\s*(0?[1-9]|1[0-2])$/,
  );
  if (match) {
    return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
  }

  match = source.match(
    /^((?:19|20|21)\d{2})\s*년\s*(0?[1-9]|1[0-2])\s*월$/,
  );
  if (match) {
    return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
  }

  match = source.match(
    /^((?:19|20|21)\d{2})\s*(?:Q([1-4])|([1-4])\s*분기)$/i,
  );
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
      regex:
        /(?:^|[_\s])(0?[1-9]|1[0-2])\s*월(?:[_\s]|$)/,
      period: (match) =>
        `${String(Number(match[1])).padStart(2, "0")}월`,
    },
    {
      regex:
        /(?:^|[_\s])((?:19|20|21)\d{2})[./-](0?[1-9]|1[0-2])(?:[_\s]|$)/,
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
  return (
    SUMMARY_LABEL_PATTERN.test(text) ||
    SUMMARY_SUFFIX_PATTERN.test(text)
  );
}

function summaryMetricBase(value = "") {
  return normalizeText(value).replace(/(?:[_/|>:\-]\s*|\s+)(?:계|합계|소계|총계)\s*$/i, "").trim();
}
function officialTotalHasDetailSiblings(value = "", allLabels = []) {
  const text = normalizeText(value);
  if (!/\s+합계\s*$/i.test(text)) return false;
  if (/[_/|>:\-]\s*합계\s*$/i.test(text)) return false;
  const base = summaryMetricBase(text);
  if (!base) return false;
  return (allLabels || []).filter((candidate) => {
    const label = normalizeText(candidate);
    return label && label !== text && label.startsWith(`${base} `) && !/(?:계|합계|소계|총계)\s*$/i.test(label);
  }).length >= 2;
}
function isSummaryMetricLabel(value = "", allLabels = []) {
  const text = normalizeText(value);
  if (!text) return false;
  if (SUMMARY_LABEL_PATTERN.test(text)) return true;
  if (/(?:[_/|>:\-]\s*|\s+)(?:소계|총계)\s*$/i.test(text)) return true;
  if (/(?:[_/|>:\-]\s*)(?:계|합계)\s*$/i.test(text)) return true;
  if (/\s+합계\s*$/i.test(text)) return !officialTotalHasDetailSiblings(text, allLabels);
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
    ) ||
    PERIOD_HEADER_PATTERN.test(normalizeText(header))
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

function inferAggregationContract({
  metricLabel = "",
  unit = "",
  column = {},
  declaredAggregation = "",
  hasTemporalAxis = false,
} = {}) {
  const fallbackAggregation =
    BUSINESS_AVERAGE_HEADER_PATTERN.test(
      [metricLabel, unit, semanticType(column), columnRole(column)]
        .map(normalizeText)
        .join(" "),
    )
      ? "average"
      : BUSINESS_SUM_HEADER_PATTERN.test(
          [metricLabel, unit, semanticType(column), columnRole(column)]
            .map(normalizeText)
            .join(" "),
        )
        ? "sum"
        : /%|퍼센트|백분율|비율|비중|구성비|점유율|증감률|달성률|지수|평균|평점|점수|시간|기록|속도|초|분초|cm|명\/천명|rate|ratio|share|percent|index|average|avg|score|duration|time/i.test(
            [metricLabel, unit].map(normalizeText).join(" "),
          )
          ? "average"
          : "sum";

  return resolveAggregationContract({
    metricLabel,
    unit,
    column,
    declaredAggregation,
    hasTemporalAxis,
    fallbackAggregation,
  });
}

function inferAggregation(options = {}) {
  return inferAggregationContract(options).operation;
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
  const canonicalMetricValueHeader =
    METRIC_VALUE_HEADER_PATTERN.test(metricValue.header);

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
    const value = normalizeText(
      rowValue(row, entry.column, entry.index),
    );
    if (value) values[entry.header] = value;
  }
  return values;
}

function shouldSkipDimensionRow(dimensionValues = {}) {
  return Object.values(dimensionValues).some(isSummaryLabel);
}

function canonicalLongSeries(table = {}, tableIndex = 0, contract = null, context = {}) {
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

  for (const [rowIndex, row] of rows.entries()) {
    let metricLabel = normalizeText(
      rowValue(
        row,
        contract.metricIdentity.column,
        contract.metricIdentity.index,
      ),
    );
    if (!metricLabel) continue;
    if (hasDetailMetric && isSummaryMetricLabel(metricLabel, metricLabels)) continue;

    const value = numericValue(
      rowValue(
        row,
        contract.metricValue.column,
        contract.metricValue.index,
      ),
    );
    if (value == null) continue;

    const dimensions = dimensionValuesForRow(
      row,
      contract.dimensions,
    );
    if (shouldSkipDimensionRow(dimensions)) continue;

    const unit = contract.unit
      ? normalizeText(
          rowValue(row, contract.unit.column, contract.unit.index),
        )
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

    const aggregationContract =
      inferAggregationContract({
        metricLabel,
        unit,
        column: contract.metricValue.column,
        declaredAggregation,
        hasTemporalAxis: Boolean(
          contract.period || contract.year,
        ),
      });
    const operation = aggregationContract.operation;

    const period = canonicalPeriodValue(
      contract.period
        ? rowValue(row, contract.period.column, contract.period.index)
        : "",
      contract.year
        ? rowValue(row, contract.year.column, contract.year.index)
        : "",
    );

    const key = [
      normalizeKey(metricLabel),
      normalizeKey(unit),
      aggregationContract.role,
      operation,
    ].join("::");

    if (!seriesMap.has(key)) {
      seriesMap.set(key, {
        key,
        tableIndex,
        metricLabel,
        unit,
        operation,
        metricRole: aggregationContract.role,
        aggregationContract,
        sourceContract: "canonical_long",
        valueHeader: contract.metricValue.header,
        protectedHeaders: contract.protectedHeaders,
        dimensionHeaders: contract.dimensions.map(
          (entry) => entry.header,
        ),
        records: [],
      });
    }

    seriesMap.get(key).records.push({
      value,
      period,
      dimensions,
      rowIndex,
    });
  }

  return [...seriesMap.values()].filter(
    (series) => series.records.length,
  );
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
    return (
      entry.profile.nonBlankCount > 0 &&
      entry.profile.numericRatio < 0.5
    );
  });

  const dimensionIndexes = new Set(
    dimensions.map((entry) => entry.index),
  );

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

  const explicit =
    normalizeText(table.metricLabel || table.measureName || "");
  if (explicit) return explicit;

  return "지표값";
}

function physicalWideSeries(table = {}, tableIndex = 0, contract = null, context = {}) {
  const rows = tableRows(table);
  const measureLabels = contract.measures.map((entry) => {
    const temporal = parseTemporalMeasureHeader(entry.header);
    return (
      temporal.metricLabel ||
      (
        temporal.period
          ? "지표값"
          : fallbackMetricLabel(table, entry.header)
      )
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
      (
        temporal.period
          ? "지표값"
          : fallbackMetricLabel(table, measure.header)
      );

    if (!metricLabel) return;
    if (hasDetailMetric && isSummaryMetricLabel(metricLabel, measureLabels)) return;

    const unit =
      normalizeText(
        measure.column.unit ||
          measure.column.measureUnit ||
          measure.column.meta?.unit ||
          "",
      ) ||
      temporal.unit;

    metricLabel = inferContextualMetricLabel({
      metricLabel,
      unit,
      table,
      context,
    });

    const aggregationContract =
      inferAggregationContract({
        metricLabel,
        unit,
        column: measure.column,
        hasTemporalAxis: Boolean(
          contract.period || contract.year || temporal.period,
        ),
      });
    const operation = aggregationContract.operation;

    const key = [
      normalizeKey(metricLabel),
      normalizeKey(unit),
      aggregationContract.role,
      operation,
    ].join("::");

    if (!seriesMap.has(key)) {
      seriesMap.set(key, {
        key,
        tableIndex,
        metricLabel,
        unit,
        operation,
        metricRole: aggregationContract.role,
        aggregationContract,
        sourceContract: "physical_wide",
        valueHeader: measure.header,
        protectedHeaders: contract.protectedHeaders,
        dimensionHeaders: contract.dimensions.map(
          (entry) => entry.header,
        ),
        records: [],
      });
    }

    for (const [rowIndex, row] of rows.entries()) {
      const value = numericValue(
        rowValue(row, measure.column, measure.index),
      );
      if (value == null) continue;

      const dimensions = dimensionValuesForRow(
        row,
        contract.dimensions,
      );
      if (shouldSkipDimensionRow(dimensions)) continue;

      const rowPeriod = temporal.period || strictPeriodValue(
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
        rowIndex,
      });
    }
  });

  return [...seriesMap.values()].filter(
    (series) => series.records.length,
  );
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
  return Array.isArray(tables) &&
    tables.some(
      (table, index) =>
        buildSemanticSeries(table, index, context).series.length > 0,
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

function comparablePeriod(value = "") {
  return normalizeText(value).toLowerCase();
}

function latestRecordSelection(records = []) {
  const list = (Array.isArray(records) ? records : []).filter(
    (record) =>
      record &&
      typeof record.value === "number" &&
      Number.isFinite(record.value),
  );
  if (!list.length) {
    return {
      records: [],
      period: "",
      method: "empty",
    };
  }

  const withPeriod = list.filter((record) => normalizeText(record.period));
  if (withPeriod.length) {
    const latestPeriod = withPeriod
      .map((record) => comparablePeriod(record.period))
      .sort((left, right) =>
        left.localeCompare(right, "ko", { numeric: true }),
      )
      .at(-1);
    return {
      records: withPeriod.filter(
        (record) => comparablePeriod(record.period) === latestPeriod,
      ),
      period: normalizeText(
        withPeriod.find(
          (record) => comparablePeriod(record.period) === latestPeriod,
        )?.period,
      ),
      method: "latest_period",
    };
  }

  const last = [...list].sort(
    (left, right) =>
      Number(left.rowIndex ?? 0) - Number(right.rowIndex ?? 0),
  ).at(-1);
  return {
    records: last ? [last] : [],
    period: "",
    method: "last_row",
  };
}

function snapshotEntityValue(record = {}, header = "") {
  return normalizeText(record?.dimensions?.[header]);
}

function snapshotEntityCompositeKey(record = {}, headers = []) {
  const values = (Array.isArray(headers) ? headers : [])
    .map((header) => snapshotEntityValue(record, header));
  if (!values.length || values.some((value) => !value)) return "";
  return values.map(normalizeKey).join("::");
}

function snapshotEntityCandidateScore({
  header = "",
  records = [],
} = {}) {
  const normalized = normalizeText(header);
  if (!normalized || SNAPSHOT_ENTITY_EXCLUDED_HEADER_PATTERN.test(normalized)) {
    return { header: normalized, score: -Infinity, eligible: false };
  }

  const values = (Array.isArray(records) ? records : [])
    .map((record) => snapshotEntityValue(record, normalized))
    .filter(Boolean);
  const rowCount = Math.max(1, records.length);
  const coverage = values.length / rowCount;
  const distinctCount = new Set(values.map(normalizeKey)).size;
  const repeatCount = Math.max(0, values.length - distinctCount);
  const identity = SNAPSHOT_ENTITY_IDENTITY_HEADER_PATTERN.test(normalized);
  const secondary = SNAPSHOT_ENTITY_SECONDARY_HEADER_PATTERN.test(normalized);

  if ((!identity && !secondary) || coverage < 0.7 || distinctCount < 2) {
    return {
      header: normalized,
      score: -Infinity,
      eligible: false,
      coverage,
      distinctCount,
      repeatCount,
      identity,
      secondary,
    };
  }

  let score = identity ? 240 : 110;
  if (/명$|이름$|명칭$/i.test(normalized)) score += 45;
  if (/(?:^|[_\s])(?:id|code)(?:$|[_\s])|번호|코드/i.test(normalized)) {
    score += 55;
  }
  if (coverage >= 0.95) score += 30;
  else if (coverage >= 0.85) score += 18;
  if (repeatCount > 0) score += 35;
  if (distinctCount <= 200) score += 20;
  else score -= 40;

  return {
    header: normalized,
    score,
    eligible: true,
    coverage,
    distinctCount,
    repeatCount,
    identity,
    secondary,
  };
}

function snapshotEntityDuplicateStats(records = [], headers = []) {
  const buckets = new Map();
  let observed = 0;
  for (const record of Array.isArray(records) ? records : []) {
    const key = snapshotEntityCompositeKey(record, headers);
    if (!key) continue;
    const period = comparablePeriod(record.period) || "__no_period__";
    const bucketKey = `${period}::${key}`;
    buckets.set(bucketKey, (buckets.get(bucketKey) || 0) + 1);
    observed += 1;
  }
  const duplicateCount = [...buckets.values()].reduce(
    (sum, count) => sum + Math.max(0, count - 1),
    0,
  );
  return {
    observed,
    duplicateCount,
    duplicateRate: observed ? duplicateCount / observed : 0,
  };
}

function resolveSnapshotEntityHeaders(series = {}) {
  const records = Array.isArray(series.records) ? series.records : [];
  if (
    series.operation !== SEMANTIC_AGGREGATION_OPERATION.LATEST ||
    series.metricRole !== SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT
  ) {
    return {
      version: SNAPSHOT_ENTITY_RESOLVER_VERSION,
      applied: false,
      headers: [],
      candidates: [],
      reason: "not_snapshot_latest",
    };
  }

  const headers = Array.from(new Set([
    ...(Array.isArray(series.dimensionHeaders)
      ? series.dimensionHeaders
      : []),
    ...records.flatMap((record) =>
      Object.keys(record?.dimensions || {}),
    ),
  ].map(normalizeText).filter(Boolean)));

  const candidates = headers
    .map((header) => snapshotEntityCandidateScore({ header, records }))
    .filter((candidate) => candidate.eligible)
    .sort((left, right) =>
      right.score - left.score ||
      right.distinctCount - left.distinctCount ||
      left.header.localeCompare(right.header, "ko"),
    );

  if (!candidates.length) {
    return {
      version: SNAPSHOT_ENTITY_RESOLVER_VERSION,
      applied: false,
      headers: [],
      candidates: [],
      reason: "no_entity_dimension",
    };
  }

  const selected = [candidates[0].header];
  const hasPeriod = records.some((record) => normalizeText(record.period));
  let duplicateStats = snapshotEntityDuplicateStats(records, selected);

  if (hasPeriod && duplicateStats.duplicateRate > 0) {
    for (const candidate of candidates.slice(1, 3)) {
      const trial = [...selected, candidate.header];
      const trialStats = snapshotEntityDuplicateStats(records, trial);
      if (trialStats.duplicateCount < duplicateStats.duplicateCount) {
        selected.push(candidate.header);
        duplicateStats = trialStats;
      }
      if (duplicateStats.duplicateRate <= 0.01) break;
    }
  }

  const keySet = new Set(
    records
      .map((record) => snapshotEntityCompositeKey(record, selected))
      .filter(Boolean),
  );

  return {
    version: SNAPSHOT_ENTITY_RESOLVER_VERSION,
    applied: keySet.size >= 2,
    headers: keySet.size >= 2 ? selected : [],
    candidates: candidates.map((candidate) => ({
      header: candidate.header,
      score: candidate.score,
      coverage: candidate.coverage,
      distinctCount: candidate.distinctCount,
      repeatCount: candidate.repeatCount,
    })),
    entityCount: keySet.size,
    duplicateCountWithinPeriod: duplicateStats.duplicateCount,
    duplicateRateWithinPeriod: duplicateStats.duplicateRate,
    reason: keySet.size >= 2
      ? "entity_dimension_resolved"
      : "insufficient_entity_count",
  };
}

function latestRecordSelectionByEntity(records = [], entityHeaders = []) {
  const list = (Array.isArray(records) ? records : []).filter(
    (record) =>
      record &&
      typeof record.value === "number" &&
      Number.isFinite(record.value),
  );
  const headers = (Array.isArray(entityHeaders) ? entityHeaders : [])
    .map(normalizeText)
    .filter(Boolean);
  if (!list.length || !headers.length) {
    return latestRecordSelection(list);
  }

  const grouped = new Map();
  for (const record of list) {
    const key = snapshotEntityCompositeKey(record, headers) ||
      `__row__${Number(record.rowIndex ?? grouped.size)}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(record);
  }

  const selectedRecords = [];
  const selectedPeriods = [];
  for (const entityRecords of grouped.values()) {
    const selected = latestRecordSelection(entityRecords);
    const chosen = [...selected.records]
      .sort((left, right) =>
        Number(left.rowIndex ?? 0) - Number(right.rowIndex ?? 0),
      )
      .at(-1);
    if (chosen) selectedRecords.push(chosen);
    if (selected.period) selectedPeriods.push(selected.period);
  }

  const uniquePeriods = Array.from(
    new Set(selectedPeriods.map(normalizeText).filter(Boolean)),
  ).sort((left, right) =>
    comparablePeriod(left).localeCompare(
      comparablePeriod(right),
      "ko",
      { numeric: true },
    ),
  );
  const period = uniquePeriods.length === 1
    ? uniquePeriods[0]
    : uniquePeriods.length > 1
      ? `${uniquePeriods[0]} ~ ${uniquePeriods.at(-1)} (엔티티별 최신)`
      : "행 순서 기준";

  return {
    records: selectedRecords,
    period,
    method: "latest_by_entity",
    entityHeaders: headers,
    entityCount: grouped.size,
    selectedPeriods: uniquePeriods,
  };
}

function operationStats(records = [], operation = "sum", options = {}) {
  const allStats = numberStats(
    records.map((record) => record.value),
  );
  if (operation === SEMANTIC_AGGREGATION_OPERATION.LATEST) {
    const entityHeaders = (Array.isArray(options.entityHeaders)
      ? options.entityHeaders
      : []).map(normalizeText).filter(Boolean);
    const selection = entityHeaders.length
      ? latestRecordSelectionByEntity(records, entityHeaders)
      : latestRecordSelection(records);
    const selectedStats = numberStats(
      selection.records.map((record) => record.value),
    );
    return {
      allStats,
      selectedStats,
      selectedRecords: selection.records,
      selectedPeriod: selection.period,
      selectionMethod: selection.method,
      entityHeaders: selection.entityHeaders || [],
      entityCount: Number(selection.entityCount || 0),
      selectedPeriods: selection.selectedPeriods || [],
      value: selectedStats.sum,
    };
  }
  return {
    allStats,
    selectedStats: allStats,
    selectedRecords: records,
    selectedPeriod: "",
    selectionMethod: operation,
    entityHeaders: [],
    entityCount: 0,
    selectedPeriods: [],
    value:
      operation === "average" ? allStats.average : allStats.sum,
  };
}

function groupRows({
  records = [],
  groupHeader = "그룹",
  operation = "sum",
  metricLabel = "지표값",
  unit = "",
  groupValue,
  entityHeaders = [],
} = {}) {
  const grouped = new Map();

  for (const record of records) {
    const key = normalizeText(groupValue(record));
    if (!key) continue;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(record);
  }

  return [...grouped.entries()]
    .map(([group, groupedRecords]) => {
      const resolved = operationStats(groupedRecords, operation, {
        entityHeaders,
      });
      return {
        [groupHeader]: group,
        operation,
        metric: metricLabel,
        value: resolved.value,
        rowCount: resolved.allStats.count,
        선택행수: resolved.selectedStats.count,
        기준기간: resolved.selectedPeriod,
        단위: unit,
        평균: resolved.selectedStats.average,
        최솟값: resolved.selectedStats.min,
        최댓값: resolved.selectedStats.max,
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

function dimensionSemanticPriority({
  header = "",
  distinctCount = 0,
} = {}) {
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
  const sectionPolicy = sectionPolicyForSeries(series, options);
  const maxDimensionsPerSeries = sectionPolicy.maxDimensions;
  const baseId = metricId(
    tableIndex,
    series.metricLabel,
    series.unit,
    series.operation,
  );
  const stats = numberStats(
    series.records.map((record) => record.value),
  );
  const additive = series.operation === "sum";
  const snapshot =
    series.operation === SEMANTIC_AGGREGATION_OPERATION.LATEST;
  const durationMetric =
    series.metricRole === SEMANTIC_METRIC_ROLE.DURATION;
  const snapshotEntityResolution = snapshot
    ? resolveSnapshotEntityHeaders(series)
    : {
        version: SNAPSHOT_ENTITY_RESOLVER_VERSION,
        applied: false,
        headers: [],
        entityCount: 0,
        reason: "not_snapshot_latest",
      };
  const snapshotEntityHeaders = snapshotEntityResolution.applied
    ? snapshotEntityResolution.headers
    : [];
  const resolvedSummary = operationStats(
    series.records,
    series.operation,
    { entityHeaders: snapshotEntityHeaders },
  );
  const summaryTitle = snapshot
    ? `${titlePrefix}${displayMetric} 최신 스냅샷`
    : `${titlePrefix}${displayMetric} 통계`;
  const summaryRows = snapshot
    ? [
        { 지표: "전체 유효값 수", 값: stats.count, 단위: "건" },
        {
          지표: "최신 기준기간",
          값: resolvedSummary.selectedPeriod || "행 순서 기준",
        },
        ...(snapshotEntityResolution.applied
          ? [
              {
                지표: "Snapshot 선택 방식",
                값: "엔티티별 최신",
              },
              {
                지표: "엔티티 기준",
                값: snapshotEntityHeaders.join(" + "),
              },
              {
                지표: "엔티티 수",
                값: resolvedSummary.entityCount,
                단위: "건",
              },
            ]
          : []),
        {
          지표: "최신 스냅샷 값",
          값: resolvedSummary.value,
          단위: series.unit,
        },
        {
          지표: "선택 행 수",
          값: resolvedSummary.selectedStats.count,
          단위: "건",
        },
        {
          지표: "최솟값",
          값: resolvedSummary.selectedStats.min,
          단위: series.unit,
        },
        {
          지표: "최댓값",
          값: resolvedSummary.selectedStats.max,
          단위: series.unit,
        },
      ]
    : additive
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
          ...(durationMetric
            ? [{
                지표: "중앙값",
                값: median(series.records.map((record) => record.value)),
                단위: series.unit,
              }]
            : []),
          { 지표: "최솟값", 값: stats.min, 단위: series.unit },
          { 지표: "최댓값", 값: stats.max, 단위: series.unit },
        ];

  const sections = [
    {
      sectionId: `${baseId}.summary`,
      title: summaryTitle,
      sectionType: snapshot
        ? "semantic_snapshot_summary"
        : additive
          ? "semantic_additive_summary"
          : "semantic_non_additive_summary",
      metricIds: [`${baseId}.summary`],
      result: {
        ok: true,
        resultType: "pivot",
        operation: snapshot
          ? "semanticLatestSnapshot"
          : additive
            ? "semanticAggregateSummary"
            : "semanticAverageRange",
        metric: {
          header: series.metricLabel,
          unit: series.unit,
        },
        rows: summaryRows,
        meta: {
          metricIds: [`${baseId}.summary`],
          complete: true,
          additive,
          snapshot,
          snapshotEntityResolverVersion:
            SNAPSHOT_ENTITY_RESOLVER_VERSION,
          snapshotEntitySelectionApplied:
            snapshotEntityResolution.applied,
          snapshotEntityHeaders:
            cloneValue(snapshotEntityHeaders),
          snapshotEntityCount:
            Number(resolvedSummary.entityCount || 0),
          snapshotSelectionMethod:
            resolvedSummary.selectionMethod,
          metricRole: series.metricRole,
          aggregationContract: cloneValue(series.aggregationContract || {}),
          unit: series.unit,
          sourceContract: series.sourceContract,
          valueHeader: series.valueHeader,
          protectedHeaders: series.protectedHeaders,
          plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
          metricRelationshipPriorityEngineVersion:
            METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
          relationshipRole: series.relationshipRole || "independent",
          representativeMetricPriority:
            Number(series.representativeMetricPriority || 0),
          componentOfMetric: series.componentOfMetric || "",
          metricRelationships:
            cloneValue(series.metricRelationships || []),
          semanticSectionBudgetEngineVersion:
            SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
          sectionBudgetPolicy: cloneValue(sectionPolicy),
          sectionBudgetPriority: sectionPolicy.budgetPriority,
          durationSummaryContractVersion: durationMetric
            ? DURATION_SUMMARY_CONTRACT_VERSION
            : "",
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
    .filter(
      (entry) =>
        entry.distinctCount >= 2 &&
        entry.distinctCount <= 200,
    )
    .sort(
      (left, right) =>
        right.semanticPriority - left.semanticPriority ||
        right.distinctCount - left.distinctCount ||
        left.header.localeCompare(right.header, "ko"),
    )
    .slice(0, maxDimensionsPerSeries);

  for (const dimension of dimensionHeaders) {
    const groupSourceRecords =
      snapshot && snapshotEntityResolution.applied
        ? resolvedSummary.selectedRecords
        : series.records;
    const groupOperation =
      snapshot && snapshotEntityResolution.applied
        ? SEMANTIC_AGGREGATION_OPERATION.SUM
        : series.operation;
    let rows = groupRows({
      records: groupSourceRecords,
      groupHeader: dimension.header,
      operation: groupOperation,
      metricLabel: series.metricLabel,
      unit: series.unit,
      groupValue: (record) => record.dimensions?.[dimension.header],
    });
    if (snapshot && snapshotEntityResolution.applied) {
      rows = rows.map((row) => ({
        ...row,
        operation: "latest_by_entity",
        기준기간: resolvedSummary.selectedPeriod,
      }));
    }
    if (!rows.length) continue;

    sections.push({
      sectionId: `${baseId}.by_${safeId(dimension.header)}`,
      title: `${titlePrefix}${dimension.header}별 ${displayMetric}`,
      sectionType: snapshot
        ? "semantic_group_snapshot"
        : additive
          ? "semantic_group_sum"
          : "semantic_group_average",
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
          snapshot,
          snapshotEntityResolverVersion:
            SNAPSHOT_ENTITY_RESOLVER_VERSION,
          snapshotEntitySelectionApplied:
            snapshotEntityResolution.applied,
          snapshotEntityHeaders:
            cloneValue(snapshotEntityHeaders),
          snapshotEntityCount:
            Number(resolvedSummary.entityCount || 0),
          snapshotSelectionMethod:
            resolvedSummary.selectionMethod,
          metricRole: series.metricRole,
          aggregationContract: cloneValue(series.aggregationContract || {}),
          unit: series.unit,
          sourceContract: series.sourceContract,
          valueHeader: series.valueHeader,
          plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
          metricRelationshipPriorityEngineVersion:
            METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
          relationshipRole: series.relationshipRole || "independent",
          representativeMetricPriority:
            Number(series.representativeMetricPriority || 0),
          componentOfMetric: series.componentOfMetric || "",
          metricRelationships:
            cloneValue(series.metricRelationships || []),
          semanticSectionBudgetEngineVersion:
            SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
          sectionBudgetPolicy: cloneValue(sectionPolicy),
          sectionBudgetPriority: sectionPolicy.budgetPriority,
          durationSummaryContractVersion: durationMetric
            ? DURATION_SUMMARY_CONTRACT_VERSION
            : "",
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
    entityHeaders: snapshotEntityHeaders,
  });

  if (sectionPolicy.includePeriod && periodRows.length) {
    sections.push({
      sectionId: `${baseId}.by_period`,
      title: `${titlePrefix}기간별 ${displayMetric}`,
      sectionType: snapshot
        ? "semantic_period_snapshot"
        : additive
          ? "semantic_period_sum"
          : "semantic_period_average",
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
          snapshot,
          snapshotEntityResolverVersion:
            SNAPSHOT_ENTITY_RESOLVER_VERSION,
          snapshotEntitySelectionApplied:
            snapshotEntityResolution.applied,
          snapshotEntityHeaders:
            cloneValue(snapshotEntityHeaders),
          snapshotEntityCount:
            Number(resolvedSummary.entityCount || 0),
          snapshotSelectionMethod:
            resolvedSummary.selectionMethod,
          metricRole: series.metricRole,
          aggregationContract: cloneValue(series.aggregationContract || {}),
          explicitPeriodOnly: true,
          unit: series.unit,
          sourceContract: series.sourceContract,
          valueHeader: series.valueHeader,
          protectedHeaders: series.protectedHeaders,
          plannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
          metricRelationshipPriorityEngineVersion:
            METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
          relationshipRole: series.relationshipRole || "independent",
          representativeMetricPriority:
            Number(series.representativeMetricPriority || 0),
          componentOfMetric: series.componentOfMetric || "",
          metricRelationships:
            cloneValue(series.metricRelationships || []),
          semanticSectionBudgetEngineVersion:
            SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
          sectionBudgetPolicy: cloneValue(sectionPolicy),
          sectionBudgetPriority: sectionPolicy.budgetPriority,
          durationSummaryContractVersion: durationMetric
            ? DURATION_SUMMARY_CONTRACT_VERSION
            : "",
        },
      },
    });
  }

  return sections.slice(0, sectionPolicy.maxSections);
}

function semanticScalarText(value, output = [], depth = 0) {
  if (depth > 5 || value == null) return output;
  if (["string", "number", "boolean"].includes(typeof value)) {
    const text = normalizeText(value);
    if (text) output.push(text);
    return output;
  }
  if (Array.isArray(value)) {
    value.slice(0, 100).forEach((item) =>
      semanticScalarText(item, output, depth + 1),
    );
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
                row && typeof row === "object"
                  ? Object.keys(row)
                  : [],
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
  if (/latest|snapshot|최신|스냅샷|기말/.test(text)) return "latest";
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

  const existingGeneric =
    GENERIC_METRIC_LABEL_PATTERN.test(existing);
  const plannedGeneric =
    GENERIC_METRIC_LABEL_PATTERN.test(planned);

  if (existingGeneric !== plannedGeneric) return false;

  return normalizeKey(existing) === normalizeKey(planned);
}

function groupHeadersSemanticallyMatch(
  existingGroup = "",
  plannedGroup = "",
) {
  const existing = normalizeText(existingGroup);
  const planned = normalizeText(plannedGroup);

  if (!existing || !planned) return true;

  if (
    BUSINESS_TIME_HEADER_PATTERN.test(existing) &&
    BUSINESS_TIME_HEADER_PATTERN.test(planned)
  ) {
    return true;
  }

  const aliases = new Set(
    semanticGroupAliases(planned).map(normalizeKey),
  );
  return aliases.has(normalizeKey(existing));
}

function isWholeMetricSummarySection(section = {}) {
  const sectionType = normalizeText(section.sectionType || "");
  const group = normalizeText(
    section.result?.groupBy?.header || "",
  );
  return !group && /summary/.test(sectionType);
}

function hasGroupedOrRankedSectionIdentity(section = {}) {
  if (normalizeText(explicitSectionGroupHeader(section))) return true;
  const title = normalizeText(section.title || "");
  return /(?:^|\s)(?:기간|연도|년도|연월|월|일|부서|팀|상태|분류|유형|지역|품목|상품|제품|자산|장비|시설|담당자|고객|거래처|기관|사업|과제|프로젝트)별|상위|하위|순위|구성비|증감률|추이/i.test(
    title,
  );
}

function existingSectionCoversPlanned(existing = {}, planned = {}) {
  const metric = normalizeText(
    planned.result?.metric?.header || "",
  );
  const group = normalizeText(
    planned.result?.groupBy?.header || "",
  );
  const text = semanticSectionText(existing);
  const existingMetric =
    explicitSectionMetricHeader(existing);
  const existingGroup =
    explicitSectionGroupHeader(existing);

  if (
    isWholeMetricSummarySection(planned) &&
    hasGroupedOrRankedSectionIdentity(existing)
  ) {
    return false;
  }

  if (!metric || !sectionContainsToken(text, metric)) return false;

  if (
    !metricHeadersSemanticallyMatch(
      existingMetric,
      metric,
    )
  ) {
    return false;
  }

  if (
    group &&
    !groupHeadersSemanticallyMatch(
      existingGroup,
      group,
    )
  ) {
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
  if (plannedFamily === "latest") {
    return existingFamily === "latest";
  }
  if (plannedFamily === "average") {
    return ["average", "summary"].includes(existingFamily) ||
      /평균|점수\s*요약/.test(normalizeText(existing.title));
  }
  if (plannedFamily === "sum") {
    return ["sum", "summary"].includes(existingFamily) ||
      /합계|금액\s*요약|수량\s*요약/.test(normalizeText(existing.title));
  }
  if (!group) {
    return ["summary", "average", "sum"].includes(existingFamily) ||
      /요약|통계/.test(normalizeText(existing.title));
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

function applyMandatorySummaryCoverageFloor({
  sections = [],
  plannedSections = [],
  expectedMetricIds = [],
} = {}) {
  const working = Array.isArray(sections)
    ? sections.map((section) => cloneValue(section))
    : [];
  const expected = new Set(uniqueMetricIds(expectedMetricIds));
  const renderedBefore = new Set(
    uniqueMetricIds(
      working.flatMap((section) => collectSectionMetricIds(section)),
    ),
  );
  const candidates = (Array.isArray(plannedSections)
    ? plannedSections
    : []
  ).filter((section) => {
    if (!isWholeMetricSummarySection(section)) return false;
    return collectSectionMetricIds(section).some(
      (metricId) => expected.has(metricId),
    );
  });

  const restoredMetricIds = [];
  const transferredMetricIds = [];
  const restoredSectionIds = [];
  const coverageActions = [];

  for (const planned of candidates) {
    const plannedMetricIds = collectSectionMetricIds(planned).filter(
      (metricId) => expected.has(metricId),
    );
    const missingMetricIds = plannedMetricIds.filter(
      (metricId) => !renderedBefore.has(metricId),
    );
    if (!missingMetricIds.length) continue;

    const targetIndex = working.findIndex((section) =>
      existingSectionCoversPlanned(section, planned),
    );
    if (targetIndex >= 0) {
      working[targetIndex] = annotateExistingSection(
        working[targetIndex],
        planned,
      );
      const targetMetricIds = collectSectionMetricIds(
        working[targetIndex],
      );
      for (const metricId of missingMetricIds) {
        renderedBefore.add(metricId);
        transferredMetricIds.push(metricId);
      }
      coverageActions.push({
        action: "transfer_to_authoritative_summary",
        plannedSectionId: normalizeText(planned.sectionId || ""),
        targetSectionId: normalizeText(
          working[targetIndex]?.sectionId || "",
        ),
        targetTitle: normalizeText(working[targetIndex]?.title || ""),
        metricIds: targetMetricIds.filter((metricId) =>
          missingMetricIds.includes(metricId),
        ),
      });
      continue;
    }

    const restored = applySectionMetricIds(planned);
    restored.result = {
      ...(restored.result || {}),
      meta: {
        ...(restored.result?.meta || {}),
        mandatorySummaryCoverageFloorVersion:
          MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
        restoredByMandatorySummaryCoverageFloor: true,
      },
    };
    const metricHeader = normalizeText(
      restored.result?.metric?.header || "",
    );
    const insertIndex = working.findIndex((section) =>
      metricHeadersSemanticallyMatch(
        explicitSectionMetricHeader(section),
        metricHeader,
      ),
    );
    if (insertIndex >= 0) working.splice(insertIndex, 0, restored);
    else working.push(restored);

    for (const metricId of missingMetricIds) {
      renderedBefore.add(metricId);
      restoredMetricIds.push(metricId);
    }
    restoredSectionIds.push(
      normalizeText(restored.sectionId || restored.title || ""),
    );
    coverageActions.push({
      action: "restore_planned_summary",
      plannedSectionId: normalizeText(planned.sectionId || ""),
      targetSectionId: normalizeText(restored.sectionId || ""),
      targetTitle: normalizeText(restored.title || ""),
      metricIds: missingMetricIds,
    });
  }

  const renderedAfter = new Set(
    uniqueMetricIds(
      working.flatMap((section) => collectSectionMetricIds(section)),
    ),
  );
  const missingBefore = [...expected].filter(
    (metricId) => !new Set(
      uniqueMetricIds(
        sections.flatMap((section) => collectSectionMetricIds(section)),
      ),
    ).has(metricId),
  );
  const missingAfter = [...expected].filter(
    (metricId) => !renderedAfter.has(metricId),
  );

  return {
    version: MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
    sections: working,
    applied: coverageActions.length > 0,
    expectedMetricIdCount: expected.size,
    missingMetricIdsBefore: missingBefore,
    missingMetricIdsAfter: missingAfter,
    restoredMetricIds: uniqueMetricIds(restoredMetricIds),
    transferredMetricIds: uniqueMetricIds(transferredMetricIds),
    restoredSectionIds,
    coverageActions,
    pass: missingAfter.length === 0,
  };
}

function sectionMetricAndGroupMatch(
  existing = {},
  planned = {},
) {
  const metric = normalizeText(
    planned.result?.metric?.header || "",
  );
  const group = normalizeText(
    planned.result?.groupBy?.header || "",
  );
  const text = semanticSectionText(existing);
  const existingMetric =
    explicitSectionMetricHeader(existing);
  const existingGroup =
    explicitSectionGroupHeader(existing);

  if (
    !metric ||
    !sectionContainsToken(text, metric)
  ) {
    return false;
  }

  if (
    !metricHeadersSemanticallyMatch(
      existingMetric,
      metric,
    )
  ) {
    return false;
  }

  if (!group) {
    return !existingGroup;
  }

  if (
    existingGroup &&
    !groupHeadersSemanticallyMatch(
      existingGroup,
      group,
    )
  ) {
    return false;
  }

  if (
    BUSINESS_TIME_HEADER_PATTERN.test(group)
  ) {
    return Boolean(
      BUSINESS_TIME_HEADER_PATTERN.test(
        existingGroup,
      ) ||
      BUSINESS_TIME_HEADER_PATTERN.test(
        normalizeText(existing.title),
      ) ||
      semanticGroupAliases(group).some(
        (alias) =>
          sectionContainsToken(text, alias),
      ),
    );
  }

  if (existingGroup) return true;

  return semanticGroupAliases(group).some(
    (alias) =>
      sectionContainsToken(text, alias),
  );
}

function sectionOperationFamilies(
  section = {},
) {
  const families = new Set();
  const directFamily =
    semanticOperationFamily(section);
  if (directFamily && directFamily !== "other") {
    families.add(directFamily);
  }

  const rows = Array.isArray(
    section.result?.rows,
  )
    ? section.result.rows.slice(0, 100)
    : [];

  for (const row of rows) {
    if (
      !row ||
      typeof row !== "object" ||
      Array.isArray(row)
    ) {
      continue;
    }

    const text = normalizeText([
      row.작업,
      row.operation,
      row.집계,
      row.집계유형,
      row.지표,
      row.metric,
    ].filter(Boolean).join(" ")).toLowerCase();

    if (!text) continue;
    if (/latest|snapshot|최신|스냅샷|기말/.test(text)) {
      families.add("latest");
    }
    if (/average|mean|avg|평균/.test(text)) {
      families.add("average");
    }
    if (/sum|total|합계|합산|총계/.test(text)) {
      families.add("sum");
    }
    if (/count|건수/.test(text)) {
      families.add("count");
    }
    if (/top|bottom|rank|상위|하위|순위/.test(text)) {
      families.add("rank");
    }
    if (/ratio|composition|비율|구성비/.test(text)) {
      families.add("ratio");
    }
  }

  return families;
}

function firstSemanticRowText(row = {}, keys = []) {
  for (const key of keys) {
    const value = normalizeText(row?.[key]);
    if (value) return value;
  }
  return "";
}

function semanticRowMetricLabel(row = {}) {
  return firstSemanticRowText(row, [
    "지표",
    "metric",
    "지표명",
    "항목",
    "label",
    "name",
  ]);
}

function semanticRowValue(row = {}) {
  for (const key of [
    "값",
    "value",
    "수치",
    "amount",
    "합계",
    "평균",
  ]) {
    const value = Number(row?.[key]);
    if (Number.isFinite(value)) {
      return { key, value };
    }
  }
  return { key: "", value: NaN };
}

function isGeneralStockSnapshotAlias(label = "") {
  const normalized = normalizeText(label);
  if (!normalized) return false;
  if (
    /기초|시작|opening|안전|safety|금액|비용|value|amount|잔여|잔량|remaining/i.test(
      normalized,
    )
  ) {
    return false;
  }
  if (
    /입고|출고|이동|조정|증감|사용|대여|판매|생산|inbound|outbound|movement|adjustment|flow/i.test(
      normalized,
    )
  ) {
    return false;
  }

  const key = normalizeKey(normalized);
  return (
    /^(?:행별|전체|총|평균|누적)?재고(?:스냅샷|수량|잔액|보유수량)(?:합계|총계|평균|값)?$/.test(
      key,
    ) ||
    /^(?:스냅샷|기말)재고(?:수량|잔액|보유수량)?(?:합계|총계|값)?$/.test(
      key,
    ) ||
    /^(?:row|all|total|average|cumulative)?(?:inventory|stock)(?:snapshot|quantity|balance|onhand)(?:sum|total|average)?$/.test(
      key,
    )
  );
}

function stockSnapshotSemanticSubtype(label = "") {
  const normalized = normalizeText(label);
  if (!normalized) return "unknown";
  if (/금액|비용|value|amount/i.test(normalized)) return "money";
  if (/기초|시작|opening/i.test(normalized)) return "opening";
  if (/안전|safety/i.test(normalized)) return "safety";
  if (/잔여|잔량|remaining/i.test(normalized)) return "remaining";
  if (/현재|기말|on\s*hand|onhand|closing/i.test(normalized)) {
    return "current";
  }
  if (isGeneralStockSnapshotAlias(normalized)) return "generic_current";
  if (/재고|inventory|stock|보유/i.test(normalized)) return "generic";
  return "unknown";
}

function semanticRoleForMetricLabel(label = "") {
  const normalized = normalizeText(label);
  const classified = classifyMetricRole({ metricLabel: normalized });
  const generalStockAlias = isGeneralStockSnapshotAlias(normalized);

  /*
   * inventoryFlowReportBuilder의 실제 overview 행은 Semantic Planner보다
   * 먼저 `재고 수량 합계` 형태로 전달된다. 이 라벨은 일반 수량 패턴에도
   * 걸리므로, 공통 stock alias를 기본 classifier보다 우선해야 한다.
   * 이후 businessTemplateExecutor가 붙이는 `행별 재고 스냅샷 합계` 라벨과
   * 동일한 의미 계약으로 취급한다.
   */
  if (generalStockAlias) {
    return {
      ...classified,
      role: SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT,
      confidence: Math.max(Number(classified.confidence || 0), 0.99),
      source: "general_stock_snapshot_alias_preclassification",
      aliasVersion: GENERAL_STOCK_SNAPSHOT_ALIAS_VERSION,
      overriddenRole: classified.role,
    };
  }

  if (
    classified.role === SEMANTIC_METRIC_ROLE.GENERIC_MEASURE &&
    /재고.*(?:스냅샷|잔액|잔량|합계)|(?:스냅샷|기말).*재고|최다.*보유|보유.*(?:수량|품목)/i.test(normalized)
  ) {
    return {
      ...classified,
      role: SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT,
      confidence: Math.max(Number(classified.confidence || 0), 0.92),
      source: "mixed_row_stock_snapshot_pattern",
      aliasVersion: "",
    };
  }
  return classified;
}

function sectionAuthoritativeMetricHeader(section = {}) {
  return normalizeText(
    section.result?.metric?.header ||
      section.metric?.header ||
      section.metricHeader ||
      "",
  );
}

function sectionAuthoritativeScalar(section = {}, family = "") {
  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  const wanted = family || semanticOperationFamily(section);
  const patterns = wanted === "latest"
    ? [/최신\s*스냅샷\s*값|최신값|기말값|latest\s*(?:snapshot\s*)?value/i]
    : [/평균|average|mean|avg/i];

  for (const pattern of patterns) {
    const row = rows.find((item) =>
      item &&
      typeof item === "object" &&
      !Array.isArray(item) &&
      pattern.test(semanticRowMetricLabel(item)),
    );
    if (!row) continue;
    const numeric = semanticRowValue(row);
    if (Number.isFinite(numeric.value)) {
      return {
        value: numeric.value,
        valueKey: numeric.key,
        row,
      };
    }
  }

  const direct = Number(
    section.result?.value ??
      section.value ??
      section.result?.total,
  );
  if (Number.isFinite(direct)) {
    return { value: direct, valueKey: "value", row: null };
  }

  return { value: NaN, valueKey: "", row: null };
}

function sectionAuthoritativeSelectedCount(section = {}) {
  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  for (const row of rows) {
    if (
      !row ||
      typeof row !== "object" ||
      Array.isArray(row)
    ) {
      continue;
    }
    const label = semanticRowMetricLabel(row);
    if (!/선택\s*행\s*수|selected\s*row\s*count/i.test(label)) {
      continue;
    }
    const numeric = semanticRowValue(row);
    if (Number.isFinite(numeric.value) && numeric.value > 0) {
      return numeric.value;
    }
  }
  return 0;
}

function semanticMetricCompatibilityScore({
  requestedLabel = "",
  candidateMetric = "",
} = {}) {
  const requested = normalizeText(requestedLabel);
  const candidate = normalizeText(candidateMetric);
  if (!requested || !candidate) return -Infinity;

  const requestedRole = semanticRoleForMetricLabel(requested);
  const candidateRole = semanticRoleForMetricLabel(candidate);
  if (requestedRole.role !== candidateRole.role) return -Infinity;

  const requestedKey = normalizeKey(requested);
  const candidateKey = normalizeKey(candidate);
  let score = 10;

  if (requestedKey.includes(candidateKey)) score += 60;
  if (candidateKey.includes(requestedKey)) score += 40;

  if (requestedRole.role === SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT) {
    const requestedSubtype = stockSnapshotSemanticSubtype(requested);
    const candidateSubtype = stockSnapshotSemanticSubtype(candidate);

    if (requestedSubtype === candidateSubtype) score += 48;
    if (["generic_current", "generic"].includes(requestedSubtype)) {
      if (candidateSubtype === "current") {
        score += requestedSubtype === "generic_current" ? 96 : 60;
      } else if (candidateSubtype === "remaining") {
        score += requestedSubtype === "generic_current" ? 42 : 36;
      } else if (candidateSubtype === "generic") {
        score += 24;
      } else if (["money", "opening", "safety"].includes(candidateSubtype)) {
        score -= requestedSubtype === "generic_current" ? 90 : 60;
      }
    } else if (
      ["money", "opening", "safety", "remaining", "current"].includes(
        requestedSubtype,
      ) &&
      ["money", "opening", "safety", "remaining", "current"].includes(
        candidateSubtype,
      ) &&
      requestedSubtype !== candidateSubtype
    ) {
      score -= 54;
    }
  }

  const sharedTokens = [
    "재고",
    "수량",
    "금액",
    "잔여",
    "현재",
    "기말",
    "기초",
    "안전",
    "단가",
    "일수",
    "기간",
  ].filter((token) =>
    requested.includes(token) && candidate.includes(token),
  );
  score += sharedTokens.length * 8;

  const requestedMoney = /금액|비용|원가|amount|cost|value/i.test(requested);
  const candidateMoney = /금액|비용|원가|amount|cost|value/i.test(candidate);
  if (requestedMoney === candidateMoney) score += 8;
  else score -= 30;

  const requestedOpening = /기초|opening/i.test(requested);
  const requestedSafety = /안전|safety/i.test(requested);
  const requestedCurrent = /현재|기말|잔여|on\s*hand|closing|remaining/i.test(requested);

  if (/현재|기말|on\s*hand|closing/i.test(candidate)) {
    score += requestedOpening || requestedSafety ? -12 : 24;
    if (requestedCurrent) score += 12;
  }
  if (/잔여|remaining/i.test(candidate)) {
    score += requestedOpening || requestedSafety ? -8 : 18;
  }
  if (/기초|opening/i.test(candidate)) {
    score += requestedOpening ? 22 : -14;
  }
  if (/안전|safety/i.test(candidate)) {
    score += requestedSafety ? 22 : -16;
  }

  return score;
}

function authoritativeSemanticSectionIndex({
  sections = [],
  requestedLabel = "",
  family = "latest",
  groupHeader = "",
  excludedIndex = -1,
} = {}) {
  let bestIndex = -1;
  let bestScore = -Infinity;

  sections.forEach((section, index) => {
    if (index === excludedIndex) return;
    if (semanticOperationFamily(section) !== family) return;

    const candidateGroup = explicitSectionGroupHeader(section);
    if (groupHeader) {
      if (!groupHeadersSemanticallyMatch(candidateGroup, groupHeader)) {
        return;
      }
    } else if (candidateGroup) {
      return;
    }

    const metric = sectionAuthoritativeMetricHeader(section);
    const scalar = sectionAuthoritativeScalar(section, family);
    const groupedRows = groupHeader
      ? candidateGroupedRows(section)
      : [];
    if (
      !metric ||
      (groupHeader
        ? groupedRows.length === 0
        : !Number.isFinite(scalar.value))
    ) {
      return;
    }

    const score = semanticMetricCompatibilityScore({
      requestedLabel,
      candidateMetric: metric,
    });
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore >= 10 ? bestIndex : -1;
}


function isInventoryFlowMixedSection(section = {}) {
  const text = normalizeText([
    section.sectionId,
    section.sectionType,
    section.title,
    section.result?.resultType,
    section.result?.operation,
  ].filter(Boolean).join(" ")).toLowerCase();

  return (
    /inventory[_\s-]*flow[_\s-]*(?:overview|summary)/i.test(text) ||
    /재고.*입출고.*(?:흐름|요약)|입출고.*(?:흐름|요약)/i.test(text)
  );
}

function semanticFamilyForRole(role = "") {
  if (role === SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT) return "latest";
  if (
    [
      SEMANTIC_METRIC_ROLE.UNIT_RATE,
      SEMANTIC_METRIC_ROLE.DURATION,
      SEMANTIC_METRIC_ROLE.PERCENTAGE_RATE,
    ].includes(role)
  ) {
    return "average";
  }
  if (
    [
      SEMANTIC_METRIC_ROLE.MONEY_FLOW,
      SEMANTIC_METRIC_ROLE.QUANTITY_FLOW,
      SEMANTIC_METRIC_ROLE.COUNT,
    ].includes(role)
  ) {
    return "sum";
  }
  return "";
}

function blockedFlowMetricTransferTargetIndex({
  sections = [],
  section = {},
  excludedIndex = -1,
} = {}) {
  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  let bestIndex = -1;
  let bestScore = -Infinity;

  for (const row of rows) {
    if (!row || typeof row !== "object" || Array.isArray(row)) continue;
    const label = semanticRowMetricLabel(row);
    if (!label) continue;
    const role = semanticRoleForMetricLabel(label);
    const wantedFamily = semanticFamilyForRole(role.role);
    if (!wantedFamily) continue;

    sections.forEach((candidate, index) => {
      if (index === excludedIndex) return;
      const sectionType = normalizeText(candidate.sectionType).toLowerCase();
      const plannerId = normalizeText(
        candidate.result?.meta?.semanticCoverage?.plannerSectionId || "",
      );
      if (!sectionType.startsWith("semantic_") && !plannerId) return;
      if (semanticOperationFamily(candidate) !== wantedFamily) return;
      if (explicitSectionGroupHeader(candidate)) return;

      const candidateMetric = sectionAuthoritativeMetricHeader(candidate);
      if (!candidateMetric) return;
      const score = semanticMetricCompatibilityScore({
        requestedLabel: label,
        candidateMetric,
      });
      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });
  }

  return bestScore >= 10 ? bestIndex : -1;
}

function mixedSectionSemanticRoles(section = {}) {
  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  return new Set(
    rows
      .map((row) => semanticRowMetricLabel(row))
      .filter(Boolean)
      .map((label) => semanticRoleForMetricLabel(label).role)
      .filter(Boolean),
  );
}

function isMixedMetricRowSection(section = {}) {
  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  if (rows.length < 2) return false;
  const roles = mixedSectionSemanticRoles(section);
  return roles.size > 1;
}

function replaceMixedSectionRows({
  section = {},
  sections = [],
  excludedIndex = -1,
} = {}) {
  if (isContractMetricSection(section)) {
    return {
      section,
      replacedRowCount: 0,
      replacements: [],
    };
  }

  if (!isMixedMetricRowSection(section)) {
    return {
      section,
      replacedRowCount: 0,
      replacements: [],
    };
  }

  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  const replacements = [];
  const nextRows = rows.map((row, rowIndex) => {
    if (
      !row ||
      typeof row !== "object" ||
      Array.isArray(row)
    ) {
      return row;
    }

    const label = semanticRowMetricLabel(row);
    const role = semanticRoleForMetricLabel(label);
    const family = role.role === SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT
      ? "latest"
      : [
          SEMANTIC_METRIC_ROLE.UNIT_RATE,
          SEMANTIC_METRIC_ROLE.DURATION,
          SEMANTIC_METRIC_ROLE.PERCENTAGE_RATE,
        ].includes(role.role)
        ? "average"
        : "";
    if (!family) return row;

    const targetIndex = authoritativeSemanticSectionIndex({
      sections,
      requestedLabel: label || explicitSectionMetricHeader(section),
      family,
      groupHeader: "",
      excludedIndex,
    });
    if (targetIndex < 0) return row;

    const target = sections[targetIndex];
    const scalar = sectionAuthoritativeScalar(target, family);
    if (!Number.isFinite(scalar.value)) return row;

    const valueInfo = semanticRowValue(row);
    const valueKey = valueInfo.key || "값";
    const metric = sectionAuthoritativeMetricHeader(target);
    const labelKey = [
      "지표",
      "metric",
      "지표명",
      "항목",
      "label",
      "name",
    ].find((key) => Object.prototype.hasOwnProperty.call(row, key)) || "지표";
    const suffix = family === "latest" ? "최신 스냅샷" : "평균";
    const nextRow = {
      ...row,
      [labelKey]: `${metric} ${suffix}`,
      [valueKey]: scalar.value,
    };

    replacements.push({
      rowIndex,
      originalLabel: label,
      replacementLabel: nextRow[labelKey],
      family,
      authoritativeSectionId: normalizeText(target.sectionId || ""),
      authoritativeMetric: metric,
      originalValue: Number.isFinite(valueInfo.value) ? valueInfo.value : null,
      replacementValue: scalar.value,
    });
    return nextRow;
  });

  if (!replacements.length) {
    return { section, replacedRowCount: 0, replacements: [] };
  }

  return {
    section: {
      ...section,
      result: {
        ...(section.result || {}),
        rows: nextRows,
        meta: {
          ...(section.result?.meta || {}),
          mixedSectionRowPrecedenceVersion:
            MIXED_SECTION_ROW_PRECEDENCE_VERSION,
          mixedSectionRowPrecedenceApplied: true,
          replacedSemanticConflictRowCount: replacements.length,
          semanticConflictRowReplacements: cloneValue(replacements),
        },
      },
    },
    replacedRowCount: replacements.length,
    replacements,
  };
}

function contractSectionOperation(section = {}) {
  return normalizeText(
    `${section.sectionType || ""} ${section.result?.operation || ""}`,
  ).toLowerCase();
}

function isContractMetricSection(section = {}) {
  return /contract/.test(contractSectionOperation(section));
}

function contractMetricOutputHeader(section = {}) {
  const groupHeader = explicitSectionGroupHeader(section);
  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  const row = rows.find((item) =>
    item && typeof item === "object" && !Array.isArray(item),
  ) || {};
  const ignored = new Set([
    groupHeader,
    "metricId",
    "순위",
    "rank",
    "작업",
    "operation",
  ].filter(Boolean));
  return Object.keys(row).find((key) =>
    !ignored.has(key) && Number.isFinite(Number(row[key])),
  ) || sectionAuthoritativeMetricHeader(section) || normalizeText(section.title);
}

function candidateGroupedRows(section = {}) {
  const groupHeader = explicitSectionGroupHeader(section);
  const rows = Array.isArray(section.result?.rows)
    ? section.result.rows
    : [];
  return rows
    .map((row) => {
      const groupValue = row?.[groupHeader];
      const numeric = semanticRowValue(row || {});
      return {
        groupValue,
        value: numeric.value,
      };
    })
    .filter((item) =>
      item.groupValue != null && Number.isFinite(item.value),
    );
}

function contractRowAggregationIntent(row = {}, section = {}) {
  const label = semanticRowMetricLabel(row);
  const evidence = normalizeText([
    label,
    row.집계유형,
    row.집계방식,
    row.aggregation,
    row.aggregate,
    row.operation,
    row.작업,
    row.metricId,
    section.result?.aggregation,
    section.result?.operation,
    section.operation,
    section.sectionType,
  ].filter(Boolean).join(" ")).toLowerCase();

  if (/평균|average|mean|avg/.test(evidence)) return "average";
  if (/총|합계|합산|총계|sum|total|grand\s*total/.test(evidence)) {
    return "total";
  }
  return "latest_total";
}

function connectContractKpisToSemanticSnapshots({ sections = [] } = {}) {
  const working = Array.isArray(sections)
    ? sections.map((section) => cloneValue(section))
    : [];
  const updates = [];

  working.forEach((section, sectionIndex) => {
    if (!isContractMetricSection(section)) return;

    const groupHeader = explicitSectionGroupHeader(section);
    const outputHeader = contractMetricOutputHeader(section);
    const sectionLabel = normalizeText(
      sectionAuthoritativeMetricHeader(section) ||
        outputHeader ||
        section.title,
    );
    const sectionRole = semanticRoleForMetricLabel(sectionLabel);

    if (groupHeader && sectionRole.role === SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT) {
      const targetIndex = authoritativeSemanticSectionIndex({
        sections: working,
        requestedLabel: sectionLabel,
        family: "latest",
        groupHeader,
        excludedIndex: sectionIndex,
      });
      if (targetIndex < 0) return;
      const target = working[targetIndex];
      const sourceRows = candidateGroupedRows(target);
      if (!sourceRows.length) return;
      const metricId = collectSectionMetricIds(section)[0] || "";
      const isRank = /rank|순위|top/i.test(contractSectionOperation(section));
      let nextRows;

      if (isRank) {
        const limit = Math.max(1, (section.result?.rows || []).length || 5);
        let rank = 0;
        let lastValue = null;
        nextRows = sourceRows
          .slice()
          .sort((a, b) => b.value - a.value)
          .slice(0, limit)
          .map((item, index) => {
            if (lastValue === null || item.value !== lastValue) {
              rank = index + 1;
              lastValue = item.value;
            }
            return {
              [groupHeader]: item.groupValue,
              [outputHeader]: item.value,
              순위: rank,
              ...(metricId ? { metricId } : {}),
            };
          });
      } else {
        nextRows = sourceRows.map((item) => ({
          [groupHeader]: item.groupValue,
          [outputHeader]: item.value,
          ...(metricId ? { metricId } : {}),
        }));
      }

      working[sectionIndex] = {
        ...section,
        result: {
          ...(section.result || {}),
          rows: nextRows,
          meta: {
            ...(section.result?.meta || {}),
            contractKpiSnapshotBridgeVersion:
              CONTRACT_KPI_SNAPSHOT_BRIDGE_VERSION,
            contractKpiSnapshotBridgeApplied: true,
            authoritativeSemanticSectionId: normalizeText(target.sectionId || ""),
            authoritativeSemanticMetric: sectionAuthoritativeMetricHeader(target),
          },
        },
      };
      updates.push({
        sectionId: normalizeText(section.sectionId || ""),
        mode: isRank ? "group_rank_latest" : "group_latest",
        rowCount: nextRows.length,
        authoritativeSectionId: normalizeText(target.sectionId || ""),
      });
      return;
    }

    if (groupHeader) return;
    const rows = Array.isArray(section.result?.rows)
      ? section.result.rows
      : [];
    let changed = false;
    const rowUpdates = [];
    const nextRows = rows.map((row, rowIndex) => {
      if (!row || typeof row !== "object" || Array.isArray(row)) return row;
      const label = semanticRowMetricLabel(row);
      const role = semanticRoleForMetricLabel(label);
      if (role.role !== SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT) return row;

      const targetIndex = authoritativeSemanticSectionIndex({
        sections: working,
        requestedLabel: label,
        family: "latest",
        groupHeader: "",
        excludedIndex: sectionIndex,
      });
      if (targetIndex < 0) return row;
      const target = working[targetIndex];
      const scalar = sectionAuthoritativeScalar(target, "latest");
      if (!Number.isFinite(scalar.value)) return row;

      const aggregationIntent = contractRowAggregationIntent(row, section);
      const isAverage = aggregationIntent === "average";
      const selectedCount = sectionAuthoritativeSelectedCount(target);
      if (isAverage && selectedCount <= 0) return row;
      const value = isAverage
        ? scalar.value / selectedCount
        : scalar.value;
      const valueInfo = semanticRowValue(row);
      const valueKey = valueInfo.key || "값";
      changed = true;
      rowUpdates.push({
        rowIndex,
        label,
        originalValue: Number.isFinite(valueInfo.value) ? valueInfo.value : null,
        replacementValue: value,
        mode: isAverage ? "snapshot_average" : "snapshot_total",
        aggregationIntent,
        selectedCount: isAverage ? selectedCount : null,
        authoritativeSectionId: normalizeText(target.sectionId || ""),
      });
      return {
        ...row,
        [valueKey]: value,
      };
    });

    if (!changed) return;
    working[sectionIndex] = {
      ...section,
      result: {
        ...(section.result || {}),
        rows: nextRows,
        meta: {
          ...(section.result?.meta || {}),
          contractKpiSnapshotBridgeVersion:
            CONTRACT_KPI_SNAPSHOT_BRIDGE_VERSION,
          contractKpiSnapshotBridgeApplied: true,
          contractKpiSnapshotRowUpdates: cloneValue(rowUpdates),
        },
      },
    };
    updates.push({
      sectionId: normalizeText(section.sectionId || ""),
      mode: "scalar_rows",
      rowCount: rowUpdates.length,
      rowUpdates,
    });
  });

  return {
    sections: working,
    applied: updates.length > 0,
    updatedSectionCount: updates.length,
    updatedRowCount: updates.reduce(
      (sum, item) => sum + Number(item.rowCount || 0),
      0,
    ),
    updates,
    version: CONTRACT_KPI_SNAPSHOT_BRIDGE_VERSION,
  };
}


function semanticConflictReason({
  existing = {},
  planned = {},
} = {}) {
  if (
    !sectionMetricAndGroupMatch(
      existing,
      planned,
    )
  ) {
    return "";
  }

  const plannedFamily =
    semanticOperationFamily(planned);
  if (
    !["latest", "average"].includes(
      plannedFamily,
    )
  ) {
    return "";
  }

  const existingFamilies =
    sectionOperationFamilies(existing);

  if (
    plannedFamily === "latest"
  ) {
    if (existingFamilies.has("latest")) {
      return "";
    }
    if (
      existingFamilies.has("sum") ||
      existingFamilies.has("summary")
    ) {
      return "LATEST_OVERRIDES_LEGACY_SUM";
    }
  }

  if (
    plannedFamily === "average"
  ) {
    if (existingFamilies.has("average")) {
      return "";
    }
    if (
      existingFamilies.has("sum") ||
      existingFamilies.has("summary")
    ) {
      return "AVERAGE_OVERRIDES_LEGACY_SUM";
    }
  }

  return "";
}

function authoritativeSectionIndex({
  sections = [],
  planned = {},
  excludedIndex = -1,
} = {}) {
  const plannedId = normalizeText(
    planned.sectionId || "",
  );

  const exactIndex = sections.findIndex(
    (section, index) =>
      index !== excludedIndex &&
      plannedId &&
      normalizeText(
        section.sectionId || "",
      ) === plannedId,
  );
  if (exactIndex >= 0) return exactIndex;

  const annotatedIndex = sections.findIndex(
    (section, index) =>
      index !== excludedIndex &&
      normalizeText(
        section.result?.meta
          ?.semanticCoverage
          ?.plannerSectionId || "",
      ) === plannedId,
  );
  if (annotatedIndex >= 0) {
    return annotatedIndex;
  }

  return sections.findIndex(
    (section, index) =>
      index !== excludedIndex &&
      existingSectionCoversPlanned(
        section,
        planned,
      ),
  );
}

function resolveSemanticContractConflicts({
  sections = [],
  plannedSections = [],
  baseSectionCount = 0,
  actualFlowEvidence = null,
} = {}) {
  const working = Array.isArray(sections)
    ? sections.map((section) =>
        cloneValue(section),
      )
    : [];
  const plans = Array.isArray(
    plannedSections,
  )
    ? plannedSections
    : [];
  const baseLimit = Math.max(
    0,
    Math.min(
      Number(baseSectionCount) || 0,
      working.length,
    ),
  );
  const removedIndexes = new Set();
  const conflicts = [];
  const mixedRowRepairs = [];
  const blockedMixedFlowSections = [];
  let transferredMetricIdCount = 0;
  let replacedMixedRowCount = 0;

  for (
    let existingIndex = 0;
    existingIndex < baseLimit;
    existingIndex += 1
  ) {
    const existing = working[existingIndex];
    if (!existing) continue;

    const flowGateEvaluated = actualFlowEvidence?.evaluated === true;
    const blockedByFlowGate =
      flowGateEvaluated &&
      isInventoryFlowMixedSection(existing) &&
      actualFlowEvidence?.pass !== true;

    if (blockedByFlowGate) {
      const legacyMetricIds = collectSectionMetricIds(existing);
      const targetIndex = blockedFlowMetricTransferTargetIndex({
        sections: working,
        section: existing,
        excludedIndex: existingIndex,
      });
      let transferredMetricIds = [];
      let authoritativeSectionId = "";
      let authoritativeTitle = "";

      if (targetIndex >= 0) {
        const targetMetricIds = collectSectionMetricIds(working[targetIndex]);
        const mergedMetricIds = uniqueMetricIds([
          targetMetricIds,
          legacyMetricIds,
        ]);
        transferredMetricIds = legacyMetricIds.filter(
          (metricId) => !targetMetricIds.includes(metricId),
        );
        transferredMetricIdCount += transferredMetricIds.length;
        working[targetIndex] = applySectionMetricIds(
          working[targetIndex],
          mergedMetricIds,
        );
        authoritativeSectionId = normalizeText(
          working[targetIndex]?.sectionId || "",
        );
        authoritativeTitle = normalizeText(
          working[targetIndex]?.title || "",
        );
      }

      removedIndexes.add(existingIndex);
      const blocked = {
        reason: "MIXED_FLOW_SECTION_BLOCKED_NO_ACTUAL_FLOW_EVIDENCE",
        legacySectionId: normalizeText(existing.sectionId || ""),
        legacyTitle: normalizeText(existing.title || ""),
        authoritativeSectionId,
        authoritativeTitle,
        transferredMetricIds,
        flowEvidenceMode:
          actualFlowEvidence?.mode || "no_actual_flow_evidence",
      };
      blockedMixedFlowSections.push(blocked);
      conflicts.push({
        ...blocked,
        metricHeader: "",
        groupHeader: "",
        rowLevel: false,
      });
      continue;
    }

    const mixedRepair = replaceMixedSectionRows({
      section: existing,
      sections: working,
      excludedIndex: existingIndex,
    });
    if (mixedRepair.replacedRowCount > 0) {
      working[existingIndex] = mixedRepair.section;
      replacedMixedRowCount += mixedRepair.replacedRowCount;
      mixedRowRepairs.push({
        legacySectionId: normalizeText(existing.sectionId || ""),
        legacyTitle: normalizeText(existing.title || ""),
        replacedRowCount: mixedRepair.replacedRowCount,
        replacements: cloneValue(mixedRepair.replacements),
      });
      conflicts.push({
        reason: "ROW_LEVEL_SEMANTIC_CONTRACT_PRECEDENCE",
        metricHeader: "",
        groupHeader: "",
        legacySectionId: normalizeText(existing.sectionId || ""),
        legacyTitle: normalizeText(existing.title || ""),
        authoritativeSectionId: "",
        authoritativeTitle: "",
        transferredMetricIds: [],
        rowLevel: true,
        replacements: cloneValue(mixedRepair.replacements),
      });
      continue;
    }

    for (const planned of plans) {
      const reason = semanticConflictReason({
        existing,
        planned,
      });
      if (!reason) continue;

      const targetIndex = authoritativeSectionIndex({
        sections: working,
        planned,
        excludedIndex: existingIndex,
      });
      if (targetIndex < 0) continue;

      const legacyMetricIds = collectSectionMetricIds(existing);
      const targetMetricIds = collectSectionMetricIds(working[targetIndex]);
      const mergedMetricIds = uniqueMetricIds([
        targetMetricIds,
        legacyMetricIds,
      ]);

      transferredMetricIdCount += mergedMetricIds.filter(
        (metricId) => !targetMetricIds.includes(metricId),
      ).length;

      working[targetIndex] = applySectionMetricIds(
        working[targetIndex],
        mergedMetricIds,
      );

      removedIndexes.add(existingIndex);
      conflicts.push({
        reason,
        metricHeader: normalizeText(
          planned.result?.metric?.header || "",
        ),
        groupHeader: normalizeText(
          planned.result?.groupBy?.header || "",
        ),
        legacySectionId: normalizeText(existing.sectionId || ""),
        legacyTitle: normalizeText(existing.title || ""),
        authoritativeSectionId: normalizeText(
          working[targetIndex]?.sectionId || "",
        ),
        authoritativeTitle: normalizeText(
          working[targetIndex]?.title || "",
        ),
        transferredMetricIds: legacyMetricIds,
        rowLevel: false,
      });
      break;
    }

  }

  const retained = working.filter(
    (_, index) => !removedIndexes.has(index),
  );
  const prunedConflicts = conflicts.filter(
    (conflict) => conflict.rowLevel !== true,
  );

  return {
    sections: retained,
    applied: conflicts.length > 0,
    prunedSectionCount: removedIndexes.size,
    prunedSectionIds: prunedConflicts.map(
      (conflict) =>
        conflict.legacySectionId ||
        conflict.legacyTitle,
    ),
    mixedSectionRowPrecedenceVersion:
      MIXED_SECTION_ROW_PRECEDENCE_VERSION,
    mixedSectionRowPrecedenceApplied:
      mixedRowRepairs.length > 0,
    repairedMixedSectionCount:
      mixedRowRepairs.length,
    replacedMixedRowCount,
    mixedRowRepairs,
    actualFlowEvidenceGateVersion: ACTUAL_FLOW_EVIDENCE_GATE_VERSION,
    actualFlowEvidence: cloneValue(actualFlowEvidence || {}),
    blockedMixedFlowSectionCount: blockedMixedFlowSections.length,
    blockedMixedFlowSectionIds: blockedMixedFlowSections.map(
      (item) => item.legacySectionId || item.legacyTitle,
    ),
    blockedMixedFlowSections,
    transferredMetricIdCount,
    conflicts,
    version: SEMANTIC_CONTRACT_PRECEDENCE_VERSION,
  };
}

const GENERIC_SECTION_CLEANUP_VERSION =
  "semantic_generic_section_cleanup_v1";

function plannedSpecificMetricHeaders(
  plannedSections = [],
) {
  return Array.from(
    new Set(
      plannedSections
        .map((section) =>
          normalizeText(
            section.result?.metric?.header || "",
          ),
        )
        .filter(
          (header) =>
            header &&
            !GENERIC_METRIC_LABEL_PATTERN.test(header),
        ),
    ),
  );
}

function sectionHasGenericMetricIdentity(
  section = {},
) {
  const explicitMetric =
    explicitSectionMetricHeader(section);

  if (explicitMetric) {
    return GENERIC_METRIC_LABEL_PATTERN.test(
      explicitMetric,
    );
  }

  const title = normalizeText(section.title);
  return /지표값|metric\s*value|measure\s*value/i.test(
    title,
  );
}

function plannedMetricHasReplacementCoverage({
  plannedSections = [],
  metricHeader = "",
} = {}) {
  const metricKey = normalizeKey(metricHeader);
  const relevant = plannedSections.filter(
    (section) =>
      normalizeKey(
        section.result?.metric?.header || "",
      ) === metricKey,
  );

  if (!relevant.length) return false;

  const hasSummary = relevant.some((section) => {
    const group = normalizeText(
      section.result?.groupBy?.header || "",
    );
    const family = semanticOperationFamily(section);
    return (
      !group &&
      ["summary", "sum", "average"].includes(
        family,
      )
    );
  });

  const hasGroupedOrPeriod = relevant.some(
    (section) =>
      normalizeText(
        section.result?.groupBy?.header || "",
      ),
  );

  return hasSummary && hasGroupedOrPeriod;
}

function pruneSupersededGenericSections({
  sections = [],
  plannedSections = [],
  baseSectionCount = 0,
} = {}) {
  const specificMetrics =
    plannedSpecificMetricHeaders(plannedSections);

  if (specificMetrics.length !== 1) {
    return {
      sections,
      prunedSectionCount: 0,
      prunedSectionIds: [],
      cleanupApplied: false,
      reason:
        "SPECIFIC_METRIC_COUNT_NOT_ONE",
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
      reason:
        "SPECIFIC_METRIC_COVERAGE_INCOMPLETE",
    };
  }

  const prunedSectionIds = [];

  const retained = sections.filter(
    (section, index) => {
      if (index >= baseSectionCount) return true;

      if (
        collectSectionMetricIds(section).length > 0
      ) {
        return true;
      }

      if (
        !sectionHasGenericMetricIdentity(section)
      ) {
        return true;
      }

      prunedSectionIds.push(
        normalizeText(
          section.sectionId ||
            section.title ||
            `base_section_${index + 1}`,
        ),
      );
      return false;
    },
  );

  return {
    sections: retained,
    prunedSectionCount:
      prunedSectionIds.length,
    prunedSectionIds,
    cleanupApplied:
      prunedSectionIds.length > 0,
    reason:
      prunedSectionIds.length > 0
        ? ""
        : "NO_UNCONTRACTED_GENERIC_SECTION",
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
  const preferredTables = selectPreferredSemanticTables(
    tables,
    { preferDimensionCompletePhysical: true },
  );
  const plannerContext = {
    ...cloneValue(context),
    templateId:
      executionResult.templateId || templateCandidate.templateId || "",
    templateTitle:
      executionResult.title || templateCandidate.title || "",
    templateDescription:
      executionResult.description || templateCandidate.description || "",
  };

  const plan = buildSemanticOutputPlan({
    tables: preferredTables,
    templateId:
      executionResult.templateId || templateCandidate.templateId ||
      "business_template",
    title:
      executionResult.title || templateCandidate.title ||
      "업무 템플릿",
    context: plannerContext,
    options: {
      includeOverview: false,
      includeDiagnostics: false,
      compactTitles: true,
      maxDimensionsPerSeries:
        options.maxDimensionsPerSeries ?? 8,
      maxSections: options.maxPlannedSections ?? 120,
      maxSectionsPerTable: options.maxSectionsPerTable ?? 28,
      includeDistinct: options.includeDistinct !== false,
      preferVirtualTables: false,
    },
  });

  const plannedSections = (plan.sections || []).filter(
    (section) =>
      !String(section.sectionType || "").startsWith(
        "semantic_table_overview",
      ) &&
      !String(section.sectionType || "").startsWith(
        "semanticSource",
      ),
  );

  const merged = [...baseSections];
  const renderedPlannerIds = [];
  let matchedExistingSectionCount = 0;
  let addedSectionCount = 0;
  const maxAddedSections = Math.max(
    0,
    Number(options.maxAddedSections ?? 64),
  );

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

    const mandatorySummary =
      isWholeMetricSummarySection(planned);
    if (
      addedSectionCount >= maxAddedSections &&
      !mandatorySummary
    ) {
      continue;
    }

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

  const actualFlowEvidence = detectActualFlowEvidence(preferredTables);
  const contractPrecedence =
    resolveSemanticContractConflicts({
      sections: merged,
      plannedSections,
      baseSectionCount: baseSections.length,
      actualFlowEvidence,
    });

  const genericCleanup =
    pruneSupersededGenericSections({
      sections: contractPrecedence.sections,
      plannedSections,
      baseSectionCount: Math.max(
        0,
        baseSections.length -
          contractPrecedence.prunedSectionCount,
      ),
    });

  const contractSnapshotBridge =
    connectContractKpisToSemanticSnapshots({
      sections: genericCleanup.sections,
    });

  const flowDirectionSemantics =
    applyFlowDirectionSemantics({
      sections: contractSnapshotBridge.sections,
      tables: preferredTables,
    });

  const representativePriority = prioritizeBusinessSections({
    sections: flowDirectionSemantics.sections,
    primaryMetricLabels:
      plan.executionMeta?.primaryMetricLabels || [],
    componentMetricLabels:
      plan.executionMeta?.componentMetricLabels || [],
  });

  const mandatorySummaryMetricIds = uniqueMetricIds(
    plannedSections
      .filter(isWholeMetricSummarySection)
      .flatMap((section) => collectSectionMetricIds(section)),
  );
  const expectedMetricIds = uniqueMetricIds([
    renderedPlannerIds,
    mandatorySummaryMetricIds,
  ]);
  const normalizedBeforeCoverageFloor =
    normalizeSectionMetricIds(
      representativePriority.sections,
    );
  const mandatorySummaryCoverageFloor =
    applyMandatorySummaryCoverageFloor({
      sections: normalizedBeforeCoverageFloor,
      plannedSections,
      expectedMetricIds,
    });
  const finalOutputQualityGate =
    applyFinalOutputQualityGate({
      sections: mandatorySummaryCoverageFloor.sections,
      expectedMetricIds,
      sectionBudgetSummaries:
        plan.executionMeta?.sectionBudgetSummaries || [],
      throwOnFailure:
        options.enforceFinalOutputQualityGate !== false,
    });
  const normalizedSections =
    normalizeSectionMetricIds(
      finalOutputQualityGate.sections,
    );

  return {
    ...cloneValue(executionResult),
    sections: normalizedSections,
    contractSummaryCoverage: {
      version: `${SEMANTIC_OUTPUT_CONTRACT_VERSION}_business_coverage_v1`,
      contractCatalogVersion:
        `${SEMANTIC_OUTPUT_CONTRACT_VERSION}_business_coverage_catalog_v1`,
      expectedMetricIds,
    },
    finalOutputQualityGate: {
      version: FINAL_OUTPUT_QUALITY_GATE_VERSION,
      completenessVersion: OUTPUT_COMPLETENESS_CONTRACT_VERSION,
      duplicateResolverVersion:
        SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
      status: finalOutputQualityGate.status,
      pass: finalOutputQualityGate.pass,
      failureReasons: cloneValue(
        finalOutputQualityGate.failureReasons || [],
      ),
      expectedMetricCount:
        finalOutputQualityGate.analysis?.expectedMetricCount || 0,
      renderedExpectedMetricCount:
        finalOutputQualityGate.analysis?.renderedExpectedMetricCount || 0,
    },
    executionMeta: {
      ...(executionResult.executionMeta || {}),
      semanticBusinessAugmentation: true,
      semanticOutputPlannerVersion:
        SEMANTIC_OUTPUT_PLANNER_VERSION,
      metricSemanticRoleEngineVersion:
        plan.executionMeta?.metricSemanticRoleEngineVersion ||
        METRIC_SEMANTIC_ROLE_ENGINE_VERSION,
      aggregationContractResolverVersion:
        plan.executionMeta?.aggregationContractResolverVersion ||
        AGGREGATION_CONTRACT_RESOLVER_VERSION,
      metricRelationshipPriorityEngineVersion:
        METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
      derivedTotalRelationVersion:
        DERIVED_TOTAL_RELATION_VERSION,
      representativeMetricPriorityVersion:
        REPRESENTATIVE_METRIC_PRIORITY_VERSION,
      metricRelationshipCount:
        Number(plan.executionMeta?.metricRelationshipCount || 0),
      metricRelationships:
        cloneValue(plan.executionMeta?.metricRelationships || []),
      primaryMetricLabels:
        cloneValue(plan.executionMeta?.primaryMetricLabels || []),
      componentMetricLabels:
        cloneValue(plan.executionMeta?.componentMetricLabels || []),
      representativeMetricPriorityApplied:
        representativePriority.applied,
      representativeMetricReorderedSectionCount:
        representativePriority.reorderedSectionCount,
      semanticSectionBudgetEngineVersion:
        SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
      mandatorySummaryCoverageFloorVersion:
        MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
      mandatorySummaryCoverageFloorApplied:
        mandatorySummaryCoverageFloor.applied,
      mandatorySummaryCoverageFloorPass:
        mandatorySummaryCoverageFloor.pass,
      mandatorySummaryCoverageMissingMetricIdsBefore:
        cloneValue(
          mandatorySummaryCoverageFloor.missingMetricIdsBefore,
        ),
      mandatorySummaryCoverageMissingMetricIdsAfter:
        cloneValue(
          mandatorySummaryCoverageFloor.missingMetricIdsAfter,
        ),
      mandatorySummaryCoverageRestoredMetricIds:
        cloneValue(
          mandatorySummaryCoverageFloor.restoredMetricIds,
        ),
      mandatorySummaryCoverageTransferredMetricIds:
        cloneValue(
          mandatorySummaryCoverageFloor.transferredMetricIds,
        ),
      mandatorySummaryCoverageRestoredSectionIds:
        cloneValue(
          mandatorySummaryCoverageFloor.restoredSectionIds,
        ),
      mandatorySummaryCoverageActions:
        cloneValue(
          mandatorySummaryCoverageFloor.coverageActions,
        ),
      outputCompletenessContractVersion:
        OUTPUT_COMPLETENESS_CONTRACT_VERSION,
      semanticOutputDuplicateResolverVersion:
        SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
      finalOutputQualityGateVersion:
        FINAL_OUTPUT_QUALITY_GATE_VERSION,
      finalOutputQualityGateApplied:
        finalOutputQualityGate.applied,
      finalOutputQualityGatePass:
        finalOutputQualityGate.pass,
      finalOutputQualityGateStatus:
        finalOutputQualityGate.status,
      finalOutputQualityGateStrict:
        options.enforceFinalOutputQualityGate !== false,
      finalOutputQualityGateFailureReasons:
        cloneValue(finalOutputQualityGate.failureReasons || []),
      finalOutputQualityGateExpectedMetricCount:
        Number(finalOutputQualityGate.analysis?.expectedMetricCount || 0),
      finalOutputQualityGateRenderedExpectedMetricCount:
        Number(
          finalOutputQualityGate.analysis?.renderedExpectedMetricCount || 0,
        ),
      finalOutputQualityGateMissingMetricIds:
        cloneValue(finalOutputQualityGate.analysis?.missingMetricIds || []),
      finalOutputQualityGateDuplicateMetricIds:
        cloneValue(finalOutputQualityGate.analysis?.duplicateMetricIds || []),
      finalOutputQualityGateEmptyRequiredSectionIds:
        cloneValue(
          finalOutputQualityGate.analysis?.emptyRequiredSectionIds || [],
        ),
      finalOutputQualityGateIncompleteRequiredSectionIds:
        cloneValue(
          finalOutputQualityGate.analysis?.incompleteRequiredSectionIds || [],
        ),
      finalOutputQualityGateInvalidNumberCount:
        Number(finalOutputQualityGate.analysis?.invalidNumbers?.length || 0),
      finalOutputQualityGateBudgetViolationCount:
        Number(
          finalOutputQualityGate.analysis?.sectionBudgetViolations?.length || 0,
        ),
      finalOutputQualityGateInputSectionCount:
        Number(
          finalOutputQualityGate.duplicateResolution?.inputSectionCount || 0,
        ),
      finalOutputQualityGateOutputSectionCount:
        Number(finalOutputQualityGate.sections?.length || 0),
      finalOutputQualityGateRemovedDuplicateSectionCount:
        Number(
          finalOutputQualityGate.duplicateResolution
            ?.removedDuplicateSectionCount || 0,
        ),
      finalOutputQualityGateRemovedDuplicateSectionIds:
        cloneValue(
          finalOutputQualityGate.duplicateResolution
            ?.removedDuplicateSectionIds || [],
        ),
      finalOutputQualityGateMergedMetricIdCount:
        Number(
          finalOutputQualityGate.duplicateResolution
            ?.mergedMetricIdCount || 0,
        ),
      finalOutputQualityGateReassignedMetricIdCount:
        Number(
          finalOutputQualityGate.metricOwnership
            ?.removedOwnershipCount || 0,
        ),
      finalOutputQualityGateRenamedSectionIdCount:
        Number(finalOutputQualityGate.renamedSectionIds?.length || 0),
      finalOutputQualityGateRenamedTitleCount:
        Number(finalOutputQualityGate.renamedTitles?.length || 0),
      durationSummaryContractVersion:
        DURATION_SUMMARY_CONTRACT_VERSION,
      distinctEntitySectionVersion:
        DISTINCT_ENTITY_SECTION_VERSION,
      sectionBudgetSummaries:
        cloneValue(plan.executionMeta?.sectionBudgetSummaries || []),
      semanticMetricRoleCounts:
        cloneValue(plan.executionMeta?.semanticMetricRoleCounts || {}),
      semanticAggregationOperationCounts:
        cloneValue(
          plan.executionMeta?.semanticAggregationOperationCounts || {},
        ),
      unsafeAggregationOverrideCount:
        Number(plan.executionMeta?.unsafeAggregationOverrideCount || 0),
      snapshotEntityResolverVersion:
        plan.executionMeta?.snapshotEntityResolverVersion ||
        SNAPSHOT_ENTITY_RESOLVER_VERSION,
      snapshotEntitySeriesCount:
        Number(plan.executionMeta?.snapshotEntitySeriesCount || 0),
      snapshotEntityAppliedSeriesCount:
        Number(plan.executionMeta?.snapshotEntityAppliedSeriesCount || 0),
      snapshotEntitySelections:
        cloneValue(plan.executionMeta?.snapshotEntitySelections || []),
      flowDirectionSemanticEngineVersion:
        FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
      flowDirectionSectionRepairVersion:
        FLOW_DIRECTION_SECTION_REPAIR_VERSION,
      flowDirectionApplied:
        flowDirectionSemantics.applied,
      flowDirectionReason:
        flowDirectionSemantics.reason,
      flowDirectionEvidence:
        cloneValue(flowDirectionSemantics.evidence || {}),
      flowDirectionSystemSummary:
        cloneValue(flowDirectionSemantics.systemSummary || {}),
      flowDirectionRepairedSectionCount:
        Number(flowDirectionSemantics.repairedSectionCount || 0),
      flowDirectionRepairedSectionIds:
        cloneValue(flowDirectionSemantics.repairedSectionIds || []),
      flowDirectionScopes:
        cloneValue(flowDirectionSemantics.scopes || []),
      flowDirectionDualEntryApplied:
        flowDirectionSemantics.dualEntryApplied === true,
      flowDirectionLocationEntryCount:
        Number(flowDirectionSemantics.locationEntryCount || 0),
      semanticContractPrecedenceVersion:
        SEMANTIC_CONTRACT_PRECEDENCE_VERSION,
      semanticContractPrecedenceApplied:
        contractPrecedence.applied,
      prunedLegacyConflictCount:
        contractPrecedence.prunedSectionCount,
      prunedLegacyConflictSectionIds:
        contractPrecedence.prunedSectionIds,
      transferredLegacyMetricIdCount:
        contractPrecedence.transferredMetricIdCount,
      semanticContractConflicts:
        cloneValue(contractPrecedence.conflicts),
      actualFlowEvidenceGateVersion:
        ACTUAL_FLOW_EVIDENCE_GATE_VERSION,
      actualFlowEvidencePass:
        actualFlowEvidence.pass,
      actualFlowEvidenceMode:
        actualFlowEvidence.mode,
      actualFlowEvidence:
        cloneValue(actualFlowEvidence),
      blockedMixedFlowSectionCount:
        contractPrecedence.blockedMixedFlowSectionCount,
      blockedMixedFlowSectionIds:
        contractPrecedence.blockedMixedFlowSectionIds,
      blockedMixedFlowSections:
        cloneValue(contractPrecedence.blockedMixedFlowSections),
      mixedSectionRowPrecedenceVersion:
        MIXED_SECTION_ROW_PRECEDENCE_VERSION,
      generalStockSnapshotAliasVersion:
        GENERAL_STOCK_SNAPSHOT_ALIAS_VERSION,
      mixedSectionRowPrecedenceApplied:
        contractPrecedence.mixedSectionRowPrecedenceApplied,
      repairedMixedSectionCount:
        contractPrecedence.repairedMixedSectionCount,
      replacedMixedRowCount:
        contractPrecedence.replacedMixedRowCount,
      mixedRowRepairs:
        cloneValue(contractPrecedence.mixedRowRepairs),
      contractKpiSnapshotBridgeVersion:
        CONTRACT_KPI_SNAPSHOT_BRIDGE_VERSION,
      contractKpiSnapshotBridgeApplied:
        contractSnapshotBridge.applied,
      contractKpiSnapshotUpdatedSectionCount:
        contractSnapshotBridge.updatedSectionCount,
      contractKpiSnapshotUpdatedRowCount:
        contractSnapshotBridge.updatedRowCount,
      contractKpiSnapshotUpdates:
        cloneValue(contractSnapshotBridge.updates),
      preferredTableCount: preferredTables.length,
      plannedSectionCount: plannedSections.length,
      matchedExistingSectionCount,
      addedSectionCount,
      expectedMetricIdCount: expectedMetricIds.length,
      maxAddedSections,
      genericSectionCleanupVersion:
        GENERIC_SECTION_CLEANUP_VERSION,
      genericSectionCleanupApplied:
        genericCleanup.cleanupApplied,
      prunedGenericSectionCount:
        genericCleanup.prunedSectionCount,
      prunedGenericSectionIds:
        genericCleanup.prunedSectionIds,
      genericSectionCleanupReason:
        genericCleanup.reason,
      genericSectionCleanupMetric:
        genericCleanup.specificMetric || "",
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
    const summary = Array.isArray(table.summaryRows)
      ? table.summaryRows
      : [];
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
      역할: (plan?.series || [])
        .map((series) => series.metricRole)
        .filter(Boolean)
        .join(", "),
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

function semanticRoleSummary(plans = []) {
  const roleCounts = {};
  const operationCounts = {};
  let unsafeAggregationOverrideCount = 0;

  for (const plan of Array.isArray(plans) ? plans : []) {
    for (const series of Array.isArray(plan.series) ? plan.series : []) {
      const role = normalizeText(series.metricRole || "generic_measure");
      const operation = normalizeText(series.operation || "sum");
      roleCounts[role] = (roleCounts[role] || 0) + 1;
      operationCounts[operation] = (operationCounts[operation] || 0) + 1;
      if (
        series.aggregationContract?.unsafeDeclaredAggregationOverridden
      ) {
        unsafeAggregationOverrideCount += 1;
      }
    }
  }

  return {
    roleCounts,
    operationCounts,
    unsafeAggregationOverrideCount,
  };
}

function snapshotEntityResolutionSummary(plans = []) {
  const selections = [];
  let snapshotSeriesCount = 0;
  let appliedSeriesCount = 0;

  for (const plan of Array.isArray(plans) ? plans : []) {
    for (const series of Array.isArray(plan.series) ? plan.series : []) {
      if (
        series.operation !== SEMANTIC_AGGREGATION_OPERATION.LATEST ||
        series.metricRole !== SEMANTIC_METRIC_ROLE.STOCK_SNAPSHOT
      ) {
        continue;
      }
      snapshotSeriesCount += 1;
      const resolution = resolveSnapshotEntityHeaders(series);
      if (resolution.applied) appliedSeriesCount += 1;
      selections.push({
        tableIndex: series.tableIndex,
        metricLabel: series.metricLabel,
        applied: resolution.applied,
        headers: cloneValue(resolution.headers || []),
        entityCount: Number(resolution.entityCount || 0),
        reason: resolution.reason,
      });
    }
  }

  return {
    version: SNAPSHOT_ENTITY_RESOLVER_VERSION,
    snapshotSeriesCount,
    appliedSeriesCount,
    selections,
  };
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
  const rawPlans = sourceTables.map((table, index) =>
    buildSemanticSeries(table, index, context),
  );
  const relationshipAnalyses = rawPlans.map((plan) =>
    applyMetricRelationshipPriorities(plan.series),
  );
  const plans = rawPlans.map((plan, index) => ({
    ...plan,
    series: relationshipAnalyses[index].series,
    relationshipAnalysis: relationshipAnalyses[index],
  }));
  const sections = [];
  const sectionBudgetSummaries = [];

  sourceTables.forEach((table, tableIndex) => {
    const plan = plans[tableIndex];
    const tableSections = [];
    if (options.includeOverview !== false) {
      tableSections.push(
        overviewSection(table, tableIndex, plan.contract),
      );
    }
    if (options.includeDistinct !== false) {
      const distinctSection = buildDistinctEntitySection({
        table,
        tableIndex,
        metricIdFactory: metricId,
      });
      if (distinctSection) tableSections.push(distinctSection);
    }
    plan.series.forEach((series) => {
      tableSections.push(
        ...seriesSections(table, tableIndex, series, options),
      );
    });
    const budget = applySemanticSectionBudget({
      sections: tableSections,
      maxSections: Number(options.maxSectionsPerTable ?? 28),
    });
    sections.push(...budget.sections);
    sectionBudgetSummaries.push({
      tableIndex,
      tableLabel: tableLabel(table, tableIndex),
      ...budget,
      sections: undefined,
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

  const preliminarySections = normalizeSectionMetricIds(
    dedupeSections(limitedSections),
  );
  const expectedMetricIds = uniqueMetricIds(
    preliminarySections.flatMap((section) =>
      collectSectionMetricIds(section),
    ),
  );
  const finalOutputQualityGate =
    applyFinalOutputQualityGate({
      sections: preliminarySections,
      expectedMetricIds,
      sectionBudgetSummaries,
      throwOnFailure:
        options.enforceFinalOutputQualityGate !== false,
    });
  const finalSections = normalizeSectionMetricIds(
    finalOutputQualityGate.sections,
  );

  const canonicalLongTableCount = plans.filter(
    (plan) => plan.contract?.type === "canonical_long",
  ).length;
  const physicalWideTableCount = plans.filter(
    (plan) => plan.contract?.type === "physical_wide",
  ).length;
  const metricSemanticSummary = semanticRoleSummary(plans);
  const snapshotEntitySummary =
    snapshotEntityResolutionSummary(plans);
  const metricRelationships = relationshipAnalyses.flatMap(
    (analysis) => cloneValue(analysis.relations || []),
  );
  const primaryMetricLabels = Array.from(new Set(
    relationshipAnalyses.flatMap(
      (analysis) => analysis.primaryMetricLabels || [],
    ),
  ));
  const componentMetricLabels = Array.from(new Set(
    relationshipAnalyses.flatMap(
      (analysis) => analysis.componentMetricLabels || [],
    ),
  ));

  return {
    ok: true,
    resultType: "businessTemplate",
    operation: "genericStructuredSummary",
    templateId,
    title,
    sections: finalSections,
    contractSummaryCoverage: {
      version: SEMANTIC_OUTPUT_CONTRACT_VERSION,
      contractCatalogVersion:
        `${SEMANTIC_OUTPUT_CONTRACT_VERSION}_catalog`,
      expectedMetricIds,
    },
    finalOutputQualityGate: {
      version: FINAL_OUTPUT_QUALITY_GATE_VERSION,
      completenessVersion: OUTPUT_COMPLETENESS_CONTRACT_VERSION,
      duplicateResolverVersion:
        SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
      status: finalOutputQualityGate.status,
      pass: finalOutputQualityGate.pass,
      failureReasons: cloneValue(
        finalOutputQualityGate.failureReasons || [],
      ),
      expectedMetricCount:
        finalOutputQualityGate.analysis?.expectedMetricCount || 0,
      renderedExpectedMetricCount:
        finalOutputQualityGate.analysis?.renderedExpectedMetricCount || 0,
    },
    executionMeta: {
      sourceDataOnly: false,
      genericSummary: true,
      patchVersion,
      semanticOutputPlannerVersion:
        SEMANTIC_OUTPUT_PLANNER_VERSION,
      metricSemanticRoleEngineVersion:
        METRIC_SEMANTIC_ROLE_ENGINE_VERSION,
      aggregationContractResolverVersion:
        AGGREGATION_CONTRACT_RESOLVER_VERSION,
      metricRelationshipPriorityEngineVersion:
        METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
      derivedTotalRelationVersion:
        DERIVED_TOTAL_RELATION_VERSION,
      representativeMetricPriorityVersion:
        REPRESENTATIVE_METRIC_PRIORITY_VERSION,
      metricRelationshipCount: metricRelationships.length,
      metricRelationships: cloneValue(metricRelationships),
      primaryMetricLabels: cloneValue(primaryMetricLabels),
      componentMetricLabels: cloneValue(componentMetricLabels),
      semanticSectionBudgetEngineVersion:
        SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
      mandatorySummaryCoverageFloorVersion:
        MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
      outputCompletenessContractVersion:
        OUTPUT_COMPLETENESS_CONTRACT_VERSION,
      semanticOutputDuplicateResolverVersion:
        SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
      finalOutputQualityGateVersion:
        FINAL_OUTPUT_QUALITY_GATE_VERSION,
      finalOutputQualityGateApplied:
        finalOutputQualityGate.applied,
      finalOutputQualityGatePass:
        finalOutputQualityGate.pass,
      finalOutputQualityGateStatus:
        finalOutputQualityGate.status,
      finalOutputQualityGateStrict:
        options.enforceFinalOutputQualityGate !== false,
      finalOutputQualityGateFailureReasons:
        cloneValue(finalOutputQualityGate.failureReasons || []),
      finalOutputQualityGateExpectedMetricCount:
        Number(finalOutputQualityGate.analysis?.expectedMetricCount || 0),
      finalOutputQualityGateRenderedExpectedMetricCount:
        Number(
          finalOutputQualityGate.analysis?.renderedExpectedMetricCount || 0,
        ),
      finalOutputQualityGateMissingMetricIds:
        cloneValue(finalOutputQualityGate.analysis?.missingMetricIds || []),
      finalOutputQualityGateDuplicateMetricIds:
        cloneValue(finalOutputQualityGate.analysis?.duplicateMetricIds || []),
      finalOutputQualityGateRemovedDuplicateSectionCount:
        Number(
          finalOutputQualityGate.duplicateResolution
            ?.removedDuplicateSectionCount || 0,
        ),
      finalOutputQualityGateRemovedDuplicateSectionIds:
        cloneValue(
          finalOutputQualityGate.duplicateResolution
            ?.removedDuplicateSectionIds || [],
        ),
      finalOutputQualityGateRenamedSectionIdCount:
        Number(finalOutputQualityGate.renamedSectionIds?.length || 0),
      finalOutputQualityGateRenamedTitleCount:
        Number(finalOutputQualityGate.renamedTitles?.length || 0),
      durationSummaryContractVersion:
        DURATION_SUMMARY_CONTRACT_VERSION,
      distinctEntitySectionVersion:
        DISTINCT_ENTITY_SECTION_VERSION,
      sectionBudgetSummaries: cloneValue(sectionBudgetSummaries),
      semanticContractPrecedenceVersion:
        SEMANTIC_CONTRACT_PRECEDENCE_VERSION,
      semanticMetricRoleCounts:
        metricSemanticSummary.roleCounts,
      semanticAggregationOperationCounts:
        metricSemanticSummary.operationCounts,
      unsafeAggregationOverrideCount:
        metricSemanticSummary.unsafeAggregationOverrideCount,
      snapshotEntityResolverVersion:
        SNAPSHOT_ENTITY_RESOLVER_VERSION,
      snapshotEntitySeriesCount:
        snapshotEntitySummary.snapshotSeriesCount,
      snapshotEntityAppliedSeriesCount:
        snapshotEntitySummary.appliedSeriesCount,
      snapshotEntitySelections:
        cloneValue(snapshotEntitySummary.selections),
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
      plannedTableCount: plans.filter(
        (plan) => plan.series.length > 0,
      ).length,
      context: cloneValue(context),
      compactTitles: options.compactTitles === true,
      maxDimensionsPerSeries: Number(
        options.maxDimensionsPerSeries ?? 3,
      ),
      maxSectionsPerTable: Number(options.maxSectionsPerTable ?? 28),
      includeDistinct: options.includeDistinct !== false,
      preferVirtualTables: options.preferVirtualTables === true,
    },
  };
}

module.exports = {
  SEMANTIC_OUTPUT_PLANNER_LEGACY_VERSION,
  SEMANTIC_OUTPUT_PLANNER_FLOW_DIRECTION_VERSION,
  SEMANTIC_OUTPUT_PLANNER_PREVIOUS_VERSION,
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  SEMANTIC_OUTPUT_CONTRACT_VERSION,
  SEMANTIC_CONTRACT_PRECEDENCE_VERSION,
  MIXED_SECTION_ROW_PRECEDENCE_VERSION,
  GENERAL_STOCK_SNAPSHOT_ALIAS_VERSION,
  CONTRACT_KPI_SNAPSHOT_BRIDGE_VERSION,
  ACTUAL_FLOW_EVIDENCE_GATE_VERSION,
  SNAPSHOT_ENTITY_RESOLVER_VERSION,
  FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
  FLOW_DIRECTION_SECTION_REPAIR_VERSION,
  METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
  DERIVED_TOTAL_RELATION_VERSION,
  REPRESENTATIVE_METRIC_PRIORITY_VERSION,
  SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
  MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
  DURATION_SUMMARY_CONTRACT_VERSION,
  DISTINCT_ENTITY_SECTION_VERSION,
  OUTPUT_COMPLETENESS_CONTRACT_VERSION,
  SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
  FINAL_OUTPUT_QUALITY_GATE_VERSION,
  buildSemanticOutputPlan,
  buildSemanticSeries,
  augmentBusinessTemplateResult,
  canPlanSemanticOutput,
  selectPreferredSemanticTables,
  existingSectionCoversPlanned,
  isWholeMetricSummarySection,
  applyMandatorySummaryCoverageFloor,
  canonicalLongContract,
  physicalWideContract,
  dimensionSemanticPriority,
  pruneSupersededGenericSections,
  strictPeriodValue,
  inferAggregation,
  inferAggregationContract,
  latestRecordSelection,
  latestRecordSelectionByEntity,
  resolveSnapshotEntityHeaders,
  snapshotEntityResolutionSummary,
  operationStats,
  resolveSemanticContractConflicts,
  replaceMixedSectionRows,
  connectContractKpisToSemanticSnapshots,
  isGeneralStockSnapshotAlias,
  stockSnapshotSemanticSubtype,
  detectActualFlowEvidence,
  tableActualFlowEvidence,
  canonicalActualFlowDirection,
  canonicalFlowDirection,
  resolveFlowDirectionEvidence,
  buildSystemFlowSummary,
  buildDirectionRows,
  buildPeriodFlowRows,
  buildEntityFlowRows,
  buildLocationLedgerRows,
  applyFlowDirectionSemantics,
  isInventoryFlowMixedSection,
  contractRowAggregationIntent,
  sectionMetricAndGroupMatch,
  sectionOperationFamilies,
  semanticConflictReason,
  semanticRoleSummary,
  applyMetricRelationshipPriorities,
  prioritizeBusinessSections,
  applySemanticSectionBudget,
  buildDistinctEntitySection,
  sectionPolicyForSeries,
  applyFinalOutputQualityGate,
};
