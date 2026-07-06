const {
  isVirtualQueryTable,
  stripVirtualTableSuffix,
  resolveSourceRefForTableId,
  getSourceTableId,
  getCanonicalTableId,
  getTableUsage,
} = require("./sourceTablePolicy");

const MULTI_SOURCE_CANDIDATE_VERSION = "multi_source_candidate_builder_v1_4";
const MULTI_SOURCE_DETECTION_HOTFIX_VERSION =
  "multi_source_candidate_detection_hotfix_v1";
const MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION =
  "multi_source_candidate_payload_fallback_hotfix_v1";
const MULTI_SOURCE_SCHEMA_GROUP_VERSION = "multi_source_schema_group_v1_1";
const MULTI_SOURCE_SCHEMA_UNION_VERSION = "multi_source_schema_union_v1_1";
const MULTI_SOURCE_COMPATIBLE_SCHEMA_VERSION =
  "multi_source_compatible_schema_v1";
const MULTI_SOURCE_INDIVIDUAL_SOURCE_VERSION =
  "multi_source_individual_source_candidates_v1";
const UNION_SOURCE_DIMENSION_HEADER = "원본데이터시트";
const UNION_SOURCE_TABLE_ID_HEADER = "원본테이블ID";
const UNION_ORIGINAL_SHEET_HEADER = "원본시트명";
const UNION_PREVIEW_ROW_LIMIT = 30;

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const result = [];

  for (const value of values) {
    const text = String(value ?? "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    result.push(text);
  }

  return result;
}

function uniqueTables(tables = []) {
  const seen = new Set();
  const result = [];

  for (const table of asArray(tables)) {
    const key =
      tableIdOf(table) ||
      sourceTableIdOf(table) ||
      table.sheetName ||
      JSON.stringify(table).slice(0, 80);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(table);
  }

  return result;
}

function normalizeText(value = "") {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKey(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[\s_\-()\[\]{}.,:;\/\\]+/g, "")
    .replace(/(전체|남자|여자|남성|여성|유효|데이터|sheet\d*)/gi, "");
}

function stableIdKey(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[\s_\-()\[\]{}.,:;\/\#]+/g, "_")
    .replace(/[^\p{Letter}\p{Number}_]+/gu, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 140);
}

function tableIdOf(table = {}) {
  const explicitId = normalizeText(
    getCanonicalTableId?.(table) ||
      table.tableId ||
      table.id ||
      table.queryTableId ||
      table.normalizedTableId ||
      table.meta?.tableId ||
      table.metadata?.tableId ||
      "",
  );
  if (explicitId) return explicitId;

  const sourceId = normalizeText(
    table.sourceTableId ||
      table.physicalTableId ||
      table.parentTableId ||
      table.transformation?.sourceTableId ||
      table.transformation?.source?.tableId ||
      "",
  );
  const kind = normalizeText(
    table.transformation?.type ||
      table.transformationType ||
      table.normalizationType ||
      table.normalizedType ||
      table.virtualTableKind ||
      "",
  );
  if (sourceId && /WIDE_LONG/i.test(kind)) return `${sourceId}#WIDE_LONG`;
  if (sourceId && /CROSS_LONG/i.test(kind)) return `${sourceId}#CROSS_LONG`;
  return "";
}

function sourceTableIdOf(table = {}) {
  const id = tableIdOf(table);
  return normalizeText(
    table.sourceTableId ||
      table.physicalTableId ||
      table.parentTableId ||
      table.originTableId ||
      table.transformation?.sourceTableId ||
      table.transformation?.source?.tableId ||
      table.transformation?.source?.id ||
      table.source?.tableId ||
      table.source?.id ||
      getSourceTableId?.(table) ||
      stripVirtualTableSuffix(id) ||
      id,
  );
}

function sourceSheetNameOf(table = {}, policy = {}) {
  const ref = resolveSourceRefForTableId(
    tableIdOf(table) || sourceTableIdOf(table),
    policy,
  );
  return normalizeText(
    ref.sourceSheetName ||
      table.sourceSheetName ||
      table.sheetName ||
      table.name ||
      table.sheet ||
      "",
  );
}

function tableUsageOf(table = {}) {
  return (
    getTableUsage?.(table) ||
    table.tableUsage ||
    table.diagnostics?.tableUsage ||
    {}
  );
}

function isAnalysisEligible(table = {}) {
  const usage = tableUsageOf(table);
  return (
    usage.analysisEligible === true ||
    table.analysisEligible === true ||
    table.isAnalysisEligible === true
  );
}

function isExplicitlyIneligible(table = {}) {
  const usage = tableUsageOf(table);
  return (
    usage.analysisEligible === false ||
    table.analysisEligible === false ||
    table.isAnalysisEligible === false
  );
}

function isVirtualTable(table = {}) {
  const id = tableIdOf(table);
  const typeText = [
    table.transformation?.type,
    table.transformationType,
    table.normalizationType,
    table.normalizedType,
    table.virtualTableKind,
    table.tableKind,
    table.tableType,
    table.type,
  ]
    .filter(Boolean)
    .join(" ");
  return (
    isVirtualQueryTable?.(table) === true ||
    table.isVirtual === true ||
    table.virtual === true ||
    Boolean(table.transformation) ||
    /#(?:WIDE_LONG|CROSS_LONG)(?:_|$)/i.test(id) ||
    /(?:WIDE_LONG|CROSS_LONG)/i.test(typeText)
  );
}

function normalizeColumnItem(item, index = 0) {
  if (item == null) return null;
  if (typeof item === "string" || typeof item === "number") {
    const header = normalizeText(item);
    return header ? { header, key: header, index } : null;
  }
  if (typeof item !== "object") return null;

  const header = normalizeText(
    item.header ||
      item.name ||
      item.key ||
      item.field ||
      item.fieldName ||
      item.columnName ||
      item.label ||
      item.title ||
      item.displayName ||
      item.id ||
      item.accessor ||
      "",
  );
  if (!header) return null;
  return { ...item, header, key: item.key || item.field || header, index };
}

function candidateColumnArrays(table = {}) {
  return [
    table.columns,
    table.columnDefs,
    table.columnDefinitions,
    table.columnProfiles,
    table.columnMeta,
    table.columnsMeta,
    table.fields,
    table.schema?.columns,
    table.schema?.fields,
    table.schema?.headers,
    table.meta?.columns,
    table.metadata?.columns,
    table.normalizedColumns,
    table.normalizedHeaders,
    table.flattenedHeaders,
    table.headers,
    table.header,
    table.headerRow,
    table.headerRows,
  ].filter((value) => Array.isArray(value) && value.length);
}

function flattenColumnArray(value) {
  const result = [];
  for (const item of asArray(value)) {
    if (Array.isArray(item)) result.push(...flattenColumnArray(item));
    else result.push(item);
  }
  return result;
}

function rowObjects(table = {}) {
  return (
    [
      table.rows,
      table.dataRows,
      table.records,
      table.values,
      table.sampleRows,
      table.previewRows,
      table.normalizedRows,
    ].find((rows) => Array.isArray(rows) && rows.length) || []
  );
}

function inferColumnsFromRows(table = {}) {
  const rows = rowObjects(table);
  const firstObject = rows.find(
    (row) => row && typeof row === "object" && !Array.isArray(row),
  );
  if (!firstObject) return [];
  return Object.keys(firstObject).map((key, index) => ({
    header: normalizeText(key),
    key,
    index,
  }));
}

function normalizedColumns(table = {}) {
  const columns = [];
  for (const columnArray of candidateColumnArrays(table)) {
    flattenColumnArray(columnArray).forEach((column, index) => {
      const normalized = normalizeColumnItem(column, index);
      if (normalized) columns.push(normalized);
    });
  }
  columns.push(...inferColumnsFromRows(table));

  const seen = new Set();
  return columns.filter((column) => {
    const key = normalizeKey(column.header);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function columnHeaders(table = {}) {
  return normalizedColumns(table)
    .map((column) => normalizeText(column.header))
    .filter(Boolean);
}

function valueForHeader(row = {}, header = "") {
  if (!row || typeof row !== "object" || Array.isArray(row)) return undefined;
  if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
  const key = Object.keys(row).find(
    (candidate) => normalizeKey(candidate) === normalizeKey(header),
  );
  return key ? row[key] : undefined;
}

function isNumericValue(value) {
  if (value == null || value === "") return false;
  if (typeof value === "number") return Number.isFinite(value);
  const text = String(value).replace(/,/g, "").replace(/%$/, "").trim();
  if (!text) return false;
  const num = Number(text);
  return Number.isFinite(num);
}

function sampleSuggestsNumeric(table = {}, header = "") {
  const rows = rowObjects(table).slice(0, 20);
  let checked = 0;
  let numeric = 0;
  for (const row of rows) {
    const value = valueForHeader(row, header);
    if (value == null || value === "") continue;
    checked += 1;
    if (isNumericValue(value)) numeric += 1;
  }
  return checked > 0 && numeric / checked >= 0.6;
}

function isNumericColumn(column = {}, table = {}) {
  const type = String(
    column.type ||
      column.profileType ||
      column.dominantType ||
      column.dataType ||
      column.valueType ||
      column.semanticType ||
      column.role ||
      "",
  ).toLowerCase();
  return (
    type.includes("number") ||
    type.includes("numeric") ||
    type.includes("integer") ||
    type.includes("float") ||
    type.includes("double") ||
    type.includes("decimal") ||
    sampleSuggestsNumeric(table, column.header)
  );
}

function isPeriodLike(header = "") {
  return /(연도|년도|연월|월|분기|기간|일자|날짜|date|year|month|quarter|period)/i.test(
    header,
  );
}

function isIdentifierLike(header = "") {
  return /(id|코드|번호|순번|index|no\.?$|식별)/i.test(header);
}

function isKnownMetricHeader(header = "") {
  return /(순매출액|매출액|판매금액|거래액|출하액|매출수량|판매수량|수량|집행금액|집행액|현금|현물|총\s*연구비|정부출연금|민간부담금|연구비|연봉|급여|지표값|값|건수|비용|금액|예산|출연금|amount|sales|count|value|budget|expense|grant|salary)/i.test(
    header,
  );
}

function metricHeaders(table = {}) {
  return normalizedColumns(table)
    .filter((column) => {
      const header = normalizeText(
        column.header || column.name || column.key || "",
      );
      if (!header) return false;
      if (isPeriodLike(header) || isIdentifierLike(header)) return false;
      return isNumericColumn(column, table) || isKnownMetricHeader(header);
    })
    .map((column) =>
      normalizeText(column.header || column.name || column.key || ""),
    )
    .filter(Boolean);
}

function dimensionHeaders(table = {}) {
  const metricSet = new Set(metricHeaders(table).map(normalizeKey));
  return normalizedColumns(table)
    .filter((column) => {
      const header = normalizeText(
        column.header || column.name || column.key || "",
      );
      if (!header) return false;
      if (isIdentifierLike(header)) return false;
      return !metricSet.has(normalizeKey(header));
    })
    .map((column) =>
      normalizeText(column.header || column.name || column.key || ""),
    )
    .filter(Boolean);
}

function overlap(left = [], right = []) {
  const rightSet = new Set(right.map(normalizeKey).filter(Boolean));
  return left.filter((item) => rightSet.has(normalizeKey(item)));
}

function sharedHeaders(tables = [], picker = columnHeaders) {
  const lists = tables
    .map((table) => picker(table))
    .filter((list) => list.length);
  if (!lists.length) return [];
  return lists.slice(1).reduce((acc, list) => overlap(acc, list), lists[0]);
}

function sourceEntries(policy = {}) {
  return asArray(policy.sourceTableEntries);
}

function sourceTablesFromPolicy(policy = {}) {
  return uniqueTables([
    ...sourceEntries(policy)
      .map((entry) => entry.table)
      .filter(Boolean),
    ...asArray(policy.sourceTables),
  ]);
}

function sourceEntryForTable(table = {}, policy = {}) {
  const sourceId = sourceTableIdOf(table);
  const id = tableIdOf(table);
  return sourceEntries(policy).find(
    (entry) =>
      entry.tableId === id ||
      entry.sourceTableId === id ||
      entry.tableId === sourceId ||
      entry.sourceTableId === sourceId,
  );
}

function physicalTablesFromNormalized(tables = []) {
  return uniqueTables(
    asArray(tables).filter(
      (table) => !isVirtualTable(table) && !isExplicitlyIneligible(table),
    ),
  );
}

function titleForPhysicalGroup(tables = []) {
  const sheetNames = unique(
    tables.map(
      (table) => table.sheetName || table.name || table.tableName || "",
    ),
  );
  const hasFitness = sheetNames.some((name) => /체력|평가|fitness/i.test(name));
  const hasGenderSplit = sheetNames.some((name) =>
    /(전체|남자|여자|남성|여성)/.test(name),
  );

  if (hasFitness && hasGenderSplit) return "전체·남자·여자 원본 통합 비교 후보";
  if (sheetNames.length >= 2)
    return `${sheetNames.slice(0, 3).join("·")} 원본 통합 비교 후보`;
  return "다중 원본 통합 비교 후보";
}

function makePhysicalMultiSourceCandidate({
  sourceTables = [],
  policy = {},
} = {}) {
  if (sourceTables.length < 2) return null;

  const sourceIds = unique(sourceTables.map(sourceTableIdOf));
  const sourceSheets = unique(
    sourceTables.map((table) => sourceSheetNameOf(table, policy)),
  );
  const sharedDimensions = sharedHeaders(sourceTables, dimensionHeaders);
  const sharedMetrics = sharedHeaders(sourceTables, metricHeaders);
  const sharedColumnHeaders = sharedHeaders(sourceTables, columnHeaders);

  if (sourceIds.length < 2) return null;

  const hasSharedSchema = Boolean(
    sharedColumnHeaders.length ||
    sharedDimensions.length ||
    sharedMetrics.length,
  );

  return {
    candidateId: `multi_source_physical_${sourceIds.map(stableIdKey).join("_").slice(0, 120)}`,
    candidateType: "multiSource",
    type: "multiSource",
    multiSourceCandidateKind: "physicalComparison",
    multiSourceCandidateVersion: MULTI_SOURCE_CANDIDATE_VERSION,
    multiSourceDetectionHotfixVersion: MULTI_SOURCE_DETECTION_HOTFIX_VERSION,
    title: titleForPhysicalGroup(sourceTables),
    description:
      "분리된 여러 원본 표를 같은 분석 단위로 묶어 비교·통합 후보로 제공합니다.",
    sourceScope: "multiTable",
    sourceTableIds: sourceIds,
    sourceSheetNames: sourceSheets,
    sourceTableId: sourceIds[0] || "",
    sourceSheetName: sourceSheets[0] || "",
    sourceTables: sourceTables.map((table, index) => {
      const entry = sourceEntryForTable(table, policy) || {};
      return {
        sourceTableIndex: entry.sourceTableIndex ?? index,
        sourceTableId: sourceTableIdOf(table),
        tableId: tableIdOf(table),
        sheetName: table.sheetName || table.name || "",
        sourceSheetName:
          entry.sourceSheetName || sourceSheetNameOf(table, policy),
        rowCount: Number(
          table.rowCount || asArray(rowObjects(table)).length || 0,
        ),
        columnCount: columnHeaders(table).length,
        analysisEligible: isAnalysisEligible(table),
      };
    }),
    sharedColumns: sharedColumnHeaders,
    sharedDimensions,
    sharedMetrics,
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    recipeIds: ["multi_source_comparison", "multi_source_dashboard"],
    confidence: sharedMetrics.length ? 0.82 : hasSharedSchema ? 0.74 : 0.66,
    priority:
      780 + Math.min(40, sourceIds.length * 5 + sharedColumnHeaders.length),
    reasonCodes: unique([
      "MULTI_SOURCE_CANDIDATE_BUILDER_V1_1",
      "MULTIPLE_PHYSICAL_SOURCE_TABLES",
      hasSharedSchema
        ? "SHARED_OR_DETECTED_SCHEMA"
        : "SCHEMA_FALLBACK_BY_TABLE_COUNT",
      sharedMetrics.length ? "SHARED_METRICS" : "NO_SHARED_METRICS",
      sourceTables.some((table) => table.isPrimary)
        ? "HAS_PRIMARY_SOURCE_TABLE"
        : "",
    ]),
    recommendationReason: hasSharedSchema
      ? "여러 원본데이터 시트가 동일하거나 유사한 구조를 가지므로 통합 비교 자동화 후보로 사용할 수 있습니다."
      : "여러 분석 가능 원본데이터 시트가 존재하므로 통합 비교 후보로 추적합니다.",
  };
}

function hashText(value = "") {
  let hash = 5381;
  const text = String(value || "");
  for (let i = 0; i < text.length; i += 1) {
    hash = (hash * 33) ^ text.charCodeAt(i);
  }
  return (hash >>> 0).toString(36);
}

function schemaColumnKind(column = {}, table = {}) {
  const header = normalizeText(
    column.header || column.name || column.key || "",
  );
  const metricSet = new Set(metricHeaders(table).map(normalizeKey));
  if (metricSet.has(normalizeKey(header))) return "metric";
  if (isPeriodLike(header)) return "period";
  if (isIdentifierLike(header)) return "id";
  return "dimension";
}

function isSchemaNoiseHeader(header = "") {
  const raw = String(header || "").trim();
  const normalized = normalizeKey(raw);
  if (!raw || !normalized) return true;

  // 서식/검증/출처/주석성 열은 schema grouping을 깨뜨리기 쉽다.
  // 특히 다중시트 회귀 케이스의 "테스트 목적_... 전체/남자/여자 ..." 열은
  // 실제 분석 축이 아니라 설명 열이므로 compatible schema 판단에서 제외한다.
  if (
    /테스트\s*목적|테스트설명|사용\s*안내|설명|비고|주석|출처|자료\s*기준|작성\s*기관|작성\s*일|조회\s*기간|다운로드|통계표|url|metadata|meta|note|comment|remark|source/i.test(
      raw,
    )
  ) {
    return true;
  }

  // 너무 긴 자연어형 헤더는 보통 메모/주석성 열이다.
  if (raw.length >= 28 && /[.:：]|\s/.test(raw)) return true;

  return false;
}

function comparableSchemaColumnsForTable(table = {}) {
  return schemaColumnsForTable(table).filter(
    (column) => !isSchemaNoiseHeader(column.header),
  );
}

function schemaColumnKeys(columns = []) {
  return asArray(columns)
    .map((column) => column.key || normalizeKey(column.header))
    .filter(Boolean);
}

function sameTableGroupKey(tables = []) {
  return unique(asArray(tables).map(sourceTableIdOf)).sort().join("|");
}

function compatibleSchemaSignatureForTable(table = {}) {
  const columns = comparableSchemaColumnsForTable(table);
  if (columns.length < 2) return "";

  const kindSeq = columns.map((column) => column.kind || "dimension").join("|");
  const metricCount = columns.filter(
    (column) => column.kind === "metric",
  ).length;
  const periodCount = columns.filter(
    (column) => column.kind === "period",
  ).length;
  const dimensionCount = columns.filter(
    (column) => column.kind === "dimension",
  ).length;

  // 헤더명이 일부 달라도 열의 역할/개수 구성이 같으면 compatible schema로 본다.
  // union 후보는 실제 row를 합칠 때 공통 헤더만 쓰므로 안전하다.
  return [
    `cols:${columns.length}`,
    `kinds:${kindSeq}`,
    `m:${metricCount}`,
    `p:${periodCount}`,
    `d:${dimensionCount}`,
  ].join("|");
}

function compatibleGroupQuality(tables = []) {
  const sourceTables = uniqueTables(tables);
  const sharedColumnHeaders = sharedHeaders(sourceTables, (table) =>
    comparableSchemaColumnsForTable(table).map((column) => column.header),
  );
  const sharedDimensions = sharedHeaders(sourceTables, dimensionHeaders).filter(
    (header) => !isSchemaNoiseHeader(header),
  );
  const sharedMetrics = sharedHeaders(sourceTables, metricHeaders).filter(
    (header) => !isSchemaNoiseHeader(header),
  );
  const rowCount = sourceTables.reduce(
    (sum, table) =>
      sum + Number(table.rowCount || rowObjects(table).length || 0),
    0,
  );

  const firstColumns = comparableSchemaColumnsForTable(sourceTables[0] || {});
  const minComparableColumnCount = Math.min(
    ...sourceTables.map(
      (table) => comparableSchemaColumnsForTable(table).length,
    ),
  );
  const maxComparableColumnCount = Math.max(
    ...sourceTables.map(
      (table) => comparableSchemaColumnsForTable(table).length,
    ),
  );

  return {
    sharedColumnHeaders,
    sharedDimensions,
    sharedMetrics,
    rowCount,
    firstColumns,
    minComparableColumnCount,
    maxComparableColumnCount,
    columnCountGap: maxComparableColumnCount - minComparableColumnCount,
    ok:
      sourceTables.length >= 2 &&
      minComparableColumnCount >= 2 &&
      rowCount > 0 &&
      (sharedColumnHeaders.length >= 2 ||
        sharedMetrics.length >= 1 ||
        sharedDimensions.length >= 2 ||
        maxComparableColumnCount === minComparableColumnCount),
  };
}

function schemaColumnsForTable(table = {}) {
  return normalizedColumns(table)
    .map((column) => {
      const header = normalizeText(
        column.header || column.name || column.key || "",
      );
      const key = normalizeKey(header);
      if (!key) return null;
      return {
        header,
        key,
        kind: schemaColumnKind(column, table),
      };
    })
    .filter(Boolean);
}

function schemaSignatureForTable(table = {}) {
  const columns = comparableSchemaColumnsForTable(table);
  if (columns.length < 2) return "";

  // 같은 schema 판단은 열 순서까지 포함한다. 단, 설명/주석성 열은 제외한다.
  return columns.map((column) => `${column.key}:${column.kind}`).join("|");
}

function groupTablesBySchemaSignature(sourceTables = []) {
  const groups = new Map();

  for (const table of asArray(sourceTables)) {
    const signature = schemaSignatureForTable(table);
    if (!signature) continue;

    if (!groups.has(signature)) {
      groups.set(signature, {
        schemaSignature: signature,
        schemaSignatureHash: hashText(signature),
        schemaColumns: comparableSchemaColumnsForTable(table),
        tables: [],
      });
    }

    groups.get(signature).tables.push(table);
  }

  return [...groups.values()]
    .map((group) => ({
      ...group,
      tables: uniqueTables(group.tables),
    }))
    .filter((group) => group.tables.length >= 2);
}

function groupTablesByCompatibleSchema(sourceTables = [], exactGroups = []) {
  const source = uniqueTables(sourceTables);
  if (source.length < 2) return [];

  const exactKeys = new Set(
    exactGroups.map((group) => sameTableGroupKey(group.tables)),
  );
  const groups = new Map();

  for (const table of source) {
    const signature = compatibleSchemaSignatureForTable(table);
    if (!signature) continue;
    if (!groups.has(signature)) {
      groups.set(signature, {
        schemaSignature: `compatible:${signature}`,
        schemaSignatureHash: hashText(`compatible:${signature}`),
        schemaCompatibilityMode: "compatible",
        schemaColumns: comparableSchemaColumnsForTable(table),
        tables: [],
      });
    }
    groups.get(signature).tables.push(table);
  }

  const compatible = [...groups.values()]
    .map((group) => {
      const tables = uniqueTables(group.tables);
      const quality = compatibleGroupQuality(tables);
      return {
        ...group,
        tables,
        schemaColumns: group.schemaColumns?.length
          ? group.schemaColumns
          : quality.firstColumns,
        compatibility: quality,
      };
    })
    .filter((group) => {
      if (group.tables.length < 2) return false;
      if (exactKeys.has(sameTableGroupKey(group.tables))) return false;
      return group.compatibility?.ok === true;
    });

  // kind/개수 grouping도 실패하면, 다중 physical table 전체를 compatible fallback으로 묶는다.
  if (!compatible.length) {
    const quality = compatibleGroupQuality(source);
    const key = sameTableGroupKey(source);
    if (quality.ok && !exactKeys.has(key)) {
      return [
        {
          schemaSignature: `compatible:fallback:${hashText(key)}`,
          schemaSignatureHash: hashText(`compatible:fallback:${key}`),
          schemaCompatibilityMode: "compatibleFallback",
          schemaColumns: quality.firstColumns,
          tables: source,
          compatibility: quality,
        },
      ];
    }
  }

  return compatible;
}

function commonSchemaHeadersForGroup(sourceTables = [], group = {}) {
  const common = sharedHeaders(sourceTables, (table) =>
    comparableSchemaColumnsForTable(table).map((column) => column.header),
  ).filter((header) => !isSchemaNoiseHeader(header));
  if (common.length) return common;

  const schemaHeaders = asArray(group.schemaColumns)
    .map((column) => column.header)
    .filter((header) => header && !isSchemaNoiseHeader(header));
  if (schemaHeaders.length) return schemaHeaders;

  return comparableSchemaColumnsForTable(sourceTables[0] || {}).map(
    (column) => column.header,
  );
}

function compactLinkedAnalysisCandidate(candidate = {}) {
  return {
    candidateId:
      candidate.candidateId || candidate.id || candidate.recipeId || "",
    recipeType:
      candidate.recipeType || candidate.type || candidate.recipeId || "",
    title: candidate.title || "",
    sourceTableId: candidate.sourceTableId || candidate.tableId || "",
    sourceSheetName: candidate.sourceSheetName || "",
    outputTypes: asArray(candidate.outputTypes),
  };
}

function linkedAnalysisCandidatesForSourceTables({
  sourceTables = [],
  analysisRecipeCandidates = [],
} = {}) {
  const sourceIds = new Set(
    sourceTables
      .flatMap((table) => [sourceTableIdOf(table), tableIdOf(table)])
      .filter(Boolean),
  );

  const linked = flattenCandidateItems(analysisRecipeCandidates).filter(
    (candidate) => {
      const ids = candidateSourceIds(candidate).flatMap((id) => [
        id,
        stripVirtualTableSuffix(id),
      ]);
      return ids.some((id) => sourceIds.has(id));
    },
  );

  const seen = new Set();
  return linked.map(compactLinkedAnalysisCandidate).filter((candidate) => {
    const key = candidate.candidateId || JSON.stringify(candidate);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildUnionPreviewRows({
  sourceTables = [],
  headers = [],
  policy = {},
  limit = UNION_PREVIEW_ROW_LIMIT,
} = {}) {
  const previewRows = [];
  let rowCount = 0;

  for (const table of sourceTables) {
    const rows = rowObjects(table);
    const sourceSheetName = sourceSheetNameOf(table, policy);
    const sourceTableId = sourceTableIdOf(table);
    const originalSheetName =
      table.sheetName || table.name || table.tableName || "";

    for (const row of rows) {
      rowCount += 1;

      if (previewRows.length >= limit) continue;

      const next = {
        [UNION_SOURCE_DIMENSION_HEADER]: sourceSheetName,
        [UNION_SOURCE_TABLE_ID_HEADER]: sourceTableId,
        [UNION_ORIGINAL_SHEET_HEADER]: originalSheetName,
      };

      for (const header of headers) {
        next[header] = valueForHeader(row, header);
      }

      previewRows.push(next);
    }
  }

  return { previewRows, rowCount };
}

function sourceTableSummaries(sourceTables = [], policy = {}) {
  return sourceTables.map((table, index) => {
    const entry = sourceEntryForTable(table, policy) || {};
    return {
      sourceTableIndex: entry.sourceTableIndex ?? index,
      sourceTableId: sourceTableIdOf(table),
      tableId: tableIdOf(table),
      sheetName: table.sheetName || table.name || "",
      sourceSheetName:
        entry.sourceSheetName || sourceSheetNameOf(table, policy),
      rowCount: Number(table.rowCount || rowObjects(table).length || 0),
      columnCount: columnHeaders(table).length,
      schemaSignature: schemaSignatureForTable(table),
      analysisEligible: isAnalysisEligible(table),
    };
  });
}

function makeSchemaUnionCandidate({
  group = {},
  groupIndex = 0,
  policy = {},
  analysisRecipeCandidates = [],
} = {}) {
  const sourceTables = uniqueTables(group.tables || []);
  if (sourceTables.length < 2) return null;

  const sourceIds = unique(sourceTables.map(sourceTableIdOf));
  const sourceSheets = unique(
    sourceTables.map((table) => sourceSheetNameOf(table, policy)),
  );
  const headers = commonSchemaHeadersForGroup(sourceTables, group);
  const sharedDimensions = sharedHeaders(sourceTables, dimensionHeaders);
  const sharedMetrics = sharedHeaders(sourceTables, metricHeaders);
  const { previewRows, rowCount } = buildUnionPreviewRows({
    sourceTables,
    headers,
    policy,
  });
  const linkedAnalysisCandidates = linkedAnalysisCandidatesForSourceTables({
    sourceTables,
    analysisRecipeCandidates,
  });
  const hash = group.schemaSignatureHash || hashText(group.schemaSignature);

  return {
    candidateId: `multi_source_schema_union_${hash}`,
    candidateType: "multiSource",
    type: "multiSource",
    multiSourceCandidateKind: "schemaUnion",
    multiSourceCandidateVersion: MULTI_SOURCE_CANDIDATE_VERSION,
    multiSourceSchemaGroupVersion: MULTI_SOURCE_SCHEMA_GROUP_VERSION,
    multiSourceSchemaUnionVersion: MULTI_SOURCE_SCHEMA_UNION_VERSION,
    title:
      sourceSheets.length >= 2
        ? `${sourceSheets.slice(0, 3).join("·")} 통합 분석 후보`
        : "동일 구조 원본 통합 분석 후보",
    description:
      "동일 schema 원본데이터들을 하나의 union 분석 단위로 묶고, 원본데이터시트를 구분 dimension으로 추가합니다.",
    sourceScope: "multiTable",
    sourceTableIds: sourceIds,
    sourceSheetNames: sourceSheets,
    sourceTableId: sourceIds[0] || "",
    sourceSheetName: sourceSheets[0] || "",
    sourceTables: sourceTableSummaries(sourceTables, policy),
    schemaGroupIndex: groupIndex,
    schemaSignature: group.schemaSignature,
    schemaSignatureHash: hash,
    schemaCompatibilityMode: group.schemaCompatibilityMode || "exact",
    schemaCompatibility: group.compatibility || null,
    schemaColumns: group.schemaColumns || [],
    sharedColumns: headers,
    sharedDimensions,
    sharedMetrics,
    unionSourceDimension: UNION_SOURCE_DIMENSION_HEADER,
    unionTable: {
      tableId: `union_${hash}`,
      tableName: "다중 원본 통합 테이블",
      virtualTableKind: "MULTI_SOURCE_UNION",
      sourceScope: "multiTable",
      isVirtual: true,
      rowCount,
      previewRowLimit: UNION_PREVIEW_ROW_LIMIT,
      columns: [
        {
          header: UNION_SOURCE_DIMENSION_HEADER,
          type: "category",
          role: "dimension",
        },
        {
          header: UNION_SOURCE_TABLE_ID_HEADER,
          type: "text",
          role: "id",
        },
        {
          header: UNION_ORIGINAL_SHEET_HEADER,
          type: "category",
          role: "dimension",
        },
        ...headers.map((header) => ({
          header,
          type: sharedMetrics.some(
            (metric) => normalizeKey(metric) === normalizeKey(header),
          )
            ? "number"
            : "text",
          role: sharedMetrics.some(
            (metric) => normalizeKey(metric) === normalizeKey(header),
          )
            ? "metric"
            : "dimension",
        })),
      ],
      previewRows,
    },
    unionRowCount: rowCount,
    unionPreviewRows: previewRows,
    linkedAnalysisCandidates,
    linkedAnalysisCandidateCount: linkedAnalysisCandidates.length,
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    recipeIds: unique([
      "multi_source_schema_union",
      "multi_source_dashboard",
      sharedMetrics.length
        ? "multi_source_metric_comparison"
        : "multi_source_count_comparison",
      ...linkedAnalysisCandidates.map((candidate) => candidate.recipeType),
    ]),
    confidence: sharedMetrics.length ? 0.86 : 0.78,
    priority:
      820 + Math.min(60, sourceIds.length * 8 + sharedMetrics.length * 4),
    reasonCodes: unique([
      "MULTI_SOURCE_CANDIDATE_BUILDER_V1_4",
      group.schemaCompatibilityMode
        ? "COMPATIBLE_SCHEMA_GROUPED"
        : "SCHEMA_SIGNATURE_GROUPED",
      "MULTI_SOURCE_SCHEMA_UNION",
      "SOURCE_SHEET_NAME_DIMENSION_ADDED",
      rowCount ? "UNION_ROWS_AVAILABLE" : "UNION_PREVIEW_EMPTY",
      sharedMetrics.length ? "SHARED_METRICS" : "NO_SHARED_METRICS",
      sharedDimensions.length ? "SHARED_DIMENSIONS" : "NO_SHARED_DIMENSIONS",
    ]),
    recommendationReason:
      "동일 구조의 원본데이터가 여러 개 있으므로 원본데이터시트 dimension을 추가해 전체 통합·시트별 비교 분석을 만들 수 있습니다.",
  };
}

function makeSchemaComparisonCandidate({
  group = {},
  groupIndex = 0,
  policy = {},
} = {}) {
  const sourceTables = uniqueTables(group.tables || []);
  const base = makePhysicalMultiSourceCandidate({ sourceTables, policy });
  if (!base) return null;
  const hash = group.schemaSignatureHash || hashText(group.schemaSignature);
  return {
    ...base,
    candidateId: `multi_source_schema_compare_${hash}`,
    multiSourceCandidateKind: "schemaComparison",
    multiSourceSchemaGroupVersion: MULTI_SOURCE_SCHEMA_GROUP_VERSION,
    title: `${base.sourceSheetNames.slice(0, 3).join("·")} 동일 구조 비교 후보`,
    schemaGroupIndex: groupIndex,
    schemaSignature: group.schemaSignature,
    schemaSignatureHash: hash,
    schemaCompatibilityMode: group.schemaCompatibilityMode || "exact",
    schemaCompatibility: group.compatibility || null,
    schemaColumns: group.schemaColumns || [],
    priority: Math.max(Number(base.priority || 0), 800),
    reasonCodes: unique([
      ...(base.reasonCodes || []),
      "MULTI_SOURCE_CANDIDATE_BUILDER_V1_4",
      group.schemaCompatibilityMode
        ? "COMPATIBLE_SCHEMA_GROUPED"
        : "SCHEMA_SIGNATURE_GROUPED",
      "SCHEMA_LEVEL_COMPARISON",
    ]),
    recommendationReason:
      "동일 schema 원본데이터끼리 묶어 시트 간 비교 후보로 제공합니다.",
  };
}

function makeIndividualSourceCandidate({
  table = {},
  policy = {},
  schemaSignature = "",
  schemaSignatureHash = "",
  analysisRecipeCandidates = [],
  index = 0,
} = {}) {
  const sourceId = sourceTableIdOf(table);
  if (!sourceId) return null;

  const sourceSheetName =
    sourceSheetNameOf(table, policy) || fallbackSourceSheetName(index);
  const metrics = metricHeaders(table);
  const dimensions = dimensionHeaders(table);
  const linkedAnalysisCandidates = linkedAnalysisCandidatesForSourceTables({
    sourceTables: [table],
    analysisRecipeCandidates,
  });
  const hash = schemaSignatureHash || hashText(schemaSignature || sourceId);

  return {
    candidateId: `multi_source_individual_${stableIdKey(sourceId).slice(0, 100)}_${hash}`,
    candidateType: "multiSource",
    type: "multiSource",
    multiSourceCandidateKind: "individualSource",
    multiSourceCandidateVersion: MULTI_SOURCE_CANDIDATE_VERSION,
    multiSourceIndividualSourceVersion: MULTI_SOURCE_INDIVIDUAL_SOURCE_VERSION,
    title: `${sourceSheetName} 개별 자동화 후보`,
    description:
      "다중 원본 파일 안의 개별 원본데이터 시트를 독립 자동화 후보로 추적합니다.",
    sourceScope: "singleTable",
    sourceTableIds: [sourceId],
    sourceSheetNames: [sourceSheetName],
    sourceTableId: sourceId,
    sourceSheetName,
    sourceTables: sourceTableSummaries([table], policy),
    schemaSignature,
    schemaSignatureHash: hash,
    sharedColumns: columnHeaders(table),
    sharedDimensions: dimensions,
    sharedMetrics: metrics,
    linkedAnalysisCandidates,
    linkedAnalysisCandidateCount: linkedAnalysisCandidates.length,
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    recipeIds: unique([
      "single_source_dashboard",
      metrics.length
        ? "single_source_metric_summary"
        : "single_source_count_summary",
      ...linkedAnalysisCandidates.map((candidate) => candidate.recipeType),
    ]),
    confidence: metrics.length ? 0.68 : 0.6,
    priority: 600 + Math.min(35, metrics.length * 4 + dimensions.length),
    reasonCodes: unique([
      "MULTI_SOURCE_CANDIDATE_BUILDER_V1_4",
      "INDIVIDUAL_SOURCE_CANDIDATE",
      "SOURCE_SHEET_NAME_SCOPED",
      metrics.length ? "HAS_SOURCE_METRICS" : "NO_SOURCE_METRICS",
    ]),
    recommendationReason:
      "다중 원본 중 특정 원본데이터 시트만 대상으로 자동화 시트를 생성할 수 있도록 개별 후보를 유지합니다.",
  };
}

function buildSchemaGroupCandidates({
  sourceTables = [],
  policy = {},
  analysisRecipeCandidates = [],
} = {}) {
  const exactGroups = groupTablesBySchemaSignature(sourceTables);
  const compatibleGroups = groupTablesByCompatibleSchema(
    sourceTables,
    exactGroups,
  );
  const groups = [...exactGroups, ...compatibleGroups];
  const candidates = [];

  groups.forEach((group, groupIndex) => {
    const comparison = makeSchemaComparisonCandidate({
      group,
      groupIndex,
      policy,
    });
    if (comparison) candidates.push(comparison);

    const union = makeSchemaUnionCandidate({
      group,
      groupIndex,
      policy,
      analysisRecipeCandidates,
    });
    if (union) candidates.push(union);
  });

  const tableToGroup = new Map();
  groups.forEach((group) => {
    const hash = group.schemaSignatureHash || hashText(group.schemaSignature);
    group.tables.forEach((table) => {
      tableToGroup.set(sourceTableIdOf(table), {
        schemaSignature: group.schemaSignature,
        schemaSignatureHash: hash,
      });
    });
  });

  sourceTables.forEach((table, index) => {
    const groupMeta = tableToGroup.get(sourceTableIdOf(table)) || {
      schemaSignature: schemaSignatureForTable(table),
      schemaSignatureHash: hashText(
        schemaSignatureForTable(table) || sourceTableIdOf(table),
      ),
    };
    const individual = makeIndividualSourceCandidate({
      table,
      policy,
      analysisRecipeCandidates,
      index,
      ...groupMeta,
    });
    if (individual) candidates.push(individual);
  });

  return candidates;
}

function virtualKind(table = {}) {
  const id = tableIdOf(table);
  const kind = [
    table.transformation?.type,
    table.transformationType,
    table.normalizationType,
    table.normalizedType,
    table.virtualTableKind,
    table.tableType,
    table.type,
  ].join(" ");
  if (/#WIDE_LONG(?:_|$)/i.test(id) || /wide_long/i.test(kind))
    return "WIDE_LONG";
  if (/#CROSS_LONG(?:_|$)/i.test(id) || /cross_long/i.test(kind))
    return "CROSS_LONG";
  return "VIRTUAL";
}

function makeVirtualLinkCandidate({
  virtualTable = {},
  physicalTable = null,
  policy = {},
} = {}) {
  const virtualId = tableIdOf(virtualTable);
  const physicalId = sourceTableIdOf(virtualTable);
  if (!virtualId || !physicalId || virtualId === physicalId) return null;

  const sourceRef = resolveSourceRefForTableId(virtualId, policy);
  const sourceSheetName =
    sourceRef.sourceSheetName ||
    sourceSheetNameOf(physicalTable || virtualTable, policy);
  const kind = virtualKind(virtualTable);
  const metrics = metricHeaders(virtualTable);
  const dimensions = dimensionHeaders(virtualTable);
  const sourceIds = unique([physicalId, virtualId]);

  return {
    candidateId: `multi_source_virtual_${stableIdKey(virtualId).slice(0, 140)}`,
    candidateType: "multiSource",
    type: "multiSource",
    multiSourceCandidateKind: "physicalComparison",
    multiSourceCandidateVersion: MULTI_SOURCE_CANDIDATE_VERSION,
    multiSourceDetectionHotfixVersion: MULTI_SOURCE_DETECTION_HOTFIX_VERSION,
    title:
      kind === "WIDE_LONG"
        ? "가로형 원본·정규화 데이터 연결 후보"
        : kind === "CROSS_LONG"
          ? "교차표 원본·정규화 데이터 연결 후보"
          : "원본·가상 테이블 연결 후보",
    description:
      "물리 원본 표와 WIDE_LONG/CROSS_LONG 정규화 표를 함께 추적하는 후보입니다.",
    sourceScope: "virtualLinkedTable",
    sourceTableIds: sourceIds,
    sourceSheetNames: unique([sourceSheetName]),
    sourceTableId: physicalId,
    sourceSheetName,
    virtualTableId: virtualId,
    virtualTableKind: kind,
    sourceTables: [
      {
        sourceTableId: physicalId,
        tableId: physicalTable ? tableIdOf(physicalTable) : physicalId,
        sheetName:
          physicalTable?.sheetName ||
          physicalTable?.name ||
          virtualTable.sheetName ||
          "",
        sourceSheetName,
        role: "physicalSource",
      },
      {
        sourceTableId: physicalId,
        tableId: virtualId,
        sheetName: virtualTable.sheetName || virtualTable.name || "",
        sourceSheetName,
        role: "virtualNormalizedTable",
      },
    ],
    sharedColumns: columnHeaders(virtualTable),
    sharedDimensions: dimensions,
    sharedMetrics: metrics,
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    recipeIds: [
      kind === "WIDE_LONG" ? "wide_time_trend" : "cross_summary",
      "virtual_source_link",
    ],
    confidence: metrics.length ? 0.8 : 0.7,
    priority: kind === "WIDE_LONG" ? 755 : 745,
    reasonCodes: unique([
      "MULTI_SOURCE_CANDIDATE_BUILDER_V1_1",
      "PHYSICAL_VIRTUAL_SOURCE_LINK",
      kind,
      metrics.length ? "HAS_VIRTUAL_METRICS" : "NO_VIRTUAL_METRICS",
      dimensions.length ? "HAS_VIRTUAL_DIMENSIONS" : "NO_VIRTUAL_DIMENSIONS",
    ]),
    recommendationReason:
      "가상 정규화 표가 원본 표에서 파생되었으므로 후보와 산출물에서 두 테이블의 연결 정보를 유지합니다.",
  };
}

function isVirtualTableId(tableId = "") {
  return /#(?:WIDE_LONG|CROSS_LONG)(?:_|$)/i.test(String(tableId || ""));
}

function virtualKindFromId(tableId = "") {
  if (/#WIDE_LONG(?:_|$)/i.test(String(tableId || ""))) return "WIDE_LONG";
  if (/#CROSS_LONG(?:_|$)/i.test(String(tableId || ""))) return "CROSS_LONG";
  return "VIRTUAL";
}

function flattenCandidateItems(candidates = []) {
  const result = [];
  const visit = (candidate) => {
    if (!candidate || typeof candidate !== "object") return;
    result.push(candidate);
    asArray(candidate.candidates).forEach(visit);
    asArray(candidate.recipes).forEach(visit);
    asArray(candidate.sections).forEach((section) => visit(section?.candidate));
    if (candidate.primaryCandidate) visit(candidate.primaryCandidate);
  };
  asArray(candidates).forEach(visit);
  return result;
}

function candidateSourceIds(candidate = {}) {
  return unique([
    ...asArray(candidate.sourceTableIds),
    candidate.sourceTableId,
    candidate.tableId,
    candidate.virtualTableId,
    candidate.source?.tableId,
    candidate.table?.tableId,
  ]);
}

function candidateSourceSheets(candidate = {}) {
  return unique([
    ...asArray(candidate.sourceSheetNames),
    candidate.sourceSheetName,
    candidate.sheetName,
    candidate.table?.sourceSheetName,
    candidate.table?.sheetName,
  ]);
}

function fallbackSourceSheetName(index = 0) {
  const safeIndex = Math.max(0, Number(index) || 0);
  return safeIndex === 0 ? "원본데이터" : `원본데이터${safeIndex + 1}`;
}

function sourceSheetNameForRef({
  id = "",
  refs = [],
  policy = {},
  index = 0,
} = {}) {
  const sourceRef = resolveSourceRefForTableId(id, policy);
  if (sourceRef?.sourceSheetName) return sourceRef.sourceSheetName;

  const matched = refs.find((ref) => ref.ids.includes(id));
  if (matched?.sourceSheetNames?.length) return matched.sourceSheetNames[0];

  const sourceId = stripVirtualTableSuffix(id);
  const matchedSource = refs.find((ref) => ref.ids.includes(sourceId));
  if (matchedSource?.sourceSheetNames?.length)
    return matchedSource.sourceSheetNames[0];

  return fallbackSourceSheetName(index);
}

function collectRefsFromAnalysisCandidates(analysisRecipeCandidates = []) {
  return flattenCandidateItems(analysisRecipeCandidates)
    .map((candidate) => {
      const ids = candidateSourceIds(candidate);
      if (!ids.length) return null;
      return {
        candidate,
        ids,
        sourceSheetNames: candidateSourceSheets(candidate),
        virtualIds: ids.filter(isVirtualTableId),
        physicalIds: ids
          .filter((id) => !isVirtualTableId(id))
          .concat(ids.filter(isVirtualTableId).map(stripVirtualTableSuffix))
          .filter(Boolean),
      };
    })
    .filter(Boolean);
}

function syntheticPhysicalTable({
  id = "",
  sourceSheetName = "",
  index = 0,
} = {}) {
  const safeId = normalizeText(id);
  return {
    tableId: safeId,
    sourceTableId: safeId,
    sheetName: sourceSheetName || fallbackSourceSheetName(index),
    sourceSheetName: sourceSheetName || fallbackSourceSheetName(index),
    tableUsage: { analysisEligible: true, templateEligible: true },
    analysisEligible: true,
    columns: [],
    rows: [],
  };
}

function syntheticVirtualTable({
  virtualId = "",
  physicalId = "",
  sourceSheetName = "",
} = {}) {
  const kind = virtualKindFromId(virtualId);
  return {
    tableId: virtualId,
    sourceTableId: physicalId || stripVirtualTableSuffix(virtualId),
    sheetName: sourceSheetName || "",
    sourceSheetName: sourceSheetName || "",
    tableUsage: { analysisEligible: true, templateEligible: true },
    analysisEligible: true,
    isVirtual: true,
    transformation: {
      type: kind,
      sourceTableId: physicalId || stripVirtualTableSuffix(virtualId),
    },
    columns: [],
    rows: [],
  };
}

function buildPayloadFallbackPhysicalCandidate({
  refs = [],
  policy = {},
} = {}) {
  const physicalIds = unique(
    refs
      .flatMap((ref) => ref.physicalIds)
      .filter((id) => !isVirtualTableId(id)),
  );
  if (physicalIds.length < 2) return null;

  const syntheticTables = physicalIds.map((id, index) =>
    syntheticPhysicalTable({
      id,
      sourceSheetName: sourceSheetNameForRef({ id, refs, policy, index }),
      index,
    }),
  );

  const candidate = makePhysicalMultiSourceCandidate({
    sourceTables: syntheticTables,
    policy,
  });
  if (!candidate) return null;

  return {
    ...candidate,
    candidateId: `multi_source_payload_physical_${physicalIds.map(stableIdKey).join("_").slice(0, 120)}`,
    title: candidate.title || "분석 후보 기반 다중 원본 연결 후보",
    multiSourcePayloadFallbackHotfixVersion:
      MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION,
    reasonCodes: unique([
      ...(candidate.reasonCodes || []),
      "PAYLOAD_FALLBACK_FROM_ANALYSIS_CANDIDATES",
      "MULTIPLE_SOURCE_IDS_FROM_CANDIDATE_PAYLOAD",
    ]),
    recommendationReason:
      candidate.recommendationReason ||
      "분석 후보 payload의 sourceTableIds를 기준으로 여러 원본 표를 연결했습니다.",
  };
}

function buildPayloadFallbackSchemaCandidates({ refs = [], policy = {} } = {}) {
  const physicalIds = unique(
    refs
      .flatMap((ref) => ref.physicalIds)
      .filter((id) => !isVirtualTableId(id)),
  );
  if (physicalIds.length < 2) return [];

  const syntheticTables = physicalIds.map((id, index) =>
    syntheticPhysicalTable({
      id,
      sourceSheetName: sourceSheetNameForRef({ id, refs, policy, index }),
      index,
    }),
  );
  const key = physicalIds.join("|");
  const group = {
    schemaSignature: `payloadFallback:${key}`,
    schemaSignatureHash: hashText(`payloadFallback:${key}`),
    schemaCompatibilityMode: "payloadFallback",
    schemaColumns: [],
    tables: syntheticTables,
    compatibility: {
      ok: true,
      source: "analysisCandidatePayload",
      sharedColumnHeaders: [],
      sharedDimensions: [],
      sharedMetrics: [],
      rowCount: 0,
    },
  };

  const comparison = makeSchemaComparisonCandidate({ group, policy });
  const union = makeSchemaUnionCandidate({ group, policy });
  const individuals = syntheticTables.map((table, index) =>
    makeIndividualSourceCandidate({
      table,
      policy,
      schemaSignature: group.schemaSignature,
      schemaSignatureHash: group.schemaSignatureHash,
      index,
    }),
  );

  return [comparison, union, ...individuals]
    .filter(Boolean)
    .map((candidate) => ({
      ...candidate,
      candidateId:
        candidate.multiSourceCandidateKind === "schemaUnion"
          ? `multi_source_payload_schema_union_${group.schemaSignatureHash}`
          : candidate.multiSourceCandidateKind === "schemaComparison"
            ? `multi_source_payload_schema_compare_${group.schemaSignatureHash}`
            : candidate.candidateId,
      multiSourcePayloadFallbackHotfixVersion:
        MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION,
      reasonCodes: unique([
        ...(candidate.reasonCodes || []),
        "PAYLOAD_FALLBACK_FROM_ANALYSIS_CANDIDATES",
        "PAYLOAD_FALLBACK_SCHEMA_GROUP",
      ]),
      recommendationReason:
        candidate.recommendationReason ||
        "분석 후보 payload의 sourceTableIds를 기준으로 schemaUnion 후보를 복구했습니다.",
    }));
}

function buildPayloadFallbackVirtualCandidates({
  refs = [],
  policy = {},
} = {}) {
  const virtualIds = unique(refs.flatMap((ref) => ref.virtualIds));
  return virtualIds
    .map((virtualId) => {
      const physicalId = stripVirtualTableSuffix(virtualId);
      if (!physicalId || physicalId === virtualId) return null;
      const sourceSheetName = sourceSheetNameForRef({
        id: virtualId,
        refs,
        policy,
      });
      const physicalTable = syntheticPhysicalTable({
        id: physicalId,
        sourceSheetName,
      });
      const virtualTable = syntheticVirtualTable({
        virtualId,
        physicalId,
        sourceSheetName,
      });
      const candidate = makeVirtualLinkCandidate({
        virtualTable,
        physicalTable,
        policy,
      });
      if (!candidate) return null;
      return {
        ...candidate,
        candidateId: `multi_source_payload_virtual_${stableIdKey(virtualId).slice(0, 140)}`,
        multiSourcePayloadFallbackHotfixVersion:
          MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION,
        reasonCodes: unique([
          ...(candidate.reasonCodes || []),
          "PAYLOAD_FALLBACK_FROM_ANALYSIS_CANDIDATES",
          "VIRTUAL_SOURCE_ID_FROM_CANDIDATE_PAYLOAD",
        ]),
        recommendationReason:
          candidate.recommendationReason ||
          "분석 후보 payload에 포함된 virtual table id를 기준으로 원본·정규화 표 연결 후보를 복구했습니다.",
      };
    })
    .filter(Boolean);
}

function attachMultiSourceDiagnostics(candidates = [], diagnostics = {}) {
  try {
    Object.defineProperty(candidates, "diagnostics", {
      value: diagnostics,
      enumerable: false,
      configurable: true,
    });
  } catch (_) {
    candidates.diagnostics = diagnostics;
  }
  return candidates;
}

function buildMultiSourceCandidates({
  normalizedQueryTables = [],
  sourceTablePolicy = {},
  analysisRecipeCandidates = [],
} = {}) {
  const tables = asArray(normalizedQueryTables);
  const sourceTables = uniqueTables([
    ...sourceTablesFromPolicy(sourceTablePolicy),
    ...physicalTablesFromNormalized(tables).filter(isAnalysisEligible),
  ]);
  const refsFromAnalysisCandidates = collectRefsFromAnalysisCandidates(
    analysisRecipeCandidates,
  );
  const candidates = [];

  const physicalCandidate = makePhysicalMultiSourceCandidate({
    sourceTables,
    policy: sourceTablePolicy,
  });
  if (physicalCandidate) {
    candidates.push({
      ...physicalCandidate,
      reasonCodes: unique([
        ...(physicalCandidate.reasonCodes || []),
        "MULTI_SOURCE_CANDIDATE_BUILDER_V1_3",
      ]),
    });
  }

  candidates.push(
    ...buildSchemaGroupCandidates({
      sourceTables,
      policy: sourceTablePolicy,
      analysisRecipeCandidates,
    }),
  );

  const physicalBySourceId = new Map(
    sourceTables.map((table) => [sourceTableIdOf(table), table]),
  );
  const physicalByTableId = new Map(
    sourceTables.map((table) => [tableIdOf(table), table]),
  );

  tables.filter(isVirtualTable).forEach((virtualTable) => {
    const physicalId = sourceTableIdOf(virtualTable);
    const candidate = makeVirtualLinkCandidate({
      virtualTable,
      physicalTable:
        physicalBySourceId.get(physicalId) ||
        physicalByTableId.get(physicalId) ||
        null,
      policy: sourceTablePolicy,
    });
    if (candidate) candidates.push(candidate);
  });

  const payloadPhysicalCandidate = buildPayloadFallbackPhysicalCandidate({
    refs: refsFromAnalysisCandidates,
    policy: sourceTablePolicy,
  });
  if (payloadPhysicalCandidate) candidates.push(payloadPhysicalCandidate);

  if (
    refsFromAnalysisCandidates.length &&
    !candidates.some(
      (candidate) => candidate.multiSourceCandidateKind === "schemaUnion",
    )
  ) {
    candidates.push(
      ...buildPayloadFallbackSchemaCandidates({
        refs: refsFromAnalysisCandidates,
        policy: sourceTablePolicy,
      }),
    );
  }

  candidates.push(
    ...buildPayloadFallbackVirtualCandidates({
      refs: refsFromAnalysisCandidates,
      policy: sourceTablePolicy,
    }),
  );

  const byId = new Map();
  for (const candidate of candidates) {
    if (!candidate || !candidate.candidateId) continue;
    if (byId.has(candidate.candidateId)) continue;
    byId.set(candidate.candidateId, {
      ...candidate,
      linkedAnalysisCandidateCount: asArray(analysisRecipeCandidates).filter(
        (item) => {
          const ids = asArray(item.sourceTableIds)
            .concat(item.sourceTableId, item.tableId, item.virtualTableId)
            .filter(Boolean);
          return ids.some((id) =>
            asArray(candidate.sourceTableIds).includes(id),
          );
        },
      ).length,
    });
  }

  const result = [...byId.values()];
  return attachMultiSourceDiagnostics(result, {
    version: MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION,
    inputTableCount: tables.length,
    sourcePolicyTableCount:
      sourceTablePolicy.counts?.sourceTableCount || sourceTables.length || 0,
    sourceTableCountFromPolicyOrTables: sourceTables.length,
    analysisRecipeCandidateCount: asArray(analysisRecipeCandidates).length,
    payloadRefCount: refsFromAnalysisCandidates.length,
    virtualRefsFromRecipes: unique(
      refsFromAnalysisCandidates.flatMap((ref) => ref.virtualIds),
    ),
    physicalRefsFromRecipes: unique(
      refsFromAnalysisCandidates.flatMap((ref) => ref.physicalIds),
    ),
    generatedFrom: unique(
      result.flatMap((candidate) => candidate.reasonCodes || []),
    ).filter((code) => /FALLBACK|PHYSICAL|VIRTUAL|MULTI_SOURCE/i.test(code)),
    schemaGroupVersion: MULTI_SOURCE_SCHEMA_GROUP_VERSION,
    schemaUnionVersion: MULTI_SOURCE_SCHEMA_UNION_VERSION,
    individualSourceVersion: MULTI_SOURCE_INDIVIDUAL_SOURCE_VERSION,
    exactSchemaGroupCount: groupTablesBySchemaSignature(sourceTables).length,
    compatibleSchemaGroupCount: groupTablesByCompatibleSchema(
      sourceTables,
      groupTablesBySchemaSignature(sourceTables),
    ).length,
    schemaGroupCount:
      groupTablesBySchemaSignature(sourceTables).length +
      groupTablesByCompatibleSchema(
        sourceTables,
        groupTablesBySchemaSignature(sourceTables),
      ).length,
    schemaUnionCandidateCount: result.filter(
      (candidate) => candidate.multiSourceCandidateKind === "schemaUnion",
    ).length,
    schemaComparisonCandidateCount: result.filter(
      (candidate) => candidate.multiSourceCandidateKind === "schemaComparison",
    ).length,
    individualSourceCandidateCount: result.filter(
      (candidate) => candidate.multiSourceCandidateKind === "individualSource",
    ).length,
    outputCandidateCount: result.length,
  });
}

module.exports = {
  MULTI_SOURCE_CANDIDATE_VERSION,
  MULTI_SOURCE_DETECTION_HOTFIX_VERSION,
  MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION,
  MULTI_SOURCE_SCHEMA_GROUP_VERSION,
  MULTI_SOURCE_SCHEMA_UNION_VERSION,
  MULTI_SOURCE_INDIVIDUAL_SOURCE_VERSION,
  UNION_SOURCE_DIMENSION_HEADER,
  buildMultiSourceCandidates,
};
