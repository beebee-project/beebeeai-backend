const {
  isVirtualQueryTable,
  stripVirtualTableSuffix,
  resolveSourceRefForTableId,
  getSourceTableId,
  getCanonicalTableId,
  getTableUsage,
} = require("./sourceTablePolicy");

const MULTI_SOURCE_CANDIDATE_VERSION = "multi_source_candidate_builder_v1_2";
const MULTI_SOURCE_DETECTION_HOTFIX_VERSION =
  "multi_source_candidate_detection_hotfix_v1";
const MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION =
  "multi_source_candidate_payload_fallback_hotfix_v1";

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
    candidateId: `multi_source_physical_${sourceIds.map(normalizeKey).join("_").slice(0, 120)}`,
    candidateType: "multiSource",
    type: "multiSource",
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
    candidateId: `multi_source_virtual_${normalizeKey(virtualId).slice(0, 140)}`,
    candidateType: "multiSource",
    type: "multiSource",
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
    candidateId: `multi_source_payload_physical_${physicalIds.map(normalizeKey).join("_").slice(0, 120)}`,
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
        candidateId: `multi_source_payload_virtual_${normalizeKey(virtualId).slice(0, 140)}`,
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
  if (physicalCandidate) candidates.push(physicalCandidate);

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
    outputCandidateCount: result.length,
  });
}

module.exports = {
  MULTI_SOURCE_CANDIDATE_VERSION,
  MULTI_SOURCE_DETECTION_HOTFIX_VERSION,
  MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION,
  buildMultiSourceCandidates,
};
