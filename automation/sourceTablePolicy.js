const SOURCE_TABLE_POLICY_VERSION = "source_table_policy_v1";
const DEFAULT_SOURCE_SHEET_BASE_NAME = "원본데이터";
const VIRTUAL_TABLE_SUFFIX_RE = /#(?:WIDE_LONG|CROSS_LONG)(?:_[^#]*)?$/i;
const VIRTUAL_TABLE_TOKEN_RE = /#(?:WIDE_LONG|CROSS_LONG)(?:_|$)/i;

function normalizeText(value = "") {
  return String(value || "").trim();
}

function firstNonEmpty(...values) {
  for (const value of values) {
    const text = normalizeText(value);
    if (text) return text;
  }
  return "";
}

function getCanonicalTableId(table = {}) {
  if (typeof table === "string") return normalizeText(table);
  return firstNonEmpty(
    table.tableId,
    table.id,
    table.queryTableId,
    table.normalizedTableId,
    table.meta?.tableId,
    table.metadata?.tableId,
    table.diagnostics?.tableId,
  );
}

function getTransformationType(table = {}) {
  return firstNonEmpty(
    table.transformation?.type,
    table.transformationType,
    table.normalizationType,
    table.normalizedType,
    table.virtualTableKind,
    table.tableKind,
    table.tableType,
    table.type,
    table.sourceScope,
  );
}

function sourceSheetNameForTableIndex(
  index = 0,
  total = 1,
  baseName = DEFAULT_SOURCE_SHEET_BASE_NAME,
) {
  const safeIndex = Math.max(0, Number(index) || 0);
  const safeBaseName =
    normalizeText(baseName) || DEFAULT_SOURCE_SHEET_BASE_NAME;
  return safeIndex === 0 ? safeBaseName : `${safeBaseName}${safeIndex + 1}`;
}

function isVirtualQueryTable(table = {}) {
  const id = getCanonicalTableId(table);
  const sourceId = firstNonEmpty(
    table.sourceTableId,
    table.transformation?.sourceTableId,
  );
  const kind = getTransformationType(table);
  return (
    table.isVirtual === true ||
    table.virtual === true ||
    Boolean(table.transformation) ||
    VIRTUAL_TABLE_TOKEN_RE.test(id) ||
    VIRTUAL_TABLE_TOKEN_RE.test(sourceId) ||
    /(?:WIDE_LONG|CROSS_LONG)/i.test(kind)
  );
}

function stripVirtualTableSuffix(tableId = "") {
  return normalizeText(tableId).replace(VIRTUAL_TABLE_SUFFIX_RE, "");
}

function getSourceTableId(table = {}) {
  const tableId = getCanonicalTableId(table);
  return firstNonEmpty(
    table.sourceTableId,
    table.physicalTableId,
    table.parentTableId,
    table.originTableId,
    table.transformation?.sourceTableId,
    table.transformation?.source?.tableId,
    table.transformation?.source?.id,
    table.source?.tableId,
    table.source?.id,
    stripVirtualTableSuffix(tableId),
    tableId,
  );
}

function getTableUsage(table = {}) {
  return table.tableUsage || table.diagnostics?.tableUsage || {};
}

function isAnalysisEligibleSourceTable(table = {}) {
  return getTableUsage(table).analysisEligible === true;
}

function isTemplateEligibleSourceTable(table = {}) {
  return getTableUsage(table).templateEligible === true;
}

function tableIsNotExplicitlyIneligible(table = {}) {
  return getTableUsage(table).analysisEligible !== false;
}

function tableIdentity(table = {}) {
  return getSourceTableId(table) || table.tableId || table.sheetName || "";
}

function tableMatchesId(table = {}, id = "") {
  const target = normalizeText(id);
  if (!target) return false;
  const sourceId = normalizeText(getSourceTableId(table));
  const tableId = normalizeText(getCanonicalTableId(table));
  return (
    tableId === target ||
    sourceId === target ||
    sourceId === stripVirtualTableSuffix(target)
  );
}

function uniqueTablesBySourceId(tables = []) {
  const seen = new Set();
  const result = [];

  for (const table of Array.isArray(tables) ? tables : []) {
    if (!table) continue;
    const key = tableIdentity(table);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(table);
  }

  return result;
}

function chooseSourceTables({ physicalTables = [], allTables = [] } = {}) {
  const physical = uniqueTablesBySourceId(
    physicalTables.filter((table) => !isVirtualQueryTable(table)),
  );
  const fallbackNonVirtual = uniqueTablesBySourceId(
    allTables.filter((table) => !isVirtualQueryTable(table)),
  );
  const candidatePool = physical.length ? physical : fallbackNonVirtual;
  const eligible = candidatePool.filter(isAnalysisEligibleSourceTable);

  if (eligible.length) return eligible;

  return candidatePool.filter(tableIsNotExplicitlyIneligible);
}

function pickPrimarySourceTable({
  sourceTables = [],
  allTables = [],
  preferredTableId = "",
} = {}) {
  const preferred = preferredTableId
    ? sourceTables.find((table) => tableMatchesId(table, preferredTableId)) ||
      allTables.find((table) => tableMatchesId(table, preferredTableId))
    : null;

  return (
    preferred ||
    sourceTables.find((table) => table?.isPrimary === true) ||
    sourceTables.find(isAnalysisEligibleSourceTable) ||
    allTables.find(
      (table) =>
        table?.isPrimary === true && tableIsNotExplicitlyIneligible(table),
    ) ||
    allTables.find(isAnalysisEligibleSourceTable) ||
    sourceTables[0] ||
    null
  );
}

function makeSourceTableEntry({
  table,
  index = 0,
  total = 1,
  sourceSheetBaseName = DEFAULT_SOURCE_SHEET_BASE_NAME,
} = {}) {
  const sourceSheetName = sourceSheetNameForTableIndex(
    index,
    total,
    sourceSheetBaseName,
  );

  return {
    sourceTableIndex: index,
    sourceTableId: getSourceTableId(table),
    tableId: getCanonicalTableId(table),
    sheetName: table?.sheetName || table?.name || table?.sheet || "",
    sourceSheetName,
    sourceScope: "singleTable",
    table,
  };
}

function buildSourceTablePolicy({
  tables = [],
  normalizedQueryTables = [],
  preferredTableId = "",
  sourceSheetBaseName = DEFAULT_SOURCE_SHEET_BASE_NAME,
} = {}) {
  const inputTables = Array.isArray(tables) ? tables.filter(Boolean) : [];
  const normalizedTables = Array.isArray(normalizedQueryTables)
    ? normalizedQueryTables.filter(Boolean)
    : [];
  const physicalTables = inputTables.filter(
    (table) => !isVirtualQueryTable(table),
  );
  const allTables = inputTables.length
    ? inputTables
    : normalizedTables.length
      ? normalizedTables
      : [];
  const virtualTables = normalizedTables.filter(isVirtualQueryTable);
  const virtualAnalysisEligibleTables = virtualTables.filter(
    isAnalysisEligibleSourceTable,
  );
  const sourceTables = chooseSourceTables({
    physicalTables,
    allTables,
  });
  const primaryTable = pickPrimarySourceTable({
    sourceTables,
    allTables,
    preferredTableId,
  });
  const sourceTableEntries = sourceTables.map((table, index) =>
    makeSourceTableEntry({
      table,
      index,
      total: sourceTables.length,
      sourceSheetBaseName,
    }),
  );
  const primaryEntry =
    sourceTableEntries.find((entry) => entry.table === primaryTable) ||
    sourceTableEntries.find((entry) =>
      primaryTable
        ? tableMatchesId(entry.table, getSourceTableId(primaryTable))
        : false,
    ) ||
    null;

  const sourceSheetByTableId = {};
  for (const entry of sourceTableEntries) {
    [entry.tableId, entry.sourceTableId].filter(Boolean).forEach((id) => {
      sourceSheetByTableId[id] = entry.sourceSheetName;
      sourceSheetByTableId[stripVirtualTableSuffix(id)] = entry.sourceSheetName;
    });
  }

  return {
    version: SOURCE_TABLE_POLICY_VERSION,
    sourceSheetBaseName,
    sourceMode: "physical-analysis-eligible-tables",
    sourceScope: sourceTables.length > 1 ? "multiTable" : "singleTable",
    primaryStrategy:
      preferredTableId &&
      primaryTable &&
      tableMatchesId(primaryTable, preferredTableId)
        ? "preferredTableId"
        : primaryTable?.isPrimary
          ? "isPrimary"
          : primaryTable && isAnalysisEligibleSourceTable(primaryTable)
            ? "analysisEligible"
            : primaryTable
              ? "fallback"
              : "none",
    preferredTableId: preferredTableId || "",
    sourceTables,
    sourceTableEntries,
    primaryTable,
    primarySourceTableId:
      primaryEntry?.sourceTableId || getSourceTableId(primaryTable || {}),
    primarySourceSheetName:
      primaryEntry?.sourceSheetName ||
      (primaryTable
        ? sourceSheetNameForTableIndex(0, 1, sourceSheetBaseName)
        : ""),
    sourceSheetByTableId,
    counts: {
      physicalTableCount: physicalTables.length,
      normalizedTableCount: normalizedTables.length,
      virtualTableCount: virtualTables.length,
      virtualAnalysisEligibleCount: virtualAnalysisEligibleTables.length,
      sourceTableCount: sourceTables.length,
      analysisEligibleSourceTableCount: sourceTables.filter(
        isAnalysisEligibleSourceTable,
      ).length,
      primaryCount: sourceTables.filter((table) => table?.isPrimary === true)
        .length,
    },
  };
}

function summarizeSourceTablePolicy(policy = {}) {
  return {
    version: policy.version || SOURCE_TABLE_POLICY_VERSION,
    sourceMode: policy.sourceMode || "",
    sourceScope: policy.sourceScope || "",
    primaryStrategy: policy.primaryStrategy || "",
    primarySourceTableId: policy.primarySourceTableId || "",
    primarySourceSheetName: policy.primarySourceSheetName || "",
    counts: policy.counts || {},
    sourceTables: (policy.sourceTableEntries || []).map((entry) => ({
      sourceTableIndex: entry.sourceTableIndex,
      sourceTableId: entry.sourceTableId,
      tableId: entry.tableId,
      sheetName: entry.sheetName,
      sourceSheetName: entry.sourceSheetName,
      sourceScope: entry.sourceScope,
    })),
  };
}

function resolveSourceRefForTableId(tableId = "", policy = {}) {
  const id = normalizeText(tableId);
  const sourceId = stripVirtualTableSuffix(id);
  const entries = policy.sourceTableEntries || [];

  const entry = entries.find(
    (item) =>
      item.tableId === id ||
      item.sourceTableId === id ||
      item.tableId === sourceId ||
      item.sourceTableId === sourceId,
  );

  if (entry) {
    return {
      sourceTablePolicyVersion: policy.version || SOURCE_TABLE_POLICY_VERSION,
      sourceScope: entry.sourceScope || "singleTable",
      sourceTableIndex: entry.sourceTableIndex,
      sourceTableId: entry.sourceTableId || sourceId,
      sourceSheetName: entry.sourceSheetName,
      tableId: id,
      virtualTableId: isVirtualQueryTable({ tableId }) ? id : "",
    };
  }

  return {
    sourceTablePolicyVersion: policy.version || SOURCE_TABLE_POLICY_VERSION,
    sourceScope: "singleTable",
    sourceTableIndex: -1,
    sourceTableId: sourceId,
    sourceSheetName: policy.sourceSheetByTableId?.[sourceId] || "",
    tableId: id,
    virtualTableId: VIRTUAL_TABLE_TOKEN_RE.test(id) ? id : "",
  };
}

function resolveSourceRefForTable(table = {}, policy = {}) {
  return resolveSourceRefForTableId(
    getCanonicalTableId(table) || getSourceTableId(table),
    policy,
  );
}

function enrichCandidateWithSourceTablePolicy(candidate = {}, policy = {}) {
  if (!candidate || typeof candidate !== "object") return candidate;
  const tableId =
    candidate.tableId ||
    candidate.sourceTableId ||
    candidate.id ||
    candidate.queryTableId ||
    "";
  const sourceRef = resolveSourceRefForTableId(tableId, policy);

  return {
    ...candidate,
    sourceTablePolicyVersion: sourceRef.sourceTablePolicyVersion,
    sourceScope: candidate.sourceScope || sourceRef.sourceScope,
    sourceTableId: candidate.sourceTableId || sourceRef.sourceTableId,
    sourceSheetName: candidate.sourceSheetName || sourceRef.sourceSheetName,
    sourceTableIndex:
      candidate.sourceTableIndex != null
        ? candidate.sourceTableIndex
        : sourceRef.sourceTableIndex,
    virtualTableId: candidate.virtualTableId || sourceRef.virtualTableId || "",
  };
}

function enrichCandidateListWithSourceTablePolicy(
  candidates = [],
  policy = {},
) {
  return (Array.isArray(candidates) ? candidates : []).map((candidate) =>
    enrichCandidateWithSourceTablePolicy(candidate, policy),
  );
}

module.exports = {
  SOURCE_TABLE_POLICY_VERSION,
  DEFAULT_SOURCE_SHEET_BASE_NAME,
  VIRTUAL_TABLE_SUFFIX_RE,
  VIRTUAL_TABLE_TOKEN_RE,
  sourceSheetNameForTableIndex,
  isVirtualQueryTable,
  stripVirtualTableSuffix,
  getCanonicalTableId,
  getSourceTableId,
  getTableUsage,
  isAnalysisEligibleSourceTable,
  isTemplateEligibleSourceTable,
  buildSourceTablePolicy,
  summarizeSourceTablePolicy,
  resolveSourceRefForTableId,
  resolveSourceRefForTable,
  enrichCandidateWithSourceTablePolicy,
  enrichCandidateListWithSourceTablePolicy,
};
