function findColumnByHeader(queryContext, hint = "") {
  const normalizedHint = normalizeHeaderText(hint);
  if (!normalizedHint) return null;

  const columns = extractColumns(queryContext);

  return (
    columns.find((c) => c.normalizedHeader === normalizedHint) ||
    columns.find((c) => c.normalizedHeader.includes(normalizedHint)) ||
    columns.find((c) => normalizedHint.includes(c.normalizedHeader)) ||
    null
  );
}

function normalizeHeaderText(value = "") {
  return String(value || "")
    .replace(/\([^)]*\)/g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function pushColumn(result, table, header, meta = {}) {
  if (!header) return;

  result.push({
    tableId: table?.tableId || table?.id || null,
    sheetName: table?.sheetName || table?.sheet || meta.sheetName || null,
    header,
    normalizedHeader: normalizeHeaderText(header),
    columnLetter: meta.columnLetter || meta.letter || null,
    columnIndex: meta.columnIndex || meta.index || null,
    dominantType: meta.dominantType || meta.type || meta.profileType || null,
    role: meta.role || meta.clusterRole || meta.inferredRole || null,
    canonicalKey: meta.canonicalKey || meta.clusterCandidate || null,
    uniqueValues: meta.uniqueValues || [],
  });
}

function extractColumns(queryContext = {}) {
  const result = [];

  const tables =
    queryContext?.normalizedQueryTables || queryContext?.tables || [];

  for (const table of tables) {
    for (const col of table.columns || []) {
      const header =
        col.header || col.originalHeader || col.name || col.key || null;
      pushColumn(result, table, header, col);
    }
  }

  const sheets =
    queryContext?.allSheetsData?.allSheetsData ||
    queryContext?.allSheetsData ||
    queryContext?.allSheetsData?.allSheetsData?.allSheetsData ||
    {};

  for (const [sheetName, sheet] of Object.entries(sheets)) {
    const metaData = sheet?.metaData || {};
    for (const [header, meta] of Object.entries(metaData)) {
      if (!meta?.columnLetter) continue;
      if (!meta?.inferredRole && !meta?.canonicalKey) continue;

      pushColumn(result, { sheetName }, header, { ...meta, sheetName });
    }
  }

  return result;
}

function findColumnsByRole(queryContext, role) {
  return extractColumns(queryContext).filter((c) => c.role === role);
}

function findColumnsByType(queryContext, type) {
  return extractColumns(queryContext).filter((c) => c.dominantType === type);
}

module.exports = {
  extractColumns,
  findColumnByHeader,
  findColumnsByRole,
  findColumnsByType,
  normalizeHeaderText,
};
