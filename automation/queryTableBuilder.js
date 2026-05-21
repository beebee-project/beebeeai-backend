const XLSX = require("xlsx");

function normalizeKey(header = "", fallback = "col") {
  const base = String(header || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[^\p{Letter}\p{Number}_]+/gu, "_")
    .replace(/^_+|_+$/g, "");

  return base || fallback;
}

function uniqueKeys(columns = []) {
  const seen = new Map();

  return columns.map((col, idx) => {
    const base = normalizeKey(col.header, `col_${idx + 1}`);
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);

    return {
      ...col,
      key: count ? `${base}_${count + 1}` : base,
    };
  });
}

function cellValue(row, colIndex1Based) {
  return row?.[colIndex1Based - 1] ?? null;
}

function coerceValue(value, type) {
  if (value == null || value === "") return null;

  if (type === "number") {
    const n = Number(String(value).replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : value;
  }

  if (type === "date") {
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    const s = String(value).trim();
    return s || null;
  }

  return String(value).trim();
}

function inferColumnType(meta = {}) {
  const profileType = String(meta.profileType || "").toLowerCase();
  const dominantType = String(meta.dominantType || "").toLowerCase();

  if (profileType === "number" || dominantType === "number") return "number";
  if (profileType === "date" || dominantType === "date") return "date";
  if (profileType === "category") return "category";
  return "text";
}

function buildFallbackBlockFromMeta(sheetName, sheetInfo = {}) {
  const meta = sheetInfo.metaData || {};
  const entries = Object.entries(meta);

  if (!entries.length) return null;

  entries.sort((a, b) => {
    const ai = Number(a[1]?.columnIndex || 9999);
    const bi = Number(b[1]?.columnIndex || 9999);
    return ai - bi;
  });

  const firstMeta = entries[0]?.[1] || {};
  const headerRow = Number(firstMeta.headerRow || 1);
  const dataStartRow = Number(
    firstMeta.startRow || sheetInfo.startRow || headerRow + 1,
  );
  const dataEndRow = Number(
    sheetInfo.lastDataRow || firstMeta.lastRow || dataStartRow,
  );

  const columns = entries.map(([header, m], idx) => ({
    header,
    originalHeader: header,
    columnIndex: Number(m.columnIndex || idx + 1),
    columnLetter: m.columnLetter,
  }));

  if (!columns.length || dataEndRow < dataStartRow) return null;

  const startCol = columns[0].columnLetter || "A";
  const endCol = columns[columns.length - 1].columnLetter || startCol;

  return {
    tableId: `${sheetName}#AUTO`,
    sheetName,
    headerRow,
    headerRows: [headerRow],
    hasMergedHeader: false,
    dataStartRow,
    dataEndRow,
    startCol,
    endCol,
    range: `'${sheetName}'!${startCol}${headerRow}:${endCol}${dataEndRow}`,
    dataRange: `'${sheetName}'!${startCol}${dataStartRow}:${endCol}${dataEndRow}`,
    columns,
    score: 0,
    isFallback: true,
  };
}

function buildQueryTablesFromWorkbook(workbook, allSheetsData = {}) {
  const tables = [];

  for (const [sheetName, sheetInfo] of Object.entries(allSheetsData || {})) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: true,
      defval: null,
    });

    let blocks = Array.isArray(sheetInfo.tableBlocks)
      ? [...sheetInfo.tableBlocks]
      : [];

    if (!blocks.length) {
      const fallbackBlock = buildFallbackBlockFromMeta(sheetName, sheetInfo);
      if (fallbackBlock) blocks.push(fallbackBlock);
    }

    for (const block of blocks) {
      const rawColumns = (block.columns || []).map((c, idx) => {
        const meta = sheetInfo.metaData?.[c.header] || {};
        return {
          header: c.header,
          originalHeader: c.originalHeader || c.header,
          columnIndex: c.columnIndex || idx + 1,
          columnLetter: c.columnLetter,
          type: inferColumnType(meta),
          sampleValues: meta.sampleValues || [],
          profileType: meta.profileType || null,
          dominantType: meta.dominantType || null,
          uniqueCount: meta.uniqueCount ?? null,
          uniqueRatio: meta.uniqueRatio ?? null,
        };
      });

      const columns = uniqueKeys(rawColumns);
      const data = [];

      for (let r = block.dataStartRow; r <= block.dataEndRow; r += 1) {
        const row = rows[r - 1] || [];
        const obj = {};

        let nonEmpty = 0;

        for (const col of columns) {
          const raw = cellValue(row, col.columnIndex);
          const value = coerceValue(raw, col.type);
          obj[col.key] = value;
          if (value !== null && value !== "") nonEmpty += 1;
        }

        if (nonEmpty > 0) data.push(obj);
      }

      tables.push({
        tableId: block.tableId,
        isFallback: !!block.isFallback,
        tableName: normalizeKey(block.tableId.replace("#", "_"), "table"),
        sheetName,
        range: block.range,
        dataRange: block.dataRange,
        headerRow: block.headerRow,
        dataStartRow: block.dataStartRow,
        dataEndRow: block.dataEndRow,
        rowCount: data.length,
        columns,
        rows: data,
      });
    }
  }

  return tables;
}

module.exports = { buildQueryTablesFromWorkbook };
