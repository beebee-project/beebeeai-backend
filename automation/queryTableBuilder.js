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

function colLetter(n) {
  let s = "";
  let x = Number(n || 0);

  while (x > 0) {
    const m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }

  return s || "A";
}

function isEmptyCell(v) {
  return v == null || String(v).trim() === "";
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

const TABLE_USAGE_QUALITY_VERSION = "table_usage_quality_v1";

function normalizeUsageText(value = "") {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "")
    .trim();
}

function includesAnyNormalized(value = "", tokens = []) {
  const normalized = normalizeUsageText(value);
  if (!normalized) return false;
  return tokens.some((token) => normalized.includes(normalizeUsageText(token)));
}

const META_OR_INSTRUCTION_SHEET_TOKENS = [
  "메타",
  "metadata",
  "meta",
  "설명",
  "테스트설명",
  "사용안내",
  "안내",
  "readme",
  "guide",
  "instruction",
];

const META_OR_INSTRUCTION_CELL_TOKENS = [
  "통계표ID",
  "조회기간",
  "자료다운일자",
  "자료유형",
  "원본파일",
  "작성기관",
  "작성일",
  "생성일",
  "데이터기준",
  "자료기준",
  "단위",
  "출처",
  "통계표URL",
  "URL",
  "비고",
  "주석",
  "설명",
  "테스트목적",
  "테스트설명",
];

const FORM_LIKE_TOKENS = [
  "입력값",
  "입력란",
  "확인",
  "검토",
  "결재",
  "승인",
  "담당자",
  "서명",
  "날인",
  "체크",
  "양식",
  "□",
  "☐",
];

function collectTableUsageTexts(table = {}) {
  const headers = (table.columns || []).map(
    (c) => c.header || c.originalHeader || "",
  );
  const samples = (table.columns || [])
    .flatMap((c) => c.sampleValues || [])
    .filter((v) => v != null)
    .slice(0, 60);
  const rowValues = (table.rows || [])
    .slice(0, 5)
    .flatMap((row) => Object.values(row || {}))
    .filter((v) => v != null)
    .slice(0, 60);

  return {
    headers,
    samples,
    rowValues,
    all: [
      table.sheetName,
      table.tableTitle,
      table.tableName,
      ...headers,
      ...samples,
      ...rowValues,
    ].map((v) => String(v ?? "")),
  };
}

function countTextsWithTokens(texts = [], tokens = []) {
  return texts.filter((value) => includesAnyNormalized(value, tokens)).length;
}

function hasUrlLikeValue(texts = []) {
  return texts.some((value) =>
    /https?:\/\/|www\.|\.go\.kr|\.or\.kr|\.com/i.test(String(value || "")),
  );
}

function scoreTableUsageQuality(table = {}) {
  const rowCount = Number(table.rowCount || 0);
  const colCount = Array.isArray(table.columns) ? table.columns.length : 0;
  const confidence = Number(table.confidence || 0);
  const numericColumns = (table.columns || []).filter(
    (c) => c.type === "number",
  ).length;
  const dateColumns = (table.columns || []).filter(
    (c) => c.type === "date",
  ).length;
  const categoryColumns = (table.columns || []).filter(
    (c) => c.type === "category",
  ).length;
  const textColumns = (table.columns || []).filter(
    (c) => c.type === "text",
  ).length;
  const rangeRows = Math.max(0, Number(table.rawDataRowCount || rowCount || 0));
  const excludedCount = Array.isArray(table.excludedRows)
    ? table.excludedRows.length
    : 0;
  const texts = collectTableUsageTexts(table);
  const headerMetaTokens = countTextsWithTokens(
    texts.headers,
    META_OR_INSTRUCTION_CELL_TOKENS,
  );
  const allMetaTokens = countTextsWithTokens(
    texts.all,
    META_OR_INSTRUCTION_CELL_TOKENS,
  );
  const formTokens = countTextsWithTokens(texts.all, FORM_LIKE_TOKENS);
  const sheetLooksMetaOrInstruction = includesAnyNormalized(
    table.sheetName || "",
    META_OR_INSTRUCTION_SHEET_TOKENS,
  );
  const hasUrl = hasUrlLikeValue(texts.all);

  const broadTableEvidence =
    (rowCount >= 5 && colCount >= 2) ||
    (rowCount >= 3 && colCount >= 3 && numericColumns >= 1) ||
    (rowCount >= 2 && colCount >= 4 && numericColumns >= 2) ||
    (rowCount >= 2 && colCount >= 4 && dateColumns >= 1 && numericColumns >= 1);

  const metadataLikeBlock = Boolean(
    !broadTableEvidence &&
    (sheetLooksMetaOrInstruction ||
      hasUrl ||
      headerMetaTokens >= 2 ||
      allMetaTokens >= 3 ||
      (rowCount <= 2 && colCount <= 6 && allMetaTokens >= 2)),
  );

  const formLikeBlock = Boolean(
    !broadTableEvidence &&
    (formTokens >= 2 ||
      (rowCount <= 3 && colCount <= 5 && formTokens >= 1 && textColumns >= 2)),
  );

  const verySmallWeakBlock = Boolean(
    !broadTableEvidence &&
    rowCount <= 1 &&
    (numericColumns + dateColumns === 0 || allMetaTokens >= 1),
  );

  const lowConfidenceWeakBlock = Boolean(
    !broadTableEvidence && confidence < 45 && rowCount <= 2,
  );

  const reasons = [];
  if (sheetLooksMetaOrInstruction) reasons.push("META_OR_INSTRUCTION_SHEET");
  if (metadataLikeBlock) reasons.push("META_OR_INSTRUCTION_TABLE");
  if (formLikeBlock) reasons.push("FORM_LIKE_TABLE_BLOCK");
  if (verySmallWeakBlock) reasons.push("VERY_SMALL_WEAK_TABLE_BLOCK");
  if (lowConfidenceWeakBlock) reasons.push("LOW_CONFIDENCE_WEAK_TABLE_BLOCK");

  const analysisEligible = Boolean(
    rowCount > 0 &&
    colCount >= 2 &&
    !metadataLikeBlock &&
    !formLikeBlock &&
    !verySmallWeakBlock &&
    !lowConfidenceWeakBlock,
  );

  const templateEligible = Boolean(
    analysisEligible &&
    (broadTableEvidence ||
      rowCount >= 3 ||
      numericColumns >= 2 ||
      categoryColumns >= 1),
  );

  return {
    version: TABLE_USAGE_QUALITY_VERSION,
    queryable: rowCount > 0 && colCount >= 2,
    analysisEligible,
    templateEligible,
    reasons: reasons.length
      ? reasons
      : analysisEligible
        ? ["TABLE_LIKE_STRUCTURE"]
        : ["LOW_TABLE_USAGE_CONFIDENCE"],
    metrics: {
      rowCount,
      rawDataRowCount: rangeRows,
      columnCount: colCount,
      numericColumnCount: numericColumns,
      dateColumnCount: dateColumns,
      categoryColumnCount: categoryColumns,
      textColumnCount: textColumns,
      excludedRowCount: excludedCount,
      confidence,
      headerMetaTokenCount: headerMetaTokens,
      allMetaTokenCount: allMetaTokens,
      formTokenCount: formTokens,
      broadTableEvidence,
      sheetLooksMetaOrInstruction,
      hasUrl,
    },
  };
}

function buildFallbackBlockFromMeta(sheetName, sheetInfo = {}, rows = []) {
  const meta = sheetInfo.metaData || {};
  const entries = Object.entries(meta);

  if (!rows.length) return null;

  let headerRow = Number(sheetInfo.startRow || 1);

  const headerRowCounts = new Map();

  for (const [, m] of entries) {
    const hr = Number(m?.headerRow || m?.startRow || 0);
    if (!hr) continue;
    headerRowCounts.set(hr, (headerRowCounts.get(hr) || 0) + 1);
  }

  if (headerRowCounts.size) {
    headerRow = Array.from(headerRowCounts.entries()).sort(
      (a, b) => b[1] - a[1] || a[0] - b[0],
    )[0][0];
  }

  const row = rows[headerRow - 1] || [];
  const columns = [];

  for (let i = 0; i < row.length; i += 1) {
    const rawHeader = row[i];

    if (isEmptyCell(rawHeader)) continue;

    const header = String(rawHeader).trim();

    // fallback에서는 실제 headerRow의 셀만 컬럼으로 채택
    const metaForHeader = meta[header] || {};

    columns.push({
      header,
      originalHeader: header,
      columnIndex: i + 1,
      columnLetter: metaForHeader.columnLetter || colLetter(i + 1),
    });
  }

  if (!columns.length) return null;

  const dataStartRow = Number(
    sheetInfo.dataStartRow || sheetInfo.startRow || headerRow + 1,
  );

  let dataEndRow = Number(sheetInfo.lastDataRow || rows.length);

  while (dataEndRow >= dataStartRow) {
    const r = rows[dataEndRow - 1] || [];
    const hasValue = columns.some((c) => !isEmptyCell(r[c.columnIndex - 1]));
    if (hasValue) break;
    dataEndRow -= 1;
  }

  if (dataEndRow < dataStartRow) return null;

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

function scoreQueryTable(table = {}) {
  const rowCount = Number(table.rowCount || 0);
  const colCount = Array.isArray(table.columns) ? table.columns.length : 0;

  let score = 0;

  if (rowCount > 0) score += 30;
  if (rowCount >= 5) score += 15;
  if (colCount >= 2) score += 20;
  if (colCount >= 4) score += 10;

  const typedCols = (table.columns || []).filter((c) =>
    ["number", "date", "category"].includes(String(c.type || "").toLowerCase()),
  ).length;

  if (typedCols > 0) score += 15;
  if (table.isFallback) score -= 5;

  return Math.max(0, Math.min(100, score));
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
      const fallbackBlock = buildFallbackBlockFromMeta(
        sheetName,
        sheetInfo,
        rows,
      );
      if (fallbackBlock) blocks.push(fallbackBlock);
    }

    for (const block of blocks) {
      const rawColumns = (block.columns || []).map((c, idx) => {
        const meta = sheetInfo.metaData?.[c.header] || {};
        return {
          header: String(
            c.header || c.originalHeader || `Column${idx + 1}`,
          ).trim(),
          originalHeader: String(
            c.originalHeader || c.header || `Column${idx + 1}`,
          ).trim(),
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

      const safeRawColumns = rawColumns.filter(
        (c) => c.header && c.header !== "null" && c.header !== "undefined",
      );

      const columns = uniqueKeys(safeRawColumns);
      const data = [];
      const excludedRowSet = new Set(
        (Array.isArray(block.excludedRows) ? block.excludedRows : [])
          .map((row) => Number(row?.row || row))
          .filter((row) => Number.isFinite(row)),
      );

      for (let r = block.dataStartRow; r <= block.dataEndRow; r += 1) {
        if (excludedRowSet.has(r)) continue;
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

      const queryTable = {
        tableId: block.tableId,
        isFallback: !!block.isFallback,
        source: block.isFallback ? "fallback" : "tableBlock",
        tableName: normalizeKey(block.tableId.replace("#", "_"), "table"),
        tableTitle: block.tableTitle || "",
        sheetName,
        range: block.range,
        dataRange: block.dataRange,
        headerRow: block.headerRow,
        headerRows: Array.isArray(block.headerRows)
          ? block.headerRows
          : block.headerRow
            ? [block.headerRow]
            : [],
        hasMergedHeader: Boolean(block.hasMergedHeader),
        dataStartRow: block.dataStartRow,
        dataEndRow: block.dataEndRow,
        rowCount: data.length,
        rawDataRowCount: Math.max(
          0,
          Number(block.dataEndRow || 0) - Number(block.dataStartRow || 0) + 1,
        ),
        excludedRows: Array.isArray(block.excludedRows)
          ? block.excludedRows
          : [],
        dataQuality: block.dataQuality || null,
        tableSegmentation: block.tableSegmentation || null,
        headerQuality: block.headerQuality || null,
        blockScore: block.score ?? null,
        columns,
        rows: data,
      };

      queryTable.confidence = scoreQueryTable(queryTable);
      queryTable.tableUsage = scoreTableUsageQuality(queryTable);
      tables.push(queryTable);
    }
  }

  let bestIdx = -1;
  let bestScore = -1;

  for (let i = 0; i < tables.length; i += 1) {
    const t = tables[i];
    const eligibleBonus = t.tableUsage?.templateEligible ? 1000 : 0;
    const analysisBonus = t.tableUsage?.analysisEligible ? 500 : 0;
    const score = eligibleBonus + analysisBonus + Number(t.confidence || 0);

    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }

  tables.forEach((t, idx) => {
    t.isPrimary = idx === bestIdx;
  });

  return tables;
}

module.exports = { buildQueryTablesFromWorkbook };
