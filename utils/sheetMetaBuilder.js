const XLSX = require("xlsx");
const {
  inferClusterCandidate,
  getClusterRole,
  inferClusterType,
} = require("./clusterSchema");

function nonEmptyCount(row = []) {
  let n = 0;
  for (const cell of row) {
    if (cell != null && String(cell).trim() !== "") n++;
  }
  return n;
}

function textLikeCount(row = []) {
  let n = 0;

  for (const cell of row) {
    if (cell == null || String(cell).trim() === "") continue;

    if (isNumericLike(cell)) continue;
    if (isDateLike(cell)) continue;
    if (isBooleanLike(cell)) continue;

    n += 1;
  }

  return n;
}

function compactCellText(v) {
  return String(v ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function nonEmptyCells(row = []) {
  return (row || []).filter(
    (cell) => cell != null && String(cell).trim() !== "",
  );
}

function rowBounds(row = []) {
  const indexes = rowNonEmptyIndexes(row);
  if (!indexes.length) return { min: -1, max: -1, span: 0, fillRatio: 0 };

  const min = Math.min(...indexes);
  const max = Math.max(...indexes);
  const span = max - min + 1;
  return {
    min,
    max,
    span,
    fillRatio: span > 0 ? indexes.length / span : 0,
  };
}

function isYearLikeHeader(v) {
  const s = compactCellText(v);
  return /^(19|20)\d{2}\s*년?$/.test(s);
}

function isMonthLikeHeader(v) {
  const s = compactCellText(v);
  return /^(0?[1-9]|1[0-2])\s*월$/.test(s);
}

function isQuarterLikeHeader(v) {
  const s = compactCellText(v).toUpperCase();
  return /^(Q[1-4]|[1-4]\s*분기)$/.test(s);
}

function isYearMonthLikeHeader(v) {
  const s = compactCellText(v);
  return /^(19|20)\d{2}\s*[-./년]\s*(0?[1-9]|1[0-2])\s*월?$/.test(s);
}

function isTemporalHeaderLike(v) {
  return (
    isYearLikeHeader(v) ||
    isMonthLikeHeader(v) ||
    isQuarterLikeHeader(v) ||
    isYearMonthLikeHeader(v)
  );
}

function isFormulaOrErrorLike(v) {
  const s = compactCellText(v);
  return /^=/.test(s) || /^#(DIV\/0|N\/A|NAME|NULL|NUM|REF|VALUE)!?$/i.test(s);
}

function isLikelyHeaderToken(v) {
  const s = compactCellText(v);
  if (!s) return false;
  if (isFormulaOrErrorLike(s)) return false;
  if (isTemporalHeaderLike(s)) return true;
  if (isBooleanLike(s)) return false;
  if (isDateLike(s) || isTimeLike(s)) return false;
  if (isNumericLike(s)) return false;

  // 도메인 키워드를 박지 않고, "짧고 구조적인 라벨" 자체를 헤더 토큰으로 본다.
  // 예: 부서, 직급, 이름, amount, category, A/B/C 같은 코드형 라벨
  return s.length <= 40;
}

function looksLikeMetaOrProseRow(row = []) {
  const cells = nonEmptyCells(row).map(compactCellText);
  if (!cells.length) return false;

  const joined = cells.join(" ");
  const longTextCount = cells.filter((c) => c.length >= 25).length;
  const keyValueLikeCount = cells.filter((c) => /[:：]/.test(c)).length;
  const sentenceMarkCount = cells.filter((c) =>
    /[.!?。]|입니다|합니다|대한|관련/.test(c),
  ).length;
  const noteLike = /^\s*(※|\*|주\)|note\b|remark\b)/i.test(joined);

  return (
    noteLike ||
    (cells.length <= 2 && longTextCount >= 1) ||
    (cells.length <= 3 && keyValueLikeCount >= 1) ||
    (cells.length <= 3 && sentenceMarkCount >= 1)
  );
}

function normalizeNoiseText(value = "") {
  return compactCellText(value)
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "")
    .trim();
}

function isBlankRow(row = []) {
  return nonEmptyCount(row) === 0;
}

function rowSliceByColumnRange(row = [], minCol = 0, maxCol = 0) {
  const out = [];
  for (let c = minCol; c <= maxCol; c += 1) out.push(row?.[c]);
  return out;
}

function nonEmptyCountInRange(row = [], minCol = 0, maxCol = 0) {
  return nonEmptyCount(rowSliceByColumnRange(row, minCol, maxCol));
}

function numericDateCount(row = []) {
  return (row || []).filter(
    (cell) =>
      cell != null &&
      String(cell).trim() !== "" &&
      (isNumericLike(cell) || isDateLike(cell) || isTimeLike(cell)),
  ).length;
}

function looksLikeNoteOrCommentRow(row = [], options = {}) {
  const cells = nonEmptyCells(row).map(compactCellText);
  if (!cells.length) return false;

  const joined = cells.join(" ");
  const normalized = normalizeNoiseText(joined);
  const minCol = Number.isInteger(options.minCol) ? options.minCol : 0;
  const maxCol = Number.isInteger(options.maxCol)
    ? options.maxCol
    : Math.max(row.length - 1, minCol);
  const span = Math.max(1, maxCol - minCol + 1);
  const inRangeCount = nonEmptyCountInRange(row, minCol, maxCol);
  const sparseInTable = inRangeCount <= Math.max(1, Math.floor(span * 0.25));

  const markerLike =
    /^\s*(※|\*|주\)|주\s*[:：.]|note\b|remark\b|source\b|출처\s*[:：]|단위\s*[:：])/i.test(
      joined,
    );
  const proseLike = looksLikeMetaOrProseRow(row);
  const mostlyLongText = cells.length <= 2 && cells.some((c) => c.length >= 18);
  const hasOnlyText = numericDateCount(row) === 0;

  return Boolean(
    markerLike ||
    (sparseInTable && hasOnlyText && mostlyLongText) ||
    (sparseInTable &&
      proseLike &&
      !/^합계|^총계|^소계|^total|^subtotal/i.test(normalized)),
  );
}

const SUMMARY_ROW_DETECTOR_VERSION = "summary_row_detector_v2";

function normalizeSummaryToken(value = "") {
  return normalizeNoiseText(value).replace(/grandtotal/g, "grandtotal");
}

function isStrongSummaryToken(value = "") {
  const normalized = normalizeSummaryToken(value);
  return /^(합계|총계|소계|누계|계|total|subtotal|grandtotal)$/.test(
    normalized,
  );
}

function isWeakSummaryToken(value = "") {
  const normalized = normalizeSummaryToken(value);
  return /^(전체|all|overall)$/.test(normalized);
}

function getLeadingNonNumericCells(slice = [], limit = 3) {
  const out = [];

  for (let idx = 0; idx < slice.length && out.length < limit; idx += 1) {
    const value = slice[idx];
    if (value == null || String(value).trim() === "") continue;
    if (isNumericLike(value) || isDateLike(value) || isTimeLike(value)) break;
    out.push({ value: compactCellText(value), index: idx });
  }

  return out;
}

function looksLikeTotalOrSubtotalRow(row = [], options = {}) {
  const cells = nonEmptyCells(row).map(compactCellText);
  if (!cells.length) return false;

  const first = normalizeNoiseText(cells[0]);
  const joined = normalizeNoiseText(cells.join(" "));
  const minCol = Number.isInteger(options.minCol) ? options.minCol : 0;
  const maxCol = Number.isInteger(options.maxCol)
    ? options.maxCol
    : Math.max(row.length - 1, minCol);
  const slice = rowSliceByColumnRange(row, minCol, maxCol);
  const nonEmpty = nonEmptyCount(slice);
  const numericLike = slice.filter(
    (cell) => cell != null && String(cell).trim() !== "" && isNumericLike(cell),
  ).length;
  const numericDateLike = numericDateCount(slice);
  const leadingContextCells = getLeadingNonNumericCells(slice, 3);
  const leadingValues = leadingContextCells.map((cell) => cell.value);
  const strongSummaryInLeadingContext =
    leadingValues.some(isStrongSummaryToken);
  const weakSummaryInLeadingContext = leadingValues.some(isWeakSummaryToken);
  const hasContextBeforeMetric = leadingContextCells.length > 0;

  const startsWithSummaryToken =
    /^(합계|총계|소계|누계|계|total|subtotal|grandtotal)$/.test(first) ||
    /^(합계|총계|소계|누계|total|subtotal|grandtotal)/.test(joined);

  const dimensionSubtotalPattern = Boolean(
    strongSummaryInLeadingContext &&
    hasContextBeforeMetric &&
    (numericLike > 0 || numericDateLike > leadingContextCells.length),
  );

  const weakTotalPattern = Boolean(
    weakSummaryInLeadingContext &&
    leadingValues.some(isStrongSummaryToken) &&
    (numericLike > 0 || numericDateLike > leadingContextCells.length),
  );

  return Boolean(
    (startsWithSummaryToken && (numericLike > 0 || nonEmpty <= 2)) ||
    dimensionSubtotalPattern ||
    weakTotalPattern,
  );
}

function normalizeHeaderCompare(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "")
    .trim();
}

function looksLikeRepeatedHeaderRow(row = [], headerRow = [], headerCols = []) {
  if (!Array.isArray(headerCols) || !headerCols.length) return false;

  let comparable = 0;
  let matched = 0;

  for (const col of headerCols) {
    const left = normalizeHeaderCompare(row?.[col]);
    const right = normalizeHeaderCompare(headerRow?.[col]);
    if (!left || !right) continue;
    comparable += 1;

    // 반복 헤더는 같은 위치의 헤더 토큰이 다시 등장하는 경우만 본다.
    // 기존의 includes 비교는 숫자 데이터(예: 1, 2)가 2021/2022 헤더에 포함된다고
    // 오인해 실제 데이터 행을 REPEATED_HEADER_ROW로 제외하는 문제가 있었다.
    if (left === right) matched += 1;
  }

  if (comparable < Math.min(2, headerCols.length)) return false;
  return matched / comparable >= 0.7;
}

function classifyTableRow(row = [], context = {}) {
  const headerRow = context.headerRow || [];
  const headerCols = context.headerCols || [];
  const minCol = Number.isInteger(context.minCol) ? context.minCol : 0;
  const maxCol = Number.isInteger(context.maxCol)
    ? context.maxCol
    : Math.max(row.length - 1, minCol);
  const span = Math.max(1, maxCol - minCol + 1);
  const nonEmpty = nonEmptyCountInRange(row, minCol, maxCol);
  const density = nonEmpty / span;

  if (isBlankRow(row)) {
    return {
      kind: "blank",
      include: false,
      stopHint: true,
      reason: "BLANK_ROW",
    };
  }

  if (looksLikeRepeatedHeaderRow(row, headerRow, headerCols)) {
    return {
      kind: "repeatedHeader",
      include: false,
      stopHint: false,
      reason: "REPEATED_HEADER_ROW",
    };
  }

  if (looksLikeNoteOrCommentRow(row, { minCol, maxCol })) {
    return {
      kind: "note",
      include: false,
      stopHint: true,
      reason: "NOTE_OR_COMMENT_ROW",
    };
  }

  if (looksLikeTotalOrSubtotalRow(row, { minCol, maxCol })) {
    return {
      kind: "total",
      include: false,
      stopHint: true,
      reason: "TOTAL_OR_SUBTOTAL_ROW",
    };
  }

  if (looksLikeSectionTitleRow(row, { minCol, maxCol })) {
    return {
      kind: "sectionTitle",
      include: false,
      stopHint: true,
      reason: "SECTION_TITLE_ROW",
    };
  }

  if (nonEmpty === 0) {
    return {
      kind: "blank",
      include: false,
      stopHint: true,
      reason: "BLANK_ROW",
    };
  }

  if (density < 0.2 && span >= 5 && numericDateCount(row) === 0) {
    return {
      kind: "sparseText",
      include: false,
      stopHint: true,
      reason: "SPARSE_TEXT_ROW",
    };
  }

  return { kind: "data", include: true, stopHint: false, reason: "DATA_ROW" };
}

function hasFutureDataRow(json = [], start = 0, end = 0, context = {}) {
  for (let r = start; r < Math.min(end, json.length); r += 1) {
    const row = json[r] || [];
    const classified = classifyTableRow(row, context);
    if (classified.include) return true;

    // 다음 표나 긴 설명 영역으로 보이면 더 멀리 보지 않는다.
    if (classified.kind === "note" || classified.kind === "total") {
      const after = json[r + 1] || [];
      if (!nonEmptyCount(after)) return false;
    }
  }

  return false;
}

function summarizeExcludedRows(excludedRows = []) {
  const counts = {};
  for (const row of excludedRows || []) {
    const reason = row.reason || "EXCLUDED";
    counts[reason] = (counts[reason] || 0) + 1;
  }
  return counts;
}

const TABLE_SEGMENTATION_VERSION = "table_segmentation_v1";
const TABLE_SEGMENTATION_OPTIONS = Object.freeze({
  maxLookaheadRows: 5,
  maxBoundaryLookbackRows: 5,
  maxTitleLookbackRows: 4,
  embeddedHeaderMinScore: 32,
  embeddedHeaderMinDataEvidence: 0.5,
});

function sortUniqueRowIndexes(values = []) {
  return [
    ...new Set(values.map(Number).filter((n) => Number.isInteger(n) && n >= 0)),
  ].sort((a, b) => a - b);
}

function compactRowPreview(row = [], max = 80) {
  const s = nonEmptyCells(row).map(compactCellText).join(" | ");
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

function firstNonEmptyIndex(row = []) {
  for (let i = 0; i < row.length; i += 1) {
    if (isNonEmptyCell(row[i])) return i;
  }
  return -1;
}

function lastNonEmptyIndex(row = []) {
  for (let i = row.length - 1; i >= 0; i -= 1) {
    if (isNonEmptyCell(row[i])) return i;
  }
  return -1;
}

function rowSpanFillRatio(row = []) {
  const first = firstNonEmptyIndex(row);
  const last = lastNonEmptyIndex(row);
  if (first < 0 || last < first) return 0;
  const span = last - first + 1;
  return span > 0 ? nonEmptyCount(row) / span : 0;
}

function looksLikeSectionTitleRow(row = [], options = {}) {
  const cells = nonEmptyCells(row).map(compactCellText);
  if (!cells.length) return false;
  if (looksLikeNoteOrCommentRow(row, options)) return false;
  if (looksLikeTotalOrSubtotalRow(row, options)) return false;
  if (structurallyHeaderishRow(row)) return false;
  if (numericDateCount(row) > 0) return false;

  const first = firstNonEmptyIndex(row);
  const last = lastNonEmptyIndex(row);
  const span = first >= 0 && last >= first ? last - first + 1 : 0;
  const fillRatio = rowSpanFillRatio(row);
  const joined = cells.join(" ");

  return Boolean(
    cells.length <= 2 &&
    joined.length <= 120 &&
    (fillRatio <= 0.5 || span <= 3 || looksLikeMetaOrProseRow(row)),
  );
}

function hasSegmentationBoundaryBefore(json = [], rowIndex = 0, options = {}) {
  const lookback = Number(
    options.maxBoundaryLookbackRows ||
      TABLE_SEGMENTATION_OPTIONS.maxBoundaryLookbackRows,
  );
  const start = Math.max(0, rowIndex - lookback);

  for (let r = rowIndex - 1; r >= start; r -= 1) {
    const row = json[r] || [];
    if (isBlankRow(row)) return true;
    if (looksLikeSectionTitleRow(row, options)) return true;
    if (looksLikeNoteOrCommentRow(row, options)) return true;
    if (looksLikeTotalOrSubtotalRow(row, options)) return true;

    // 값이 많은 데이터 행이 바로 이어져 있으면 같은 표 내부일 가능성이 높다.
    if (nonEmptyCount(row) >= 2 && numericDateCount(row) > 0) return false;
  }

  return rowIndex === 0;
}

function isSegmentHeaderCandidate(json = [], rowIndex = 0, options = {}) {
  const row = json[rowIndex] || [];
  if (!nonEmptyCount(row)) return null;
  if (nonEmptyCount(row) < 2) return null;
  if (looksLikeMetaOrProseRow(row)) return null;
  if (looksLikeNoteOrCommentRow(row, options)) return null;
  if (looksLikeTotalOrSubtotalRow(row, options)) return null;

  const analysis = analyzeHeaderRowCandidate(row, {
    rowIndex,
    prevRow: rowIndex > 0 ? json[rowIndex - 1] || [] : null,
    nextRows: json.slice(rowIndex + 1, Math.min(json.length, rowIndex + 6)),
  });

  const enoughScore =
    Number(analysis.score || 0) >=
    Number(
      options.embeddedHeaderMinScore ||
        TABLE_SEGMENTATION_OPTIONS.embeddedHeaderMinScore,
    );
  const enoughDataEvidence =
    Number(analysis.nextDataEvidence || 0) >=
    Number(
      options.embeddedHeaderMinDataEvidence ||
        TABLE_SEGMENTATION_OPTIONS.embeddedHeaderMinDataEvidence,
    );
  const enoughHeaderShape =
    Number(analysis.headerLikeRatio || 0) >= 0.5 ||
    Number(analysis.temporalCount || 0) >= 2 ||
    structurallyHeaderishRow(row);

  if (!enoughScore || !enoughDataEvidence || !enoughHeaderShape) return null;

  return analysis;
}

function looksLikeHeaderFalsePositiveDataRow(row = [], analysis = {}) {
  const numericDateRatio = Number(analysis.numericDateRatio || 0);
  const temporalCount = Number(analysis.temporalCount || 0);

  return Boolean(
    isLikelyDataRow(row) ||
    (numericDateRatio > 0.2 && temporalCount < 2 && numericDateCount(row) > 0),
  );
}

function looksLikeRepeatedSegmentHeader(
  json = [],
  rowIndex = 0,
  previousHeaderIndex = null,
) {
  if (!Number.isInteger(previousHeaderIndex) || previousHeaderIndex < 0) {
    return false;
  }

  const candidateRow = json[rowIndex] || [];
  const previousRow = json[previousHeaderIndex] || [];
  const previousCols = rowNonEmptyIndexes(previousRow);

  if (!previousCols.length) return false;

  return looksLikeRepeatedHeaderRow(candidateRow, previousRow, previousCols);
}

function shouldKeepSegmentHeader(
  json = [],
  rowIndex = 0,
  kept = [],
  options = {},
) {
  const analysis = isSegmentHeaderCandidate(json, rowIndex, options);
  if (!analysis)
    return { keep: false, analysis: null, reason: "NOT_HEADER_CANDIDATE" };

  const hasBoundary = hasSegmentationBoundaryBefore(json, rowIndex, options);
  const previousKept = kept.length ? kept[kept.length - 1] : null;
  const dataLike = looksLikeHeaderFalsePositiveDataRow(
    json[rowIndex] || [],
    analysis,
  );
  const repeatedHeaderInsideOpenTable =
    previousKept != null &&
    !hasBoundary &&
    looksLikeRepeatedSegmentHeader(json, rowIndex, previousKept);

  if (repeatedHeaderInsideOpenTable) {
    return {
      keep: false,
      analysis,
      reason: "REPEATED_HEADER_INSIDE_TABLE",
    };
  }

  if (dataLike && !hasBoundary && previousKept != null) {
    return { keep: false, analysis, reason: "DATA_ROW_FALSE_POSITIVE" };
  }

  return { keep: true, analysis, reason: "SEGMENT_HEADER" };
}

function buildSegmentationHeaderIndexes(
  json = [],
  headerRowIndexes = [],
  options = {},
) {
  const original = sortUniqueRowIndexes(headerRowIndexes);
  const first = Number.isInteger(options.firstNonEmpty)
    ? options.firstNonEmpty
    : 0;
  const last = Number.isInteger(options.lastNonEmpty)
    ? options.lastNonEmpty
    : Math.max(0, json.length - 1);

  const candidates = new Map();
  for (const rowIndex of original) {
    candidates.set(rowIndex, { source: "detected" });
  }

  for (let r = first; r <= last; r += 1) {
    if (candidates.has(r)) continue;

    const analysis = isSegmentHeaderCandidate(json, r, options);
    if (!analysis) continue;
    if (!hasSegmentationBoundaryBefore(json, r, options)) continue;

    candidates.set(r, { source: "segmentation", analysis });
  }

  const kept = [];
  const added = [];
  const removed = [];

  for (const rowIndex of sortUniqueRowIndexes([...candidates.keys()])) {
    const item = candidates.get(rowIndex) || {};
    const decision = shouldKeepSegmentHeader(json, rowIndex, kept, options);

    if (!decision.keep) {
      removed.push({
        rowIndex,
        source: item.source || "detected",
        reason: decision.reason,
        preview: compactRowPreview(json[rowIndex] || []),
      });
      continue;
    }

    kept.push(rowIndex);

    if (item.source === "segmentation") {
      const analysis = decision.analysis || item.analysis || {};
      added.push({
        rowIndex,
        score: analysis.score,
        nextDataEvidence: analysis.nextDataEvidence,
        reasons: analysis.reasons || [],
        preview: compactRowPreview(json[rowIndex] || []),
      });
    }
  }

  return {
    version: TABLE_SEGMENTATION_VERSION,
    originalRows: original,
    addedRows: added,
    removedRows: removed,
    headerRowIndexes: kept,
  };
}

function findPrecedingSectionTitleRows(
  json = [],
  headerStartIndex = 0,
  options = {},
) {
  const maxLookback = Number(
    options.maxTitleLookbackRows ||
      TABLE_SEGMENTATION_OPTIONS.maxTitleLookbackRows,
  );
  const start = Math.max(0, headerStartIndex - maxLookback);
  const rows = [];

  for (let r = headerStartIndex - 1; r >= start; r -= 1) {
    const row = json[r] || [];
    if (isBlankRow(row)) break;
    if (looksLikeSectionTitleRow(row, options)) {
      rows.unshift({ row: r + 1, text: compactRowPreview(row, 120) });
      continue;
    }
    if (looksLikeNoteOrCommentRow(row, options)) continue;
    break;
  }

  return rows;
}

function nearestFutureHeaderIndex(
  selectedHeaders = [],
  headerIndex = 0,
  fallback = 0,
) {
  return selectedHeaders.find((idx) => idx > headerIndex) ?? fallback;
}

function shouldStopAtExcludedRow({ classified = {}, hasFuture = false } = {}) {
  if (classified.kind === "blank") return !hasFuture;
  if (classified.kind === "total") return !hasFuture;
  if (classified.kind === "note") return !hasFuture;
  if (classified.kind === "sectionTitle") return !hasFuture;
  if (classified.kind === "sparseText") return !hasFuture;
  return false;
}

function structurallyHeaderishRow(row = []) {
  const nonEmpty = nonEmptyCount(row);
  if (nonEmpty < 2) return false;

  const temporalCount = nonEmptyCells(row).filter(isTemporalHeaderLike).length;
  const headerLikeCount = nonEmptyCells(row).filter(isLikelyHeaderToken).length;
  const numericLike = row.filter(
    (cell) => cell != null && String(cell).trim() !== "" && isNumericLike(cell),
  ).length;
  const dateLike = row.filter(
    (cell) => cell != null && String(cell).trim() !== "" && isDateLike(cell),
  ).length;
  const numericDateRatio = (numericLike + dateLike) / nonEmpty;

  if (temporalCount >= 2) return true;
  return headerLikeCount / nonEmpty >= 0.6 && numericDateRatio <= 0.15;
}

function scoreFollowingDataEvidence(row = [], nextRows = []) {
  const nonEmpty = nonEmptyCount(row);
  if (!nonEmpty || !Array.isArray(nextRows) || !nextRows.length) return 0;

  let evidence = 0;
  const maxLookahead = Math.min(nextRows.length, 4);

  for (let i = 0; i < maxLookahead; i += 1) {
    const next = nextRows[i] || [];
    const n = nonEmptyCount(next);
    if (!n) continue;

    const nextText = textLikeCount(next);
    const nextNumeric = next.filter(
      (cell) =>
        cell != null && String(cell).trim() !== "" && isNumericLike(cell),
    ).length;
    const nextDate = next.filter(
      (cell) => cell != null && String(cell).trim() !== "" && isDateLike(cell),
    ).length;

    const widthCompatible = n >= Math.max(2, Math.ceil(nonEmpty * 0.5));
    const hasDataTypeSignal = nextNumeric + nextDate > 0 || nextText > 0;

    if (widthCompatible && hasDataTypeSignal) {
      evidence += i === 0 ? 1 : 0.65;
    }
  }

  return Math.min(1, evidence);
}

function analyzeHeaderRowCandidate(row = [], options = {}) {
  const nonEmpty = nonEmptyCount(row);
  const cells = nonEmptyCells(row);
  if (!nonEmpty) {
    return {
      score: 0,
      isCandidate: false,
      nonEmpty: 0,
      textLike: 0,
      textRatio: 0,
      headerLikeRatio: 0,
      numericDateRatio: 0,
      reasons: ["EMPTY_ROW"],
    };
  }

  const textLike = textLikeCount(row);
  const textRatio = textLike / nonEmpty;
  const headerLike = cells.filter(isLikelyHeaderToken).length;
  const headerLikeRatio = headerLike / nonEmpty;
  const temporalCount = cells.filter(isTemporalHeaderLike).length;
  const temporalRatio = temporalCount / nonEmpty;
  const numericLike = row.filter(
    (cell) => cell != null && String(cell).trim() !== "" && isNumericLike(cell),
  ).length;
  const dateLike = row.filter(
    (cell) => cell != null && String(cell).trim() !== "" && isDateLike(cell),
  ).length;
  const numericDateRatio = (numericLike + dateLike) / nonEmpty;
  const bounds = rowBounds(row);
  const nextDataEvidence = scoreFollowingDataEvidence(
    row,
    options.nextRows || [],
  );
  const prevRow = Array.isArray(options.prevRow) ? options.prevRow : null;
  const prevLooksHeader = prevRow ? structurallyHeaderishRow(prevRow) : false;
  const metaOrProse = looksLikeMetaOrProseRow(row);

  let score = 0;
  const reasons = [];

  if (nonEmpty >= 2) {
    score += Math.min(nonEmpty, 12) * 2;
    reasons.push("ENOUGH_CELLS");
  } else {
    score -= 18;
    reasons.push("TOO_FEW_CELLS");
  }

  score += textRatio * 12;
  score += headerLikeRatio * 16;

  if (temporalCount >= 2) {
    score += 14;
    reasons.push("TEMPORAL_HEADER_TOKENS");
  } else if (temporalCount === 1 && nonEmpty >= 2) {
    score += 4;
    reasons.push("TEMPORAL_HEADER_TOKEN");
  }

  if (bounds.fillRatio >= 0.7) {
    score += 6;
    reasons.push("DENSE_HEADER_SPAN");
  } else if (bounds.fillRatio < 0.35) {
    score -= 5;
    reasons.push("SPARSE_ROW");
  }

  if (nextDataEvidence > 0) {
    score += nextDataEvidence * 18;
    reasons.push("FOLLOWING_DATA_EVIDENCE");
  }

  // 일반 숫자/날짜값이 섞인 행은 데이터 행일 가능성이 높다.
  // 단, 연도/월/분기처럼 시간 축 헤더로 보이는 숫자는 감점하지 않는다.
  if (numericDateRatio > 0.25 && temporalRatio < 0.5) {
    score -= 14;
    reasons.push("DATA_VALUE_RATIO_PENALTY");
  }

  if (prevLooksHeader && numericDateRatio > 0.15 && temporalRatio < 0.5) {
    score -= 22;
    reasons.push("PREVIOUS_HEADER_DATA_ROW_PENALTY");
  }

  if (metaOrProse) {
    score -= 22;
    reasons.push("META_OR_PROSE_ROW_PENALTY");
  }

  if (isLikelyDataRow(row) && headerLikeRatio < 0.75 && temporalCount < 2) {
    score -= 16;
    reasons.push("LIKELY_DATA_ROW_PENALTY");
  }

  const isCandidate =
    nonEmpty >= 2 &&
    score >= 30 &&
    (headerLikeRatio >= 0.5 || temporalCount >= 2 || textRatio >= 0.6);

  return {
    score: Math.round(score * 100) / 100,
    isCandidate,
    nonEmpty,
    textLike,
    textRatio: Math.round(textRatio * 1000) / 1000,
    headerLike,
    headerLikeRatio: Math.round(headerLikeRatio * 1000) / 1000,
    temporalCount,
    numericDateRatio: Math.round(numericDateRatio * 1000) / 1000,
    nextDataEvidence: Math.round(nextDataEvidence * 1000) / 1000,
    fillRatio: Math.round(bounds.fillRatio * 1000) / 1000,
    reasons,
  };
}

function scoreHeaderRowCandidate(row = [], options = {}) {
  return analyzeHeaderRowCandidate(row, options).score;
}

function indexToColumnLetter(idx) {
  let n = idx + 1,
    s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// 숫자 비슷한 값인지 (기존 isNumericLike 로직과 동일하게 유지)
function isNumericLike(v) {
  if (v === null || v === undefined) return false;
  if (v instanceof Date) return false;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "" || /[^\d.+\-eE]/.test(s)) return false;
  const n = Number(s);
  return Number.isFinite(n);
}

function isBooleanLike(v) {
  if (typeof v === "boolean") return true;
  const s = String(v).trim().toLowerCase();
  return ["true", "false", "yes", "no", "y", "n", "예", "아니오"].includes(s);
}

// 단순 날짜 패턴 (JS Date 객체이거나, 문자열 패턴 기반)
function isDateLike(v) {
  if (v instanceof Date) return true;
  if (v == null) return false;
  const s = String(v).trim();
  if (!s) return false;

  // YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD
  if (/^\d{4}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])/.test(s)) {
    return true;
  }

  // "2025-01-01 12:34" 같이 날짜+시간
  if (
    /^\d{4}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])\s+\d{1,2}:\d{2}/.test(
      s,
    )
  ) {
    return true;
  }

  return false;
}

// 시간만 있는 패턴 (HH:MM 또는 HH:MM:SS)
function isTimeLike(v) {
  if (v == null) return false;
  const s = String(v).trim();
  if (!s) return false;

  // 0:00 ~ 23:59(:59)
  return /^([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/.test(s);
}

function isTimeLikeString(s) {
  // 09:30, 9:30, 09:30:15
  return /^([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/.test(s);
}

function isDateLikeString(s) {
  // 2024-01-01 / 2024.01.01 / 2024/01/01
  if (/^(19|20)\d{2}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])$/.test(s))
    return true;
  // 20240101 같은 8자리 숫자도 간단히 케이스 추가 가능
  if (/^(19|20)\d{6}$/.test(s)) return true;
  return false;
}

function isDateTimeLikeString(s) {
  // 2024-01-01 09:30 / 2024-01-01T09:30:00
  return /^(19|20)\d{2}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])[ T]([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/.test(
    s,
  );
}

// 셀 하나의 타입 분류 - number / date / time / text / empty
function detectCellType(v) {
  if (v === null || v === undefined) return "empty";
  if (v instanceof Date) return "date";

  // 숫자 타입은 그냥 숫자로 본다 (엑셀에서 날짜를 Date로 주는 경우가 많음)
  if (typeof v === "number") {
    return "number";
  }

  const s = String(v).trim();
  if (!s) return "empty";

  if (isDateLike(s)) return "date";
  if (isTimeLike(s)) return "time";
  if (isNumericLike(s)) return "number";

  return "text";
}

function analyzeSamples(values) {
  let numeric = 0;
  let date = 0;
  let datetime = 0;
  let time = 0;
  let bool = 0;
  let text = 0;

  for (const v of values) {
    if (v == null) continue;

    if (v instanceof Date) {
      // 엑셀 날짜가 Date로 들어오는 경우
      datetime++;
      continue;
    }

    const s = String(v).trim();
    if (!s) continue;

    if (isBooleanLike(v)) {
      bool++;
    } else if (isDateTimeLikeString(s)) {
      datetime++;
    } else if (isDateLikeString(s)) {
      date++;
    } else if (isTimeLikeString(s)) {
      time++;
    } else if (isNumericLike(v)) {
      numeric++;
    } else {
      text++;
    }
  }

  const total = numeric + date + datetime + time + bool + text || 1;

  const ratios = {
    numericRatio: numeric / total,
    dateRatio: date / total,
    datetimeRatio: datetime / total,
    timeRatio: time / total,
    booleanRatio: bool / total,
    textRatio: text / total,
  };

  // 대표 타입(dominantType) 추론
  const entries = [
    ["number", ratios.numericRatio],
    ["date", ratios.dateRatio + ratios.datetimeRatio],
    ["time", ratios.timeRatio],
    ["boolean", ratios.booleanRatio],
    ["text", ratios.textRatio],
  ].sort((a, b) => b[1] - a[1]);

  const [topType, topRatio] = entries[0];
  const dominantType = topRatio >= 0.5 ? topType : "mixed";

  return {
    ...ratios,
    dominantType,
    sampleCount: total,
  };
}

function normalizeSampleValue(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v ?? "").trim();
}

function buildColumnProfile(header, values, stats) {
  const nonEmptyValues = values
    .map(normalizeSampleValue)
    .filter((v) => v !== "");

  const uniqueValues = [...new Set(nonEmptyValues)];
  const sampleCount = Number(stats?.sampleCount || nonEmptyValues.length || 0);
  const uniqueCount = uniqueValues.length;
  const uniqueRatio = sampleCount > 0 ? uniqueCount / sampleCount : 0;

  const dominantType = String(stats?.dominantType || "mixed").toLowerCase();

  let profileType = dominantType;
  if (Number(stats?.numericRatio || 0) >= 0.8) profileType = "number";
  else if (
    Number(stats?.dateRatio || 0) + Number(stats?.datetimeRatio || 0) >=
    0.6
  ) {
    profileType = "date";
  } else if (
    Number(stats?.textRatio || 0) >= 0.6 &&
    uniqueRatio <= 0.4 &&
    uniqueCount <= 30
  ) {
    profileType = "category";
  } else if (Number(stats?.textRatio || 0) >= 0.6) {
    profileType = "text";
  }

  return {
    profileType,
    uniqueCount,
    uniqueRatio,
    uniqueValues: uniqueValues.slice(0, 20),
  };
}

function isNonEmptyCell(v) {
  return v != null && String(v).trim() !== "";
}

function rowNonEmptyIndexes(row = []) {
  const out = [];
  for (let i = 0; i < row.length; i++) {
    if (isNonEmptyCell(row[i])) out.push(i);
  }
  return out;
}

function firstNonEmptyCell(row = []) {
  const idx = row.findIndex((cell) => isNonEmptyCell(cell));
  return idx >= 0 ? row[idx] : null;
}

function isLikelyDataRow(row = []) {
  const nonEmpty = nonEmptyCount(row);
  if (!nonEmpty) return false;

  const first = firstNonEmptyCell(row);
  const firstLooksData = isNumericLike(first) || isDateLike(first);

  const numericLike = row.filter(
    (cell) => cell != null && String(cell).trim() !== "" && isNumericLike(cell),
  ).length;

  const dateLike = row.filter(
    (cell) => cell != null && String(cell).trim() !== "" && isDateLike(cell),
  ).length;

  const numericDateRatio = (numericLike + dateLike) / nonEmpty;

  return firstLooksData && numericDateRatio > 0;
}

function normalizeHeaderText(v) {
  return String(v ?? "").trim();
}

function normalizeHeaderPart(v) {
  return String(v ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function isSameHeaderPart(a = "", b = "") {
  const x = normalizeHeaderPart(a).toLowerCase();
  const y = normalizeHeaderPart(b).toLowerCase();
  return Boolean(x && y && x === y);
}

function uniqueHeaderParts(parts = []) {
  const out = [];

  for (const part of parts.map(normalizeHeaderPart).filter(Boolean)) {
    if (out.some((existing) => isSameHeaderPart(existing, part))) continue;
    out.push(part);
  }

  return out;
}

function joinHeaderParts(parts = []) {
  const safeParts = uniqueHeaderParts(parts);
  if (!safeParts.length) return "";
  return safeParts.join("_");
}

function countTemporalHeaderTokens(row = []) {
  return nonEmptyCells(row).filter(isTemporalHeaderLike).length;
}

function countHeaderLikeTokens(row = []) {
  return nonEmptyCells(row).filter(isLikelyHeaderToken).length;
}

function rowLooksLikeSparseHeaderLayer(row = [], lowerRow = []) {
  const cells = nonEmptyCells(row);
  if (!cells.length) return false;
  if (looksLikeMetaOrProseRow(row)) return false;

  const nonEmpty = cells.length;
  const temporalCount = cells.filter(isTemporalHeaderLike).length;
  const lowerTemporalCount = countTemporalHeaderTokens(lowerRow);

  // 연도/분기/월처럼 시간축 상위 라벨은 셀이 하나만 있어도 헤더 레이어로 본다.
  if (temporalCount >= 1) return true;

  // 바로 아래 행이 월/분기형 헤더라면 짧은 단일 상위 라벨도 허용한다.
  // 예: ["2024"] / ["1월", "2월"] 또는 ["상반기"] / ["1월", "2월"]
  if (nonEmpty === 1 && lowerTemporalCount >= 2) {
    const text = normalizeHeaderPart(cells[0]);
    return text.length > 0 && text.length <= 20 && !isLikelyDataRow(row);
  }

  return false;
}

function rowLooksLikeHeaderLayer(row = [], lowerRow = []) {
  const nonEmpty = nonEmptyCount(row);
  if (!nonEmpty) return false;
  if (looksLikeMetaOrProseRow(row)) return false;

  if (rowLooksLikeSparseHeaderLayer(row, lowerRow)) return true;

  const headerLikeCount = countHeaderLikeTokens(row);
  const headerLikeRatio = headerLikeCount / nonEmpty;
  const temporalCount = countTemporalHeaderTokens(row);

  if (nonEmpty >= 2 && headerLikeRatio >= 0.5) return true;
  if (temporalCount >= 2) return true;

  return false;
}

function rowLooksLikeDataLayerAboveHeader(row = [], lowerRow = []) {
  const nonEmpty = nonEmptyCount(row);
  const lowerNonEmpty = nonEmptyCount(lowerRow);
  if (!nonEmpty || !lowerNonEmpty) return false;

  const temporalCount = countTemporalHeaderTokens(row);
  const numericDate = numericDateCount(row);
  const sameWidthAsLower =
    nonEmpty >= Math.max(2, Math.ceil(lowerNonEmpty * 0.6));
  const lowerHeaderLikeRatio = countHeaderLikeTokens(lowerRow) / lowerNonEmpty;
  const lowerTemporalCount = countTemporalHeaderTokens(lowerRow);
  const lowerHeaderKeywordCount = nonEmptyCells(lowerRow).filter((cell) => {
    const token = normalizeHeaderCompare(cell);
    return (
      /^(19|20)\d{2}$/.test(token) ||
      /^(계|합계|총계|소계|남자|여자|구분|분류|항목|전체)$/.test(token)
    );
  }).length;
  const lowerLooksHeaderish =
    lowerHeaderLikeRatio >= 0.35 ||
    lowerTemporalCount >= 1 ||
    lowerHeaderKeywordCount >= Math.min(2, lowerNonEmpty) ||
    structurallyHeaderishRow(lowerRow);

  // 반복 헤더 바로 위의 실제 데이터행(예: 개발/대리/5000, 배임/2042)이
  // 상위 병합 헤더 레이어로 붙는 것을 막는다.
  if (sameWidthAsLower && lowerLooksHeaderish && numericDate > 0) {
    return temporalCount === 0 || isLikelyDataRow(row);
  }

  return false;
}

function getFlattenHeaderBandRows(
  json = [],
  headerIndex,
  maxDepth = 4,
  options = {},
) {
  const rows = [headerIndex];
  const selectedHeaderSet = new Set(
    (options.headerRowIndexes || options.selectedHeaderIndexes || [])
      .map(Number)
      .filter((n) => Number.isInteger(n) && n >= 0),
  );

  for (let r = headerIndex - 1; r >= 0 && rows.length < maxDepth; r -= 1) {
    const row = json[r] || [];
    const lowerRow = json[r + 1] || [];

    if (!nonEmptyCount(row)) break;
    if (!rowLooksLikeHeaderLayer(row, lowerRow)) break;
    if (rowLooksLikeDataLayerAboveHeader(row, lowerRow)) break;

    const selectedAsHeader = selectedHeaderSet.has(r);
    const sparseHeaderLayer = rowLooksLikeSparseHeaderLayer(row, lowerRow);
    const temporalLayer = countTemporalHeaderTokens(row) >= 1;

    // 폭이 넓은 일반 텍스트 행은 실제 데이터행일 가능성이 높다.
    // 명시 헤더 후보로 선택된 행, 희소/시간축 레이어만 상위 헤더로 허용한다.
    if (!selectedAsHeader && !sparseHeaderLayer && !temporalLayer) break;

    rows.unshift(r);
  }

  return rows;
}

function expandHeaderRowIndexesForMergeFill(json = [], headerRowIndexes = []) {
  const expanded = new Set();

  for (const idx of headerRowIndexes || []) {
    for (const rowIndex of getFlattenHeaderBandRows(json, idx, 4, {
      headerRowIndexes,
    })) {
      expanded.add(rowIndex);
    }
  }

  return [...expanded].sort((a, b) => a - b);
}

function shouldSkipUpperHeaderLayer(
  json = [],
  headerIndex,
  laterHeaderIndexes = [],
  allHeaderRowIndexes = [],
) {
  return laterHeaderIndexes.some((laterIndex) => {
    if (laterIndex <= headerIndex) return false;
    if (laterIndex - headerIndex > 4) return false;
    return getFlattenHeaderBandRows(json, laterIndex, 4, {
      headerRowIndexes: allHeaderRowIndexes,
    }).includes(headerIndex);
  });
}

function effectiveHeaderRowIndexes(json = [], headerRowIndexes = []) {
  const sorted = [...headerRowIndexes].sort((a, b) => a - b);

  return sorted.filter((idx, pos) => {
    const later = sorted.slice(pos + 1);
    return !shouldSkipUpperHeaderLayer(json, idx, later, sorted);
  });
}

function horizontallyFillHeaderLayer(row = [], minCol = 0, maxCol = 0) {
  const out = [];
  let carry = "";

  for (let c = minCol; c <= maxCol; c += 1) {
    const value = normalizeHeaderText(row[c]);
    if (value) carry = value;
    out[c] = value || carry || "";
  }

  return out;
}

function directHeaderLayer(row = [], minCol = 0, maxCol = 0) {
  const out = [];
  for (let c = minCol; c <= maxCol; c += 1) {
    out[c] = normalizeHeaderText(row[c]);
  }
  return out;
}

function mergeHeaderRows(json = [], headerRowIndexes = [], headerIndex) {
  const bandRows = getFlattenHeaderBandRows(json, headerIndex, 4, {
    headerRowIndexes,
  });
  const rows = bandRows.map((idx) => json[idx] || []);
  const indexes = rows.flatMap((row) => rowNonEmptyIndexes(row));

  if (!indexes.length) return null;

  const minCol = Math.min(...indexes);
  const maxCol = Math.max(...indexes);

  const layers = rows.map((row, idx) =>
    idx === rows.length - 1
      ? directHeaderLayer(row, minCol, maxCol)
      : horizontallyFillHeaderLayer(row, minCol, maxCol),
  );

  const merged = [];
  const partsByColumn = {};

  for (let c = minCol; c <= maxCol; c += 1) {
    const parts = layers.map((layer) => layer[c]).filter(Boolean);
    const header = joinHeaderParts(parts);

    if (!header) continue;
    merged[c] = header;
    partsByColumn[c] = uniqueHeaderParts(parts);
  }

  const mergedNonEmpty = nonEmptyCount(merged);
  if (mergedNonEmpty < 2) return null;

  return {
    headerRows: bandRows.map((idx) => idx + 1),
    merged,
    partsByColumn,
    depth: bandRows.length,
    strategy:
      bandRows.length > 1 ? "merged_header_flatten_v2" : "single_header_row",
  };
}

function normalizedHeaderSignatureValue(value = "") {
  return normalizeHeaderCompare(value)
    .replace(/column\d+$/i, "")
    .trim();
}

function buildHeaderSignatureFromRow(row = [], headerCols = []) {
  return (headerCols || [])
    .map((col) => normalizedHeaderSignatureValue(row?.[col]))
    .filter(Boolean);
}

function buildHeaderSignatureInfo(
  json = [],
  headerRowIndexes = [],
  headerIndex = 0,
) {
  const mergedHeaderInfo = mergeHeaderRows(json, headerRowIndexes, headerIndex);
  const headerRow = json[headerIndex] || [];
  const effectiveHeaderRow = mergedHeaderInfo?.merged || headerRow;
  const headerCols = rowNonEmptyIndexes(effectiveHeaderRow);

  return {
    headerIndex,
    headerRows: mergedHeaderInfo?.headerRows?.length
      ? mergedHeaderInfo.headerRows.map((row) => row - 1)
      : [headerIndex],
    effectiveHeaderRow,
    headerCols,
    signature: buildHeaderSignatureFromRow(effectiveHeaderRow, headerCols),
  };
}

function headerSignatureSimilarity(left = [], right = []) {
  const a = (left || []).filter(Boolean);
  const b = (right || []).filter(Boolean);
  if (!a.length || !b.length) return 0;

  const bSet = new Set(b);
  const matched = a.filter((value) => bSet.has(value)).length;
  return matched / Math.max(a.length, b.length);
}

function looksLikeSameHeaderSignature(left = [], right = []) {
  const a = (left || []).filter(Boolean);
  const b = (right || []).filter(Boolean);
  if (a.length < 2 || b.length < 2) return false;

  const similarity = headerSignatureSimilarity(a, b);
  if (similarity >= 0.86) return true;

  const leadingComparable = Math.min(6, a.length, b.length);
  let leadingMatched = 0;
  for (let i = 0; i < leadingComparable; i += 1) {
    if (a[i] && b[i] && a[i] === b[i]) leadingMatched += 1;
  }

  if (leadingComparable >= 4 && leadingMatched / leadingComparable >= 0.75) {
    return true;
  }

  // 넓은 표는 반복 헤더가 일부 컬럼만 비어도 같은 헤더로 본다.
  return a.length >= 8 && b.length >= 8 && similarity >= 0.78;
}

function hasStrongHeaderBoundaryBefore(json = [], headerStartIndex = 0) {
  for (
    let r = headerStartIndex - 1;
    r >= Math.max(0, headerStartIndex - 3);
    r -= 1
  ) {
    const row = json[r] || [];
    if (isBlankRow(row)) return true;
    if (looksLikeNoteOrCommentRow(row)) return true;
    if (looksLikeSectionTitleRow(row)) return true;

    // 제목/섹션명처럼 텍스트 1~2개만 있는 행이 바로 앞에 있으면 별도 표일 수 있다.
    if (numericDateCount(row) === 0 && nonEmptyCount(row) <= 2) return true;

    if (nonEmptyCount(row) >= 2) return false;
  }

  return false;
}

function isMissingValuePlaceholderCell(value) {
  const s = compactCellText(value);
  if (!s) return false;
  return /^(?:[-–—]+|[.]+|N\/A|NA|NULL|없음|미상)$/i.test(s);
}

function isGenericDimensionHeaderToken(value) {
  const s = normalizeNoiseText(value);
  return /^(구분|분류|항목|유형|종류|지역|국가|국가별|산업|산업별|성별|연령|연령별|기간|연도|월|분기|category|item|type|group|class)$/.test(
    s,
  );
}

function countDataValueSignals(cells = []) {
  return (cells || []).filter((cell) => {
    if (cell == null || String(cell).trim() === "") return false;
    return (
      isNumericLike(cell) ||
      isDateLike(cell) ||
      isTimeLike(cell) ||
      isMissingValuePlaceholderCell(cell)
    );
  }).length;
}

function looksLikeContinuationDataCandidateRow(row = []) {
  const firstIdx = firstNonEmptyIndex(row);
  if (firstIdx < 0) return false;

  const cells = nonEmptyCells(row);
  const nonEmpty = cells.length;
  if (nonEmpty < 3) return false;

  const temporalCount = cells.filter(isTemporalHeaderLike).length;
  if (temporalCount >= 2) return false;

  const first = row[firstIdx];
  const firstText = compactCellText(first);
  if (!firstText) return false;

  const firstIsDataValue =
    isNumericLike(first) ||
    isDateLike(first) ||
    isTimeLike(first) ||
    isMissingValuePlaceholderCell(first) ||
    isTemporalHeaderLike(first);
  if (firstIsDataValue) return false;

  const afterFirst = row
    .slice(firstIdx + 1)
    .filter((cell) => cell != null && String(cell).trim() !== "");
  if (afterFirst.length < 2) return false;

  const dataSignals = countDataValueSignals(afterFirst);
  const placeholderSignals = afterFirst.filter(
    isMissingValuePlaceholderCell,
  ).length;
  const signalRatio = dataSignals / afterFirst.length;

  // blank row 뒤에 이어지는 실제 데이터 행이 헤더 후보로 승격되는 케이스 방지.
  // 예: ["지붕", "-", "-", ...], ["비상구", 45, 23, ...]
  // 반대로 ["구분", "2021", "2022", "2023"] 같은 진짜 기간형 헤더는 temporalCount로 제외된다.
  if (placeholderSignals > 0 && signalRatio >= 0.45) return true;

  if (!isGenericDimensionHeaderToken(first) && signalRatio >= 0.65) {
    return true;
  }

  return false;
}

function headerCandidateLooksLikeDataBlock(json = [], candidateInfo = {}) {
  const rows = (candidateInfo.headerRows || []).map((idx) => json[idx] || []);
  if (!rows.length) return false;

  const continuationDataLikeRows = rows.filter(
    looksLikeContinuationDataCandidateRow,
  ).length;
  if (continuationDataLikeRows >= Math.ceil(rows.length / 2)) {
    return true;
  }

  let rawNumericDateCount = 0;
  let rawTemporalCount = 0;
  let rawNonEmptyCount = 0;

  for (const row of rows) {
    rawNumericDateCount += numericDateCount(row);
    rawTemporalCount += countTemporalHeaderTokens(row);
    rawNonEmptyCount += nonEmptyCount(row);
  }

  // 데이터 행 1~2개가 헤더 후보로 묶인 경우는 숫자값이 많지만
  // 2024 같은 시간축 헤더 토큰은 거의 없다.
  if (rawNonEmptyCount > 0) {
    const rawNumericRatio = rawNumericDateCount / rawNonEmptyCount;
    if (
      rawNumericDateCount >= 3 &&
      rawNumericRatio >= 0.25 &&
      rawTemporalCount <= 2
    ) {
      return true;
    }
  }

  const dataLikeRows = rows.filter((row) => {
    const nonEmpty = nonEmptyCount(row);
    if (!nonEmpty) return false;
    const numericDate = numericDateCount(row);
    const numericDateRatio = numericDate / nonEmpty;
    const temporalCount = countTemporalHeaderTokens(row);

    if (isLikelyDataRow(row)) return true;
    if (numericDateRatio >= 0.25 && temporalCount === 0) {
      return true;
    }
    return false;
  }).length;

  if (!dataLikeRows) return false;
  return dataLikeRows >= Math.ceil(rows.length / 2);
}

function resolveNextDistinctHeaderInfo({
  json = [],
  selectedHeaders = [],
  currentPosition = 0,
  currentSignature = [],
  allHeaderRowIndexes = [],
} = {}) {
  const repeatedHeaderBands = [];
  let nextHeaderIndex = json.length;

  for (let j = currentPosition + 1; j < selectedHeaders.length; j += 1) {
    const candidateIndex = selectedHeaders[j];
    const candidateInfo = buildHeaderSignatureInfo(
      json,
      allHeaderRowIndexes,
      candidateIndex,
    );
    const candidateStart = Math.min(...candidateInfo.headerRows);
    const sameHeader = looksLikeSameHeaderSignature(
      currentSignature,
      candidateInfo.signature,
    );

    if (sameHeader && !hasStrongHeaderBoundaryBefore(json, candidateStart)) {
      repeatedHeaderBands.push({
        headerIndex: candidateIndex,
        rows: candidateInfo.headerRows,
        signatureSimilarity: Number(
          headerSignatureSimilarity(
            currentSignature,
            candidateInfo.signature,
          ).toFixed(3),
        ),
      });
      continue;
    }

    // 원래 헤더 후보로 잡혔더라도 실제 데이터/소계 행으로 보이면
    // 새 table boundary로 사용하지 않는다.
    if (headerCandidateLooksLikeDataBlock(json, candidateInfo)) {
      continue;
    }

    nextHeaderIndex = candidateIndex;
    break;
  }

  const repeatedHeaderRowSet = new Set(
    repeatedHeaderBands.flatMap((band) => band.rows || []),
  );

  return {
    nextHeaderIndex,
    repeatedHeaderBands,
    repeatedHeaderRowSet,
  };
}

function cloneJsonRows(json = []) {
  return json.map((row) => (Array.isArray(row) ? [...row] : []));
}

function fillMergedCellsForHeaderRows(
  json = [],
  merges = [],
  headerRowIndexes = [],
) {
  const out = cloneJsonRows(json);
  if (!Array.isArray(merges) || !merges.length) return out;

  const headerSet = new Set(headerRowIndexes);

  for (const m of merges) {
    const startRow = m?.s?.r;
    const endRow = m?.e?.r;
    const startCol = m?.s?.c;
    const endCol = m?.e?.c;

    if (
      startRow == null ||
      endRow == null ||
      startCol == null ||
      endCol == null
    ) {
      continue;
    }

    // 헤더 후보 행과 겹치는 병합셀만 처리
    let touchesHeader = false;
    for (let r = startRow; r <= endRow; r++) {
      if (headerSet.has(r)) {
        touchesHeader = true;
        break;
      }
    }
    if (!touchesHeader) continue;

    const source = out[startRow]?.[startCol];
    if (source == null || String(source).trim() === "") continue;

    for (let r = startRow; r <= endRow; r++) {
      if (!headerSet.has(r)) continue;

      if (!out[r]) out[r] = [];
      for (let c = startCol; c <= endCol; c++) {
        if (out[r][c] == null || String(out[r][c]).trim() === "") {
          out[r][c] = source;
        }
      }
    }
  }

  return out;
}

function detectTableBlocks(
  json = [],
  headerRowIndexes = [],
  sheetName = "",
  options = {},
) {
  const blocks = [];
  if (!Array.isArray(json) || !Array.isArray(headerRowIndexes)) return blocks;

  const segmentation =
    options.segmentation ||
    buildSegmentationHeaderIndexes(json, headerRowIndexes, options);
  const segmentationHeaderRowIndexes =
    segmentation.headerRowIndexes || headerRowIndexes;
  const selectedHeaders = effectiveHeaderRowIndexes(
    json,
    segmentationHeaderRowIndexes,
  );
  let previousBlockEnd = -1;

  for (let i = 0; i < selectedHeaders.length; i++) {
    const headerIndex = selectedHeaders[i];
    if (headerIndex <= previousBlockEnd) continue;

    const headerRow = json[headerIndex] || [];
    const mergedHeaderInfo = mergeHeaderRows(
      json,
      segmentationHeaderRowIndexes,
      headerIndex,
    );
    const effectiveHeaderRow = mergedHeaderInfo?.merged || headerRow;
    const headerCols = rowNonEmptyIndexes(effectiveHeaderRow);
    const headerStartIndex = mergedHeaderInfo?.headerRows?.length
      ? mergedHeaderInfo.headerRows[0] - 1
      : headerIndex;

    if (headerCols.length < 2) continue;

    const minCol = Math.min(...headerCols);
    const maxCol = Math.max(...headerCols);
    const baseContext = {
      headerRow: effectiveHeaderRow,
      headerCols,
      minCol,
      maxCol,
    };

    const titleRows = findPrecedingSectionTitleRows(
      json,
      headerStartIndex,
      baseContext,
    );
    const currentSignature = buildHeaderSignatureFromRow(
      effectiveHeaderRow,
      headerCols,
    );
    const nextHeaderInfo = resolveNextDistinctHeaderInfo({
      json,
      selectedHeaders,
      currentPosition: i,
      currentSignature,
      allHeaderRowIndexes: segmentationHeaderRowIndexes,
    });
    const nextHeaderIndex = nextHeaderInfo.nextHeaderIndex;
    const repeatedHeaderRowSet = nextHeaderInfo.repeatedHeaderRowSet;

    let dataStart = headerIndex + 1;
    const excludedRows = [];
    const summaryRows = [];

    while (dataStart < Math.min(nextHeaderIndex, json.length)) {
      const classified = classifyTableRow(json[dataStart] || [], baseContext);
      if (classified.include) break;

      const excludedRow = {
        row: dataStart + 1,
        reason: classified.reason,
        phase: "beforeDataStart",
        kind: classified.kind,
      };
      excludedRows.push(excludedRow);
      if (classified.kind === "total") {
        summaryRows.push({
          ...excludedRow,
          detectorVersion: SUMMARY_ROW_DETECTOR_VERSION,
          rawCells: rowSliceByColumnRange(
            json[dataStart] || [],
            minCol,
            maxCol,
          ),
        });
      }
      dataStart += 1;
    }

    if (dataStart >= json.length || dataStart >= nextHeaderIndex) continue;

    let dataEnd = dataStart - 1;
    let blankStreak = 0;
    let stopReason = "NEXT_HEADER_OR_EOF";

    for (let r = dataStart; r < Math.min(nextHeaderIndex, json.length); r++) {
      const row = json[r] || [];

      if (repeatedHeaderRowSet.has(r)) {
        excludedRows.push({
          row: r + 1,
          reason: "REPEATED_HEADER_ROW",
          phase: "insideDataRange",
          kind: "repeatedHeader",
        });
        blankStreak = 0;
        continue;
      }

      const classified = classifyTableRow(row, baseContext);

      if (classified.include) {
        blankStreak = 0;
        dataEnd = r;
        continue;
      }

      const excludedRow = {
        row: r + 1,
        reason: classified.reason,
        phase: "insideDataRange",
        kind: classified.kind,
      };
      excludedRows.push(excludedRow);
      if (classified.kind === "total") {
        summaryRows.push({
          ...excludedRow,
          detectorVersion: SUMMARY_ROW_DETECTOR_VERSION,
          rawCells: rowSliceByColumnRange(row, minCol, maxCol),
        });
      }

      if (classified.kind === "blank") blankStreak += 1;

      const hasFuture = hasFutureDataRow(
        json,
        r + 1,
        Math.min(
          nextHeaderIndex,
          r + 1 + TABLE_SEGMENTATION_OPTIONS.maxLookaheadRows,
        ),
        baseContext,
      );

      if (blankStreak >= 2 && !hasFuture) {
        stopReason = "MULTI_BLANK_BOUNDARY";
        break;
      }

      if (shouldStopAtExcludedRow({ classified, hasFuture })) {
        stopReason = classified.reason || "EXCLUDED_ROW_BOUNDARY";
        break;
      }
    }

    if (dataEnd < dataStart) continue;

    // summary_row_detector_v2_1:
    // 마지막 행이 소계/총계로 분류되어 rows에서는 제외되더라도,
    // summaryRows에는 보존되므로 table range/dataRange도 해당 행까지 포함해야 한다.
    // BLANK_ROW 같은 단순 경계 행은 기존처럼 range 밖에 둘 수 있도록 total 행만 확장 대상으로 삼는다.
    const summaryEnd = summaryRows.reduce((maxRowIndex, row) => {
      if (row.phase !== "insideDataRange") return maxRowIndex;
      const rowIndex = Number(row.row || 0) - 1;
      return Number.isFinite(rowIndex)
        ? Math.max(maxRowIndex, rowIndex)
        : maxRowIndex;
    }, dataEnd);
    const blockEnd = Math.max(dataEnd, summaryEnd);

    previousBlockEnd = Math.max(previousBlockEnd, blockEnd);

    const columns = headerCols
      .map((idx) => {
        const header = String(effectiveHeaderRow[idx] || "").trim();
        const originalHeader = String(headerRow[idx] || "").trim();
        if (!header) return null;

        return {
          header,
          originalHeader: originalHeader || header,
          columnIndex: idx + 1,
          columnLetter: indexToColumnLetter(idx),
        };
      })
      .filter(Boolean);

    const excludedInsideRange = excludedRows.filter(
      (row) => row.row >= dataStart + 1 && row.row <= blockEnd + 1,
    );
    const includedRowCount = Math.max(
      0,
      blockEnd - dataStart + 1 - excludedInsideRange.length,
    );
    const addedBySegmentation = (segmentation.addedRows || []).some(
      (row) => row.rowIndex === headerIndex,
    );

    blocks.push({
      tableId: `${sheetName || "Sheet"}#T${blocks.length + 1}`,
      sheetName,
      tableTitle: titleRows.map((row) => row.text).join(" / ") || "",
      headerRow: headerIndex + 1,
      headerRows: mergedHeaderInfo?.headerRows || [headerIndex + 1],
      hasMergedHeader: Boolean(mergedHeaderInfo && mergedHeaderInfo.depth > 1),
      headerFlattenStrategy: mergedHeaderInfo?.strategy || "single_header_row",
      dataStartRow: dataStart + 1,
      dataEndRow: blockEnd + 1,
      startCol: indexToColumnLetter(minCol),
      endCol: indexToColumnLetter(maxCol),
      startColIndex: minCol + 1,
      endColIndex: maxCol + 1,
      range: `'${sheetName}'!${indexToColumnLetter(minCol)}${headerStartIndex + 1}:${indexToColumnLetter(maxCol)}${blockEnd + 1}`,
      dataRange: `'${sheetName}'!${indexToColumnLetter(minCol)}${dataStart + 1}:${indexToColumnLetter(maxCol)}${blockEnd + 1}`,
      columns,
      headerPartsByColumn: mergedHeaderInfo?.partsByColumn || {},
      excludedRows,
      summaryRows,
      dataQuality: {
        version: "table_data_row_filter_v1",
        includedRowCount,
        excludedRowCount: excludedRows.length,
        summaryRowCount: summaryRows.length,
        excludedReasonCounts: summarizeExcludedRows(excludedRows),
      },
      tableSegmentation: {
        version: TABLE_SEGMENTATION_VERSION,
        addedBySegmentation,
        titleRows,
        repeatedHeaderBands: (nextHeaderInfo.repeatedHeaderBands || []).map(
          (band) => ({
            headerRow: band.headerIndex + 1,
            rows: (band.rows || []).map((row) => row + 1),
            signatureSimilarity: band.signatureSimilarity,
          }),
        ),
        previousBlockEnd: previousBlockEnd >= 0 ? previousBlockEnd + 1 : null,
        nextHeaderRow:
          nextHeaderIndex < json.length ? nextHeaderIndex + 1 : null,
        stopReason,
      },
      score:
        scoreHeaderRowCandidate(effectiveHeaderRow) +
        Math.min(Math.max(includedRowCount, 0), 50) +
        (titleRows.length ? 3 : 0) +
        (addedBySegmentation ? -2 : 0),
    });
  }

  detectTableBlocks.lastSegmentation = segmentation;
  return blocks;
}

// workbook (XLSX.read 결과) → allSheetsData
function buildAllSheetsData(workbook) {
  const allSheetsData = {};

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws || !ws["!ref"]) continue;

    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rowCount = range.e.r + 1;

    // 2차원 배열 (각 원소가 행)
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (!json || json.length < 1) continue;

    // 1) 시트 전체에서 "값이 있는 첫 행 / 마지막 행" 찾기
    let firstNonEmpty = null;
    let lastNonEmpty = null;
    for (let i = 0; i < json.length; i++) {
      const row = json[i] || [];
      if (row.some((c) => c != null && String(c).trim() !== "")) {
        if (firstNonEmpty === null) firstNonEmpty = i;
        lastNonEmpty = i;
      }
    }
    if (firstNonEmpty === null) continue;

    // 2) 헤더처럼 보이는 모든 행 인덱스를 수집
    //    - 고정 키워드가 아니라 행 자체의 구조 + 다음 행의 데이터 증거로 판정
    //    - 숫자형 연도/월/분기 헤더는 temporal header token으로 인정
    const headerScan = [];
    const headerRowIndexes = [];
    for (let i = firstNonEmpty; i <= lastNonEmpty; i++) {
      const row = json[i] || [];
      const analysis = analyzeHeaderRowCandidate(row, {
        rowIndex: i,
        prevRow: i > firstNonEmpty ? json[i - 1] || [] : null,
        nextRows: json.slice(i + 1, Math.min(json.length, i + 5)),
      });

      if (!analysis.nonEmpty) continue;

      const item = {
        rowIndex: i,
        ...analysis,
      };
      headerScan.push(item);

      if (analysis.isCandidate) {
        headerRowIndexes.push(i);
      }
    }

    // 보수적 fallback: strict 기준을 통과한 행이 없으면 가장 높은 후보 1개만 채택한다.
    // 이 경우도 점수/구조 증거가 최소 기준 이상이어야 한다.
    if (headerRowIndexes.length === 0) {
      const fallback = headerScan
        .filter(
          (c) =>
            c.nonEmpty >= 2 &&
            c.score >= 24 &&
            !c.reasons.includes("META_OR_PROSE_ROW_PENALTY"),
        )
        .sort((a, b) => b.score - a.score)[0];

      if (fallback) headerRowIndexes.push(fallback.rowIndex);
    }

    if (headerRowIndexes.length === 0) continue;

    const headerCandidates = headerRowIndexes
      .map((idx) => {
        const row = json[idx] || [];
        const scanned = headerScan.find((item) => item.rowIndex === idx);
        return {
          rowIndex: idx,
          score:
            scanned?.score ??
            scoreHeaderRowCandidate(row, {
              rowIndex: idx,
              prevRow: idx > firstNonEmpty ? json[idx - 1] || [] : null,
              nextRows: json.slice(idx + 1, Math.min(json.length, idx + 5)),
            }),
          nonEmpty: scanned?.nonEmpty ?? nonEmptyCount(row),
          textLike: scanned?.textLike ?? textLikeCount(row),
          textRatio: scanned?.textRatio ?? null,
          headerLikeRatio: scanned?.headerLikeRatio ?? null,
          temporalCount: scanned?.temporalCount ?? 0,
          numericDateRatio: scanned?.numericDateRatio ?? null,
          nextDataEvidence: scanned?.nextDataEvidence ?? null,
          fillRatio: scanned?.fillRatio ?? null,
          reasons: scanned?.reasons || [],
        };
      })
      .sort((a, b) => b.score - a.score);

    const bestHeaderCandidate = headerCandidates[0] || null;
    const initialMergeFillHeaderRows = expandHeaderRowIndexesForMergeFill(
      json,
      headerRowIndexes,
    );
    let headerJson = fillMergedCellsForHeaderRows(
      json,
      ws["!merges"] || [],
      initialMergeFillHeaderRows,
    );
    let tableSegmentation = buildSegmentationHeaderIndexes(
      headerJson,
      headerRowIndexes,
      {
        firstNonEmpty,
        lastNonEmpty,
      },
    );
    const mergeFillHeaderRows = expandHeaderRowIndexesForMergeFill(
      headerJson,
      tableSegmentation.headerRowIndexes,
    );
    headerJson = fillMergedCellsForHeaderRows(
      json,
      ws["!merges"] || [],
      mergeFillHeaderRows,
    );
    tableSegmentation = buildSegmentationHeaderIndexes(
      headerJson,
      headerRowIndexes,
      {
        firstNonEmpty,
        lastNonEmpty,
      },
    );
    const selectedHeaderRowIndexes = effectiveHeaderRowIndexes(
      headerJson,
      tableSegmentation.headerRowIndexes,
    );
    const tableBlocks = detectTableBlocks(
      headerJson,
      tableSegmentation.headerRowIndexes,
      sheetName,
      { segmentation: tableSegmentation },
    );

    const blockByHeaderRow = new Map();
    for (const block of tableBlocks) {
      for (const rowNo of block.headerRows || [block.headerRow]) {
        blockByHeaderRow.set(Number(rowNo), block);
      }
    }

    // 3) 헤더처럼 보이는 각 행에서 metaData 채우기
    const metaData = {};

    for (const headerIndex of selectedHeaderRowIndexes) {
      const mergedHeaderInfo = mergeHeaderRows(
        headerJson,
        tableSegmentation.headerRowIndexes,
        headerIndex,
      );
      const headers = mergedHeaderInfo?.merged || headerJson[headerIndex] || [];
      const relatedBlock = blockByHeaderRow.get(headerIndex + 1) || null;
      const excludedRowSet = new Set(
        (relatedBlock?.excludedRows || []).map((row) => Number(row.row)),
      );

      let dataStart = relatedBlock
        ? Number(relatedBlock.dataStartRow || headerIndex + 2) - 1
        : headerIndex + 1;
      let dataEnd = relatedBlock
        ? Number(relatedBlock.dataEndRow || lastNonEmpty + 1) - 1
        : lastNonEmpty;

      if (!relatedBlock) {
        for (let r = dataStart; r < json.length; r++) {
          const row = json[r] || [];
          if (row.some((c) => c != null && String(c).trim() !== "")) {
            dataStart = r;
            break;
          }
        }
      }

      headers.forEach((header, idx) => {
        const name = String(header || "").trim();
        if (!name) return;

        // 같은 이름의 헤더가 위에서 이미 등록되었으면 첫 번째 것만 사용
        if (metaData[name]) return;

        const values = [];
        const maxSampleRow = Math.min(
          json.length - 1,
          dataEnd,
          dataStart + 199,
        );
        for (let r = dataStart; r <= maxSampleRow; r++) {
          if (excludedRowSet.has(r + 1)) continue;
          const val = json[r]?.[idx];
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            values.push(val);
          }
        }

        const stats = analyzeSamples(values);

        const sampleValues = values
          .slice(0, 5)
          .map((v) => String(v).trim())
          .filter(Boolean);
        const profile = buildColumnProfile(name, values, stats);
        const clusterCandidate = inferClusterCandidate(
          name,
          sampleValues,
          stats.dominantType,
        );
        const inferredRole = getClusterRole(clusterCandidate);
        const clusterType = inferClusterType(
          clusterCandidate,
          name,
          sampleValues,
          stats.dominantType,
        );

        metaData[name] = {
          columnLetter: indexToColumnLetter(idx),
          startRow: dataStart + 1,
          lastRow: dataEnd + 1,
          headerRow: headerIndex + 1,
          headerRows: mergedHeaderInfo?.headerRows || [headerIndex + 1],
          isFlattenedHeader: Boolean(
            mergedHeaderInfo && mergedHeaderInfo.depth > 1,
          ),
          headerParts: mergedHeaderInfo?.partsByColumn?.[idx] || [name],
          excludedRows: relatedBlock?.excludedRows || [],
          columnIndex: idx + 1,
          canonicalKey: clusterCandidate || null,
          sampleValues,
          profileType: profile.profileType,
          uniqueCount: profile.uniqueCount,
          uniqueRatio: profile.uniqueRatio,
          uniqueValues: profile.uniqueValues,
          clusterCandidate,
          inferredRole,
          clusterType,
          ...stats,
        };
      });
    }

    if (Object.keys(metaData).length === 0) continue;

    // 4) 전체 범위 정보 (fallback 용도)
    const firstHeaderRow = selectedHeaderRowIndexes.length
      ? Math.min(...selectedHeaderRowIndexes)
      : Math.min(...tableSegmentation.headerRowIndexes);
    const startRow = firstHeaderRow + 2; // 0-based → 1-based
    const lastDataRow = tableBlocks.length
      ? Math.max(...tableBlocks.map((block) => Number(block.dataEndRow || 0)))
      : lastNonEmpty + 1;

    const dataQuality = {
      version: "sheet_data_row_filter_v1",
      tableCount: tableBlocks.length,
      excludedRowCount: tableBlocks.reduce(
        (sum, block) => sum + Number(block.excludedRows?.length || 0),
        0,
      ),
      excludedReasonCounts: tableBlocks.reduce((acc, block) => {
        const counts = block.dataQuality?.excludedReasonCounts || {};
        for (const [reason, count] of Object.entries(counts)) {
          acc[reason] = (acc[reason] || 0) + count;
        }
        return acc;
      }, {}),
      segmentation: {
        version: TABLE_SEGMENTATION_VERSION,
        detectedTableCount: tableBlocks.length,
        addedHeaderRows: (tableSegmentation.addedRows || []).map((row) => ({
          row: row.rowIndex + 1,
          score: row.score,
          nextDataEvidence: row.nextDataEvidence,
          preview: row.preview,
        })),
        removedHeaderRows: (tableSegmentation.removedRows || []).map((row) => ({
          row: row.rowIndex + 1,
          source: row.source,
          reason: row.reason,
          preview: row.preview,
        })),
        titleRowCount: tableBlocks.reduce(
          (sum, block) =>
            sum + Number(block.tableSegmentation?.titleRows?.length || 0),
          0,
        ),
      },
    };

    allSheetsData[sheetName] = {
      sheetName,
      rowCount,
      startRow,
      lastDataRow,
      metaData,
      tableBlocks,
      dataQuality,
      headerDetection: {
        version: "header_detector_v2_score_based",
        scannedRows: headerScan.length,
        selectedRows: tableSegmentation.headerRowIndexes.map((idx) => idx + 1),
        originallyDetectedRows: headerRowIndexes.map((idx) => idx + 1),
        segmentationAddedRows: (tableSegmentation.addedRows || []).map(
          (row) => ({
            row: row.rowIndex + 1,
            score: row.score,
            nextDataEvidence: row.nextDataEvidence,
            preview: row.preview,
          }),
        ),
        segmentationRemovedRows: (tableSegmentation.removedRows || []).map(
          (row) => ({
            row: row.rowIndex + 1,
            source: row.source,
            reason: row.reason,
            preview: row.preview,
          }),
        ),
        effectiveRows: selectedHeaderRowIndexes.map((idx) => idx + 1),
        mergeFillRows: mergeFillHeaderRows.map((idx) => idx + 1),
        topCandidates: headerScan
          .slice()
          .sort((a, b) => b.score - a.score)
          .slice(0, 10)
          .map((item) => ({
            rowIndex: item.rowIndex + 1,
            score: item.score,
            nonEmpty: item.nonEmpty,
            textRatio: item.textRatio,
            headerLikeRatio: item.headerLikeRatio,
            temporalCount: item.temporalCount,
            numericDateRatio: item.numericDateRatio,
            nextDataEvidence: item.nextDataEvidence,
            reasons: item.reasons,
          })),
      },
      headerCandidates,
      bestHeaderRow:
        bestHeaderCandidate?.rowIndex != null
          ? bestHeaderCandidate.rowIndex + 1
          : startRow - 1,
    };
  }

  return allSheetsData;
}

module.exports = { buildAllSheetsData, detectCellType };
