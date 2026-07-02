const XLSX = require("xlsx");
const { buildNarrativeSections } = require("./reportNarrativeBuilder");
const { recommendChartSpec } = require("./chartRecommendationBuilder");
const { buildReportSections } = require("./reportSectionBuilder");
const {
  FORMULA_SPEC_TYPES,
  buildColumnRange,
  buildGroupAggregateFormula,
  buildRankValueFormula,
  buildRankLabelFormula,
  buildRunningSumFormula,
  buildGrowthRateFormula,
  buildMaxIfFormula,
  buildCountIfsFormula,
  buildPivotAverageFormula,
  buildFormulaFromSpec,
  createFormulaCellFromSpec,
} = require("./formulaEngine/internalFormulaEngine");
const {
  SHEET_NAMES,
  normalizeSummarySheetMode,
  isFormulaEnabledMode,
  sourceSheetNameForTableIndex,
} = require("./config/automationSheetConfig");
const {
  getSourceColumnHeader,
  createSourceColumnMap,
  resolveSourceColumn,
} = require("./utils/headerMatcher");
const {
  isRankingLikeSection,
  normalizeAggregateOperation,
  isSimpleAggregateOperation,
  resolveCriteriaColumnIndex,
  resolveAggregateFormulaTargets,
} = require("./utils/aggregateResolver");

function buildChartDataRows(result = {}) {
  if (result.resultType === "grouped") {
    if (
      result.resultType === "grouped" &&
      (result.operation === "multiAggregate" ||
        result.operation === "pipelineCombine")
    ) {
      return result.rows || [];
    }
    const groupHeader = result.groupBy?.header || "그룹";
    const metricHeader = result.metric?.header || "값";

    return (result.rows || []).map((r) => ({
      [groupHeader]: r[groupHeader] ?? "",
      [metricHeader]: r.value,
      행수: r.rowCount,
    }));
  }

  if (result.resultType === "pivot") {
    return result.rows || [];
  }

  return [];
}

function sanitizeSheetName(name) {
  return (
    String(name || "Sheet")
      .replace(/[:\\/?*\[\]]/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 31) || "Sheet"
  );
}

function appendSheetSafe(wb, ws, name) {
  let safeName = sanitizeSheetName(name);
  const existing = new Set(wb.SheetNames || []);

  if (!existing.has(safeName)) {
    XLSX.utils.book_append_sheet(wb, ws, safeName);
    return;
  }

  let idx = 2;
  while (existing.has(`${safeName}_${idx}`.slice(0, 31))) {
    idx += 1;
  }

  XLSX.utils.book_append_sheet(wb, ws, `${safeName}_${idx}`.slice(0, 31));
}

function objectToAoa(obj = {}, prefix = "") {
  const rows = [];

  function walk(value, keyPath) {
    if (value == null || typeof value !== "object") {
      rows.push([keyPath, value ?? ""]);
      return;
    }

    if (Array.isArray(value)) {
      value.forEach((item, idx) => {
        walk(item, `${keyPath}[${idx}]`);
      });
      return;
    }

    for (const [key, child] of Object.entries(value)) {
      walk(child, keyPath ? `${keyPath}.${key}` : key);
    }
  }

  walk(obj, prefix);

  return rows.length ? rows : [["", ""]];
}

function buildChartSpec(result = {}) {
  return recommendChartSpec(result);
}

function buildInsightRows(result = {}) {
  const rows = [];

  if (!Array.isArray(result.rows) || !result.rows.length) {
    return [["요약", "분석 결과가 없습니다."]];
  }

  rows.push([
    "요약",
    `${result.operation || "분석"} 결과 ${result.rows.length}건이 생성되었습니다.`,
  ]);

  if (result.resultType === "pivot") {
    rows.push(["분석유형", "Pivot 교차 분석 결과입니다."]);
    rows.push(["행 기준", result.pivot?.rowGroup?.header || ""]);
    rows.push(["열 기준", result.pivot?.columnGroup?.header || ""]);
    rows.push(["열 항목 수", result.pivot?.columns?.length || 0]);
    rows.push(["결과 행 수", result.rows?.length || 0]);

    return rows;
  }

  if (result.resultType === "grouped") {
    const groupHeader = result.groupBy?.header || "그룹";
    const valueRows = result.rows
      .filter((r) => Number.isFinite(Number(r.value)))
      .map((r) => ({
        label: r[groupHeader],
        value: Number(r.value),
      }));

    if (valueRows.length) {
      const max = valueRows.reduce((a, b) => (b.value > a.value ? b : a));
      const min = valueRows.reduce((a, b) => (b.value < a.value ? b : a));

      rows.push(["최대값", `${max.label}: ${max.value}`]);
      rows.push(["최소값", `${min.label}: ${min.value}`]);
    }

    const growthRows = result.rows.filter((r) =>
      Number.isFinite(Number(r["증감률"])),
    );

    if (growthRows.length) {
      const maxGrowth = growthRows.reduce((a, b) =>
        Number(b["증감률"]) > Number(a["증감률"]) ? b : a,
      );

      rows.push([
        "최대 증감률",
        `${maxGrowth[groupHeader]}: ${Number(maxGrowth["증감률"]).toFixed(2)}%`,
      ]);
    }
  }

  return rows;
}

function formatKstTimestamp(date = new Date()) {
  const kst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  const y = kst.getUTCFullYear();
  const m = String(kst.getUTCMonth() + 1).padStart(2, "0");
  const d = String(kst.getUTCDate()).padStart(2, "0");
  const hh = String(kst.getUTCHours()).padStart(2, "0");
  const mm = String(kst.getUTCMinutes()).padStart(2, "0");
  return `${y}-${m}-${d} ${hh}:${mm} KST`;
}

function sectionRowsForSummary(section = {}) {
  const rows = resultToRows(section.result || {});
  return Array.isArray(rows) ? rows : [];
}

function buildBusinessWorkbookSummaryRows({
  fileName,
  message,
  result,
  businessSections = [],
  summarySheetMode = "static",
  includeSourceDataSheet = true,
  sourceTables = [],
} = {}) {
  const sectionRows = businessSections.map((section, index) => {
    const rows = sectionRowsForSummary(section);
    const chartSpec = buildChartSpec(section.result || {});
    return {
      no: index + 1,
      title: section.title || section.sectionId || `섹션_${index + 1}`,
      rowCount: rows.length,
      chart: chartSpec?.recommendedType || "-",
      type: section.sectionType || section.result?.resultType || "-",
    };
  });

  const totalRows = sectionRows.reduce((sum, row) => sum + row.rowCount, 0);
  const chartCount = sectionRows.filter(
    (row) => row.chart && row.chart !== "-",
  ).length;

  return [
    ["항목", "내용", "비고"],
    ["요청", message || "", ""],
    ["원본 파일", fileName || "", ""],
    ["템플릿", result?.title || result?.templateId || "", ""],
    ["결과 유형", result?.resultType || "businessTemplate", ""],
    ["섹션 수", businessSections.length, ""],
    ["결과 행 수", totalRows, "전체 섹션 합산"],
    ["차트 후보", chartCount, "추천 가능한 섹션 수"],
    ["수식 모드", summarySheetMode, "static/formula/hybrid"],
    [
      "원본데이터 포함",
      includeSourceDataSheet ? "예" : "아니오",
      `${sourceTables.length || 0}개 테이블`,
    ],
    ["생성일시", formatKstTimestamp(), ""],
    [],
    ["섹션", "행 수", "차트/유형"],
    ...sectionRows.map((row) => [
      row.title,
      row.rowCount,
      `${row.chart} / ${row.type}`,
    ]),
  ];
}

function setColumnWidths(ws, rows = []) {
  if (!rows.length) return;

  const keys = Object.keys(rows[0] || {});
  ws["!cols"] = keys.map((key) => {
    const maxLen = Math.max(
      String(key || "").length,
      ...rows.map((r) => String(r?.[key] ?? "").length),
    );

    return {
      wch: Math.min(Math.max(maxLen + 2, 10), 30),
    };
  });
}

function setAoaColumnWidths(ws, aoa = []) {
  const colCount = Math.max(...aoa.map((r) => r.length), 0);

  ws["!cols"] = Array.from({ length: colCount }).map((_, colIdx) => {
    const maxLen = Math.max(
      ...aoa.map((row) => String(row?.[colIdx] ?? "").length),
    );

    return {
      wch: Math.min(Math.max(maxLen + 2, 10), 40),
    };
  });
}

function inferNumberFormat(header = "", result = {}) {
  const h = String(header || "");

  if (
    h.includes("%") ||
    h.includes("률") ||
    h.includes("비율") ||
    (result.metric?.type === "rate" && h === "값")
  ) {
    return {
      z: "0.00%",
      divideBy100: true,
    };
  }

  if (result.metric?.type === "number") {
    return {
      z: Number.isInteger(result.value) ? "#,##0" : "#,##0.00",
      divideBy100: false,
    };
  }

  return {
    z: "#,##0.00",
    divideBy100: false,
  };
}

function formatNumberCells(ws, result = {}) {
  if (!ws["!ref"]) return;

  const range = XLSX.utils.decode_range(ws["!ref"]);

  for (let r = range.s.r + 1; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];

      if (!cell || typeof cell.v !== "number") continue;

      const headerAddr = XLSX.utils.encode_cell({ r: range.s.r, c });
      const header = ws[headerAddr]?.v || "";

      const fmt = inferNumberFormat(header, result);
      cell.z = fmt.z;

      if (fmt.divideBy100) {
        cell.v = cell.v / 100;
      }
    }
  }
}

function styleHeaderRow(ws) {
  if (!ws["!ref"]) return;

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const headerRow = range.s.r;

  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = XLSX.utils.encode_cell({ r: headerRow, c });
    const cell = ws[addr];

    if (!cell) continue;

    cell.s = {
      font: { bold: true },
      alignment: { horizontal: "center", vertical: "center" },
    };
  }
}

function applyDefaultSheetOptions(ws) {
  if (!ws) return;

  ws["!freeze"] = {
    xSplit: 0,
    ySplit: 1,
  };

  if (ws["!ref"]) {
    ws["!autofilter"] = {
      ref: ws["!ref"],
    };
  }
}

function resultToRows(result = {}) {
  if (result.resultType === "grouped") {
    if (
      result.operation === "multiAggregate" ||
      result.operation === "pipelineCombine"
    ) {
      return result.rows || [];
    }
    const groupHeader = result.groupBy?.header || "그룹";
    const extraKeys = ["기준값", "비교값", "증감률"].filter((k) =>
      (result.rows || []).some((r) =>
        Object.prototype.hasOwnProperty.call(r, k),
      ),
    );

    return (result.rows || []).map((r) => {
      const base = {
        [groupHeader]: r[groupHeader] ?? "",
        작업: r.operation,
        지표: r.metric,
        값: r.value,
        행수: r.rowCount,
      };

      for (const key of extraKeys) {
        base[key] = r[key];
      }

      return base;
    });
  }

  if (result.resultType === "scalar") {
    return [
      {
        지표: result.metric?.header || result.operation,
        값: result.value,
        행수: result.rowCount,
      },
    ];
  }

  if (result.resultType === "pivot") {
    return result.rows || [];
  }

  return result.rows || [];
}

function getWorksheetHeaderInfo(ws) {
  if (!ws?.["!ref"]) return { headers: [], headerRow: 0, range: null };

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const headerRow = range.s.r;
  const headers = [];

  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = XLSX.utils.encode_cell({ r: headerRow, c });
    headers.push(String(ws[addr]?.v ?? "").trim());
  }

  return { headers, headerRow, range };
}

function applySimpleAggregateFormulaSection({
  ws,
  rows = [],
  formulaPlan = null,
  formulaContext = null,
} = {}) {
  if (!ws || !formulaPlan?.enabled || !formulaContext?.enabled) return 0;
  if (!Array.isArray(rows) || !rows.length) return 0;
  if (!formulaPlan.group?.letter) return 0;

  const operation = normalizeAggregateOperation(
    formulaPlan.operation,
    formulaPlan,
  );

  if (!isSimpleAggregateOperation(operation)) return 0;
  if (operation !== "count" && !formulaPlan.metric?.letter) return 0;

  const { headers, headerRow, range } = getWorksheetHeaderInfo(ws);
  if (!range || !headers.length) return 0;

  const criteriaColIndex = resolveCriteriaColumnIndex({
    headers,
    formulaPlan,
  });

  if (criteriaColIndex < 0) return 0;

  const targets = resolveAggregateFormulaTargets({
    headers,
    operation,
    formulaPlan,
    criteriaColIndex,
  });

  if (!targets.length) {
    formulaPlan.applied = false;
    formulaPlan.formulaCount = 0;
    formulaPlan.reason = "NO_PRECISE_FORMULA_TARGET_COLUMN";
    return 0;
  }

  const sourceSheetName =
    formulaPlan.sourceSheetName ||
    formulaContext.sourceSheetName ||
    SHEET_NAMES.SOURCE_DATA;

  let appliedCount = 0;

  for (let r = headerRow + 1; r <= range.e.r; r += 1) {
    const criteriaAddr = XLSX.utils.encode_cell({
      r,
      c: criteriaColIndex,
    });

    const criteriaValue = ws[criteriaAddr]?.v;
    if (criteriaValue == null || criteriaValue === "") continue;

    for (const target of targets) {
      if (target.operation !== "count" && !formulaPlan.metric?.letter) {
        continue;
      }

      const targetAddr = XLSX.utils.encode_cell({
        r,
        c: target.columnIndex,
      });

      const previousCell = ws[targetAddr] || {};
      const formulaCell = createFormulaCellFromSpec({
        type: FORMULA_SPEC_TYPES.GROUP_AGGREGATE,
        operation: target.operation,
        sheetName: sourceSheetName,
        groupLetter: formulaPlan.group.letter,
        metricLetter: formulaPlan.metric?.letter,
        criteriaCell: criteriaAddr,
        value: Number.isFinite(Number(previousCell.v))
          ? Number(previousCell.v)
          : 0,
        cellType: "n",
      });

      if (previousCell.z) formulaCell.z = previousCell.z;
      ws[targetAddr] = formulaCell;
      appliedCount += 1;
    }
  }

  formulaPlan.applied = appliedCount > 0;
  formulaPlan.formulaCount = appliedCount;
  formulaPlan.aggregateOperation = operation;
  formulaPlan.targets = targets.map((target) => ({
    operation: target.operation,
    columnIndex: target.columnIndex,
    header: headers[target.columnIndex],
  }));
  formulaPlan.reason =
    appliedCount > 0 ? "SIMPLE_AGGREGATE_FORMULA_APPLIED" : "NO_ELIGIBLE_ROWS";

  return appliedCount;
}

function createFormulaEngineMeta(mode = "static") {
  return {
    prepared: true,
    applied: false,
    mode,
    formulaCount: 0,
  };
}

function recordFormulaApplication(meta, count = 0) {
  if (!meta || !count) return;
  meta.formulaCount += count;
  meta.applied = meta.formulaCount > 0;
}

function attachWorkbookFormulaEngineMeta(wb, meta) {
  if (!wb || !meta) return wb;
  wb["!beebeeFormulaEngine"] = meta;
  return wb;
}

function buildFormulaEngineContext({
  sourceTables = [],
  summarySheetMode = "static",
  sourceSheetName = SHEET_NAMES.SOURCE_DATA,
  formulaOptions = {},
} = {}) {
  const mode = normalizeSummarySheetMode(summarySheetMode);
  const enabled = isFormulaEnabledMode(mode);
  const tables = Array.isArray(sourceTables)
    ? sourceTables.filter(Boolean)
    : [];
  const primaryTable =
    tables.find((table) => table?.isPrimary) || tables[0] || null;
  const primaryTableIndex = primaryTable
    ? Math.max(0, tables.indexOf(primaryTable))
    : -1;
  const resolvedSourceSheetName =
    tables.length > 1 && primaryTableIndex >= 0
      ? sourceSheetNameForTableIndex(primaryTableIndex, tables.length)
      : sourceSheetName;

  const columnMap = primaryTable ? createSourceColumnMap(primaryTable) : null;
  const rows = primaryTable ? getSourceTableRows(primaryTable) : [];

  return {
    enabled,
    mode,
    sourceSheetName: resolvedSourceSheetName,
    primaryTable,
    primaryTableIndex,
    columnMap,
    sourceRowCount: rows.length,
    formulaOptions,
    engine: {
      FORMULA_SPEC_TYPES,
      buildColumnRange,
      buildFormulaFromSpec,
      createFormulaCellFromSpec,
    },
  };
}

function buildSectionFormulaPlan({
  section = {},
  rows = [],
  formulaContext = null,
} = {}) {
  if (!formulaContext?.enabled) {
    return {
      enabled: false,
      reason: "FORMULA_MODE_DISABLED",
    };
  }

  if (!formulaContext.primaryTable || !formulaContext.columnMap) {
    return {
      enabled: false,
      reason: "SOURCE_TABLE_NOT_FOUND",
    };
  }

  const sectionResult = section.result || {};
  const groupHeader =
    sectionResult.groupBy?.header ||
    section.candidate?.columns?.dimension ||
    section.chartHint?.categoryField ||
    "";

  const metricHeader =
    sectionResult.metric?.header ||
    section.candidate?.columns?.metric ||
    section.chartHint?.valueField ||
    "";

  const groupColumn = resolveSourceColumn(
    formulaContext.columnMap,
    groupHeader,
    {
      preferNumeric: false,
    },
  );
  const metricColumn = resolveSourceColumn(
    formulaContext.columnMap,
    metricHeader,
    {
      preferNumeric: true,
    },
  );

  const operation = String(
    sectionResult.operation ||
      section.sectionType ||
      section.candidate?.recipeType ||
      "",
  );

  const aggregateOperation = normalizeAggregateOperation(operation, section);
  const rankingLike = isRankingLikeSection(section);
  const canApplySimpleAggregate =
    !rankingLike &&
    isSimpleAggregateOperation(aggregateOperation) &&
    Boolean(groupColumn) &&
    (aggregateOperation === "count" || Boolean(metricColumn));

  return {
    enabled: canApplySimpleAggregate,
    reason: canApplySimpleAggregate
      ? "SIMPLE_AGGREGATE_FORMULA_READY"
      : rankingLike
        ? "RANKING_SECTION_SKIPPED"
        : "MISSING_GROUP_OR_METRIC_COLUMN",
    mode: formulaContext.mode,
    sourceSheetName: formulaContext.sourceSheetName,
    sectionId: section.sectionId || "",
    sectionType: section.sectionType || "",
    operation,
    aggregateOperation,
    rowCount: Array.isArray(rows) ? rows.length : 0,
    group: groupColumn
      ? {
          header: groupColumn.header,
          letter: groupColumn.letter,
        }
      : null,
    metric: metricColumn
      ? {
          header: metricColumn.header,
          letter: metricColumn.letter,
        }
      : null,
  };
}

function attachFormulaPlan(ws, formulaPlan = null) {
  if (!ws || !formulaPlan) return ws;

  // xlsx 저장 결과에는 영향을 주지 않는 내부 메타.
  // 다음 단계에서 실제 formula cell 적용 시 이 구조를 사용한다.
  ws["!beebeeFormulaPlan"] = formulaPlan;
  return ws;
}

function getSourceTableRows(table = {}) {
  if (Array.isArray(table.rows)) return table.rows;
  if (Array.isArray(table.data)) return table.data;
  if (Array.isArray(table.records)) return table.records;
  if (Array.isArray(table.values)) return table.values;
  return [];
}

function getSourceColumnLabel(column = {}, index = 0) {
  return getSourceColumnHeader(column) || `컬럼_${index + 1}`;
}

function normalizeSourceDataHeader(header = "", index = 0) {
  const normalized = String(header || "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  return normalized || `컬럼_${index + 1}`;
}

function makeUniqueHeaders(headers = []) {
  const seen = new Map();

  return headers.map((header, index) => {
    const base = normalizeSourceDataHeader(header, index);
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);

    return count ? `${base}(${count + 1})` : base;
  });
}

function isUsableSourceColumn(column = {}) {
  const label = getSourceColumnLabel(column);
  return Boolean(String(label || "").trim());
}

function getOrderedSourceColumns(table = {}) {
  const columns = Array.isArray(table.columns) ? table.columns : [];
  return columns
    .filter(isUsableSourceColumn)
    .slice()
    .sort((a, b) => {
      const ai = Number.isFinite(Number(a.columnIndex))
        ? Number(a.columnIndex)
        : Number.MAX_SAFE_INTEGER;
      const bi = Number.isFinite(Number(b.columnIndex))
        ? Number(b.columnIndex)
        : Number.MAX_SAFE_INTEGER;

      if (ai !== bi) return ai - bi;
      return String(getSourceColumnLabel(a)).localeCompare(
        String(getSourceColumnLabel(b)),
        "ko",
      );
    });
}

function buildCleanSourceDataAoa(table = {}) {
  const rows = getSourceTableRows(table);
  const orderedColumns = getOrderedSourceColumns(table);

  if (orderedColumns.length) {
    const headers = makeUniqueHeaders(
      orderedColumns.map((column, index) =>
        getSourceColumnLabel(column, index),
      ),
    );
    const body = rows.map((row) =>
      orderedColumns.map((column, index) =>
        getSourceRowValue(row, column, index),
      ),
    );

    return { headers, orderedColumns, sourceRows: [headers, ...body] };
  }

  if (rows.length && typeof rows[0] === "object" && !Array.isArray(rows[0])) {
    const headers = makeUniqueHeaders([
      ...new Set(rows.flatMap((row) => Object.keys(row || {}))),
    ]);

    return {
      headers,
      orderedColumns: headers.map((header) => ({ header, key: header })),
      sourceRows: [
        headers,
        ...rows.map((row) => headers.map((header) => row?.[header] ?? "")),
      ],
    };
  }

  if (rows.length && Array.isArray(rows[0])) {
    const width = Math.max(...rows.map((row) => row.length), 1);
    const headers = makeUniqueHeaders(
      Array.from({ length: width }, (_, index) =>
        normalizeSourceDataHeader(rows[0]?.[index], index),
      ),
    );
    const body = rows
      .slice(1)
      .map((row) =>
        Array.from({ length: width }, (_, index) => row?.[index] ?? ""),
      );

    return {
      headers,
      orderedColumns: headers.map((header, index) => ({
        header,
        columnIndex: index + 1,
        columnLetter: XLSX.utils.encode_col(index),
      })),
      sourceRows: [headers, ...body],
    };
  }

  return {
    headers: ["데이터 없음"],
    orderedColumns: [],
    sourceRows: [["데이터 없음"]],
  };
}

function resolveColumnPosition(orderedColumns = [], targetColumn = null) {
  if (!targetColumn) return -1;

  const byReference = orderedColumns.indexOf(targetColumn);
  if (byReference >= 0) return byReference;

  const targetHeader = getSourceColumnLabel(targetColumn);
  const targetKey =
    targetColumn.key || targetColumn.accessor || targetColumn.name;
  const targetIndex = targetColumn.columnIndex;

  return orderedColumns.findIndex((column) => {
    if (targetIndex != null && column.columnIndex === targetIndex) return true;
    if (
      targetKey &&
      (column.key === targetKey || column.accessor === targetKey)
    ) {
      return true;
    }
    return targetHeader && getSourceColumnLabel(column) === targetHeader;
  });
}

function resolveSourceDataColumnLetter(
  orderedColumns = [],
  targetColumn = null,
  fallbackLetter = "A",
) {
  const position = resolveColumnPosition(orderedColumns, targetColumn);
  return position >= 0 ? XLSX.utils.encode_col(position) : fallbackLetter;
}

function flattenReasonList(value) {
  if (Array.isArray(value)) return value.filter(Boolean).join(", ");
  if (value == null) return "";
  return String(value);
}

function buildSummaryRowsAoa(table = {}) {
  const summaryRows = Array.isArray(table.summaryRows) ? table.summaryRows : [];
  const columns = getOrderedSourceColumns(table);
  const dataHeaders = makeUniqueHeaders(
    columns.map((column, index) => getSourceColumnLabel(column, index)),
  );

  const header = [
    "sourceTableId",
    "sourceSheetName",
    "sourceRow",
    "reason",
    "phase",
    "kind",
    ...dataHeaders,
  ];

  const body = summaryRows.map((summaryRow) => {
    const values = summaryRow.values || {};
    const rawCells = Array.isArray(summaryRow.rawCells)
      ? summaryRow.rawCells
      : [];

    return [
      table.tableId || "",
      table.sheetName || "",
      summaryRow.row ?? "",
      summaryRow.reason || "",
      summaryRow.phase || "",
      summaryRow.kind || "",
      ...columns.map((column, index) => {
        const value = getSourceRowValue(values, column, index);
        return value === "" && rawCells.length
          ? (rawCells[index] ?? "")
          : value;
      }),
    ];
  });

  return [header, ...body];
}

function buildDiagnosticsAoa({
  fileName = "",
  message = "",
  primaryTable = {},
  tables = [],
  sourceRows = [],
} = {}) {
  const dataRowCount = Math.max(0, sourceRows.length - 1);
  const primaryUsage = primaryTable.tableUsage || {};
  const primarySelection = primaryTable.primarySelection || {};
  const dataQuality = primaryTable.dataQuality || {};

  const overviewRows = [
    ["항목", "값", "비고"],
    ["원본 파일", fileName || "", "사용자가 업로드한 원본 파일명"],
    ["요청", message || "", "자동화 템플릿 생성 요청"],
    [
      "원본데이터 시트",
      SHEET_NAMES.SOURCE_DATA,
      "1행 헤더, 2행부터 정리된 분석 데이터",
    ],
    ["대표 테이블", primaryTable.tableId || "", "analysisEligible 기준 선택"],
    ["원본 범위", primaryTable.range || "", "업로드 파일 내 감지 범위"],
    ["데이터 범위", primaryTable.dataRange || "", "업로드 파일 내 데이터 범위"],
    ["정리 데이터 행 수", dataRowCount, "원본데이터 시트의 데이터 행 수"],
    [
      "summaryRows",
      dataQuality.summaryRowCount || 0,
      "요약행 시트로 분리된 행 수",
    ],
    [
      "excludedRows",
      dataQuality.excludedRowCount || 0,
      "분석에서 제외된 행 수",
    ],
    ["tableUsage.version", primaryUsage.version || "", "품질 필터 버전"],
    [
      "primarySelection.reason",
      primarySelection.reason || "",
      "대표 테이블 선택 사유",
    ],
    [],
    [
      "tableId",
      "sheetName",
      "isPrimary",
      "analysisEligible",
      "templateEligible",
      "rowCount",
      "rawDataRowCount",
      "summaryRowCount",
      "excludedRowCount",
      "range",
      "dataRange",
      "tableUsage.reasons",
      "primarySelection.reason",
    ],
  ];

  const tableRows = (Array.isArray(tables) ? tables : []).map((table) => {
    const usage = table.tableUsage || {};
    const selection = table.primarySelection || {};
    const quality = table.dataQuality || {};

    return [
      table.tableId || "",
      table.sheetName || "",
      table.isPrimary === true ? "TRUE" : "FALSE",
      usage.analysisEligible === true ? "TRUE" : "FALSE",
      usage.templateEligible === true ? "TRUE" : "FALSE",
      table.rowCount ?? "",
      table.rawDataRowCount ?? "",
      quality.summaryRowCount || 0,
      quality.excludedRowCount || 0,
      table.range || "",
      table.dataRange || "",
      flattenReasonList(usage.reasons),
      selection.reason || "",
    ];
  });

  return [...overviewRows, ...tableRows];
}

function getSourceRowValue(row = {}, column = {}, index = 0) {
  if (Array.isArray(row)) {
    return row[index] ?? "";
  }

  const keys = [
    column.header,
    column.originalHeader,
    column.name,
    column.key,
    column.accessor,
  ].filter(Boolean);

  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return row[key] ?? "";
    }
  }

  return Object.values(row || {})[index] ?? "";
}

function buildSourceTableAoa(table = {}) {
  const rows = getSourceTableRows(table);
  const columns = Array.isArray(table.columns) ? table.columns : [];

  if (columns.length) {
    const headers = columns.map(getSourceColumnLabel);
    const body = rows.map((row) =>
      columns.map((column, index) => getSourceRowValue(row, column, index)),
    );

    return [headers, ...body];
  }

  if (rows.length && typeof rows[0] === "object" && !Array.isArray(rows[0])) {
    const headers = [...new Set(rows.flatMap((row) => Object.keys(row || {})))];

    return [
      headers,
      ...rows.map((row) => headers.map((header) => row?.[header] ?? "")),
    ];
  }

  if (rows.length && Array.isArray(rows[0])) {
    return rows;
  }

  return [["데이터 없음"]];
}

function appendSourceDataSheets(
  wb,
  sourceTables = [],
  { includeSourceDataSheet = true } = {},
) {
  if (!includeSourceDataSheet) return;
  if (!Array.isArray(sourceTables) || !sourceTables.length) return;

  const tables = sourceTables.filter(Boolean);
  if (!tables.length) return;

  tables.forEach((table, index) => {
    const aoa = buildSourceTableAoa(table);
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    setAoaColumnWidths(ws, aoa);
    styleHeaderRow(ws);
    applyDefaultSheetOptions(ws);

    appendSheetSafe(wb, ws, sourceSheetNameForTableIndex(index, tables.length));
  });
}

function buildSummaryWorkbook({
  fileName,
  message,
  intent,
  result,
  sourceTables = [],
  summarySheetMode = "static",
  includeSourceDataSheet = true,
  formulaOptions = {},
}) {
  const normalizedSummarySheetMode =
    normalizeSummarySheetMode(summarySheetMode);
  const formulaContext = buildFormulaEngineContext({
    sourceTables,
    summarySheetMode: normalizedSummarySheetMode,
    formulaOptions,
  });
  const wb = XLSX.utils.book_new();
  const formulaEngineMeta = createFormulaEngineMeta(normalizedSummarySheetMode);

  const summaryRows = [
    ["요청", message || ""],
    ["원본 파일", fileName || ""],
    ["테이블", result?.table?.tableName || intent?.table?.tableName || ""],
    ["결과 유형", result?.resultType || ""],
    ["작업", result?.operation || intent?.operation || ""],
    ["지표", result?.metric?.header || intent?.metric?.header || ""],
    ["그룹 기준", result?.groupBy?.header || intent?.groupBy?.header || ""],
    ["결과 행 수", Array.isArray(result?.rows) ? result.rows.length : 0],
    ["생성일시", formatKstTimestamp()],
    [],
  ];

  const businessSections = Array.isArray(result?.sections)
    ? result.sections
    : [];

  if (businessSections.length) {
    const wb = XLSX.utils.book_new();

    const summaryRows = buildBusinessWorkbookSummaryRows({
      fileName,
      message,
      result,
      businessSections,
      summarySheetMode: normalizedSummarySheetMode,
      includeSourceDataSheet,
      sourceTables,
    });

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    setAoaColumnWidths(wsSummary, summaryRows);
    styleHeaderRow(wsSummary);
    applyDefaultSheetOptions(wsSummary);
    appendSheetSafe(wb, wsSummary, SHEET_NAMES.SUMMARY);

    appendSourceDataSheets(wb, sourceTables, {
      includeSourceDataSheet,
      summarySheetMode: normalizedSummarySheetMode,
    });

    businessSections.forEach((section, index) => {
      const sectionResult = section.result || {};
      const rows = resultToRows(sectionResult);
      const outputRows = rows.length
        ? rows
        : [
            {
              섹션: section.title || section.sectionId || `섹션_${index + 1}`,
              결과: "데이터 없음",
            },
          ];
      const ws = XLSX.utils.json_to_sheet(outputRows);
      const formulaPlan = buildSectionFormulaPlan({
        section,
        rows,
        formulaContext,
      });

      setColumnWidths(ws, outputRows);
      formatNumberCells(ws, sectionResult);
      styleHeaderRow(ws);
      applyDefaultSheetOptions(ws);
      ws["!freeze"] = { xSplit: 0, ySplit: 1 };
      const formulaCount = applySimpleAggregateFormulaSection({
        ws,
        rows,
        formulaPlan,
        formulaContext,
      });
      recordFormulaApplication(formulaEngineMeta, formulaCount);
      attachFormulaPlan(ws, formulaPlan);

      appendSheetSafe(
        wb,
        ws,
        section.title || section.sectionId || `섹션_${index + 1}`,
      );
    });

    attachWorkbookFormulaEngineMeta(wb, formulaEngineMeta);
    return wb;
  }

  const rows = resultToRows(result);

  const narrative = buildNarrativeSections(result, {
    message,
    fileName,
  });

  const reportSections = buildReportSections({
    fileName,
    message,
    result,
  });

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  setAoaColumnWidths(wsSummary, summaryRows);
  styleHeaderRow(wsSummary);
  applyDefaultSheetOptions(wsSummary);
  appendSheetSafe(wb, wsSummary, SHEET_NAMES.SUMMARY);

  appendSourceDataSheets(wb, sourceTables, {
    includeSourceDataSheet,
    summarySheetMode: normalizedSummarySheetMode,
  });

  const wsResult = XLSX.utils.json_to_sheet(rows);
  const resultFormulaPlan = buildSectionFormulaPlan({
    section: {
      sectionId: "analysis_result",
      sectionType: result?.operation || result?.resultType || "analysis",
      result,
    },
    rows,
    formulaContext,
  });

  setColumnWidths(wsResult, rows);
  formatNumberCells(wsResult, result);
  styleHeaderRow(wsResult);
  applyDefaultSheetOptions(wsResult);
  wsResult["!freeze"] = { xSplit: 0, ySplit: 1 };
  const resultFormulaCount = applySimpleAggregateFormulaSection({
    ws: wsResult,
    rows,
    formulaPlan: resultFormulaPlan,
    formulaContext,
  });
  recordFormulaApplication(formulaEngineMeta, resultFormulaCount);
  attachFormulaPlan(wsResult, resultFormulaPlan);
  appendSheetSafe(wb, wsResult, SHEET_NAMES.ANALYSIS_RESULT);

  if (result?.plan) {
    const planRows = objectToAoa(result.plan);
    const wsPlan = XLSX.utils.aoa_to_sheet([["항목", "값"], ...planRows]);
    setAoaColumnWidths(wsPlan, [["항목", "값"], ...planRows]);
    styleHeaderRow(wsPlan);
    applyDefaultSheetOptions(wsPlan);
    XLSX.utils.book_append_sheet(wb, wsPlan, SHEET_NAMES.EXECUTION_PLAN);
  }

  if (result?.executionMeta) {
    const metaRows = objectToAoa(result.executionMeta);
    const wsMeta = XLSX.utils.aoa_to_sheet([["항목", "값"], ...metaRows]);
    setAoaColumnWidths(wsMeta, [["항목", "값"], ...metaRows]);
    styleHeaderRow(wsMeta);
    applyDefaultSheetOptions(wsMeta);
    XLSX.utils.book_append_sheet(wb, wsMeta, SHEET_NAMES.EXECUTION_META);
  }

  const chartDataRows = buildChartDataRows(result);
  if (chartDataRows.length) {
    const wsChartData = XLSX.utils.json_to_sheet(chartDataRows);
    setColumnWidths(wsChartData, chartDataRows);
    formatNumberCells(wsChartData, result);
    styleHeaderRow(wsChartData);
    applyDefaultSheetOptions(wsChartData);
    wsChartData["!freeze"] = { xSplit: 0, ySplit: 1 };
    appendSheetSafe(wb, wsChartData, SHEET_NAMES.CHART_DATA);
  }

  const chartSpec = buildChartSpec(result);

  if (chartSpec) {
    const chartSpecRows = objectToAoa(chartSpec);

    const wsChartSpec = XLSX.utils.aoa_to_sheet([
      ["항목", "값"],
      ...chartSpecRows,
    ]);

    setAoaColumnWidths(wsChartSpec, [["항목", "값"], ...chartSpecRows]);
    styleHeaderRow(wsChartSpec);
    applyDefaultSheetOptions(wsChartSpec);
    appendSheetSafe(wb, wsChartSpec, SHEET_NAMES.CHART_CONFIG);
  }

  const narrativeRows = [
    ["제목", narrative.title],
    ["요약", narrative.summary],
    ...(narrative.highlights || []).map((v, idx) => [`핵심 ${idx + 1}`, v]),
    [],
  ];

  const insightRows = [...narrativeRows, ...buildInsightRows(result)];

  if (insightRows.length) {
    const wsInsight = XLSX.utils.aoa_to_sheet([
      ["항목", "내용"],
      ...insightRows,
    ]);
    setAoaColumnWidths(wsInsight, [["항목", "내용"], ...insightRows]);
    styleHeaderRow(wsInsight);
    applyDefaultSheetOptions(wsInsight);
    appendSheetSafe(wb, wsInsight, SHEET_NAMES.INSIGHTS);
  }

  const reportSectionRows = objectToAoa(reportSections);

  if (reportSectionRows.length) {
    const wsReportSections = XLSX.utils.aoa_to_sheet([
      ["항목", "값"],
      ...reportSectionRows,
    ]);

    setAoaColumnWidths(wsReportSections, [
      ["항목", "값"],
      ...reportSectionRows,
    ]);
    styleHeaderRow(wsReportSections);
    applyDefaultSheetOptions(wsReportSections);
    appendSheetSafe(wb, wsReportSections, SHEET_NAMES.REPORT_SECTIONS);
  }

  attachWorkbookFormulaEngineMeta(wb, formulaEngineMeta);
  return wb;
}

function workbookToBuffer(workbook) {
  return XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
  });
}

function getRowValueByColumn(row = {}, col = {}) {
  const candidates = [
    col.header,
    col.key,
    col.originalHeader,
    col.name,
    col.accessor,
  ].filter(Boolean);

  for (const key of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return row[key] ?? "";
    }
  }

  return "";
}

function getFirstResultDimensionHeader(result = {}) {
  const firstRow = Array.isArray(result?.rows) ? result.rows[0] : null;
  if (!firstRow) return "";

  const excluded = new Set([
    "operation",
    "metric",
    "value",
    "rowCount",
    "기준값",
    "비교값",
    "증감률",
  ]);

  return (
    Object.keys(firstRow).find((key) => !excluded.has(key)) ||
    result?.groupBy?.header ||
    ""
  );
}

function findSourceDateColumn(columns = []) {
  return (
    columns.find((c) => c.role === "date" || c.inferredRole === "date") ||
    columns.find((c) => c.type === "date" || c.dominantType === "date") ||
    columns.find((c) => String(c.header || "").includes("일")) ||
    null
  );
}

function ensureDerivedGroupColumn({
  headers,
  orderedColumns,
  sourceRows,
  columns,
  groupHeader,
}) {
  if (!groupHeader) {
    return { headers, orderedColumns, sourceRows, groupLetter: null };
  }

  const existingIndex = headers.findIndex((h) => h === groupHeader);
  if (existingIndex >= 0) {
    return {
      headers,
      orderedColumns,
      sourceRows,
      groupLetter: XLSX.utils.encode_col(existingIndex),
    };
  }

  const dateCol = findSourceDateColumn(columns);
  if (!dateCol) {
    return { headers, orderedColumns, sourceRows, groupLetter: null };
  }

  const dateColIndex = resolveColumnPosition(orderedColumns, dateCol);
  if (dateColIndex < 0) {
    return { headers, orderedColumns, sourceRows, groupLetter: null };
  }

  const derivedIndex = headers.length;
  const nextHeaders = [...headers, groupHeader];

  const nextRows = sourceRows.map((row, rowIndex) => {
    if (rowIndex === 0) return nextHeaders;

    const dateValue = row[dateColIndex];
    const date = new Date(dateValue);

    let derivedValue = "";
    if (!Number.isNaN(date.getTime())) {
      if (String(groupHeader).includes("연월")) {
        derivedValue = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
      } else if (String(groupHeader).includes("연도")) {
        derivedValue = date.getFullYear();
      }
    }

    return [...row, derivedValue];
  });

  return {
    headers: nextHeaders,
    orderedColumns: [...orderedColumns, null],
    sourceRows: nextRows,
    groupLetter: XLSX.utils.encode_col(derivedIndex),
  };
}

function isAnalysisEligibleSourceTable(table = {}) {
  return table?.tableUsage?.analysisEligible === true;
}

function getAutomationSourceTables(tables = []) {
  const list = Array.isArray(tables) ? tables.filter(Boolean) : [];
  const eligibleTables = list.filter(isAnalysisEligibleSourceTable);

  if (eligibleTables.length) return eligibleTables;

  return list.filter((table) => table?.tableUsage?.analysisEligible !== false);
}

function pickAutomationPrimaryTable(sourceTables = [], allTables = []) {
  return (
    sourceTables.find((t) => t?.isPrimary === true) ||
    sourceTables.find(isAnalysisEligibleSourceTable) ||
    allTables.find(
      (t) => t?.isPrimary && t?.tableUsage?.analysisEligible !== false,
    ) ||
    allTables.find(isAnalysisEligibleSourceTable) ||
    sourceTables[0] ||
    allTables[0] ||
    {}
  );
}

function getPrimarySourceSheetName(primaryTable = {}, sourceTables = []) {
  const primaryIndex = Math.max(0, sourceTables.indexOf(primaryTable));
  return sourceSheetNameForTableIndex(primaryIndex, sourceTables.length || 1);
}

function buildAutomationTemplateWorkbook({
  fileName = "",
  message = "",
  intent = null,
  result = null,
  tables = [],
}) {
  const wb = XLSX.utils.book_new();

  const allTables = Array.isArray(tables) ? tables.filter(Boolean) : [];
  const sourceTables = getAutomationSourceTables(allTables);
  const primaryTable = pickAutomationPrimaryTable(sourceTables, allTables);
  const primarySourceSheetName = getPrimarySourceSheetName(
    primaryTable,
    sourceTables.length ? sourceTables : [primaryTable],
  );
  const columns = primaryTable.columns || [];
  const cleanSourceData = buildCleanSourceDataAoa(primaryTable);

  let headers = cleanSourceData.headers;
  let orderedColumns = cleanSourceData.orderedColumns;
  let sourceRows = cleanSourceData.sourceRows;

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ["자동화 템플릿 사용방법"],
      ["1", "원본데이터 시트에 실제 데이터를 붙여넣습니다."],
      ["2", "자동화설정 시트에서 기준열/값열/집계방식을 수정합니다."],
      ["3", "자동화시트 시트의 수식 결과를 확인합니다."],
      ["4", "필요 시 자동화시트를 복사해 실제 업무 파일에 붙여넣습니다."],
      [],
      ["파일명", fileName],
      ["요청", message],
      ["작업", result?.operation || intent?.operation || ""],
    ]),
    SHEET_NAMES.AUTOMATION_GUIDE,
  );

  const resultDimensionHeader = getFirstResultDimensionHeader(result);

  const groupHeader =
    result?.groupBy?.header ||
    intent?.groupBy?.header ||
    resultDimensionHeader ||
    "";

  const metricHeader = result?.metric?.header || intent?.metric?.header || "";

  const derivedGroup = ensureDerivedGroupColumn({
    headers,
    orderedColumns,
    sourceRows,
    columns,
    groupHeader,
  });

  headers = derivedGroup.headers;
  orderedColumns = derivedGroup.orderedColumns;
  sourceRows = derivedGroup.sourceRows;

  const derivedGroupLetter = derivedGroup.groupLetter || null;

  const groupCol =
    columns.find((c) => c.header === groupHeader) ||
    columns.find((c) => c.role === "group" || c.inferredRole === "group") ||
    columns[0] ||
    {};

  const metricCol =
    columns.find((c) => c.header === metricHeader) ||
    columns.find((c) => c.role === "metric" || c.inferredRole === "metric") ||
    columns.find((c) => c.type === "number" || c.dominantType === "number") ||
    columns[1] ||
    {};

  const groupLetter =
    derivedGroupLetter ||
    resolveSourceDataColumnLetter(orderedColumns, groupCol, "A");
  const metricLetter = resolveSourceDataColumnLetter(
    orderedColumns,
    metricCol,
    "B",
  );
  const rawOperation = result?.operation || intent?.operation || "average";

  console.log("[automation-template]", {
    groupHeader,
    derivedGroupLetter,
    groupLetter,
    metricHeader,
    metricLetter,
    rawOperation,
  });

  const operation =
    rawOperation === "cumulativeSum" ||
    rawOperation === "rollingAverage" ||
    rawOperation === "growthRate"
      ? "sum"
      : rawOperation;

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ["항목", "값", "설명"],
      ["원본시트명", primarySourceSheetName, "데이터가 들어있는 시트명"],
      ["기준열", groupLetter, "부서/월/분류 등 그룹 기준 열"],
      ["값열", metricLetter, "합계/평균 계산 대상 열"],
      ["집계방식", operation, "average, sum, count 중 선택"],
      ["요청문", message, "자동 생성 기준 요청"],
    ]),
    SHEET_NAMES.AUTOMATION_SETTINGS,
  );

  const resultRows = Array.isArray(result?.rows) ? result.rows : [];
  const autoGroupHeader = groupHeader || "기준";

  const uniqueValues = resultRows
    .map((row) => row[autoGroupHeader])
    .filter((v) => v !== undefined && v !== null && v !== "");

  const labelRange = buildColumnRange(primarySourceSheetName, groupLetter);
  const valueRange = buildColumnRange(primarySourceSheetName, metricLetter);

  const isListTemplate = rawOperation === "list";
  const isCumulativeTemplate = rawOperation === "cumulativeSum";
  const isRollingTemplate = rawOperation === "rollingAverage";
  const isGrowthTemplate = rawOperation === "growthRate";
  const isPivotTemplate = rawOperation === "pivot";
  const isMultiAggregateTemplate = rawOperation === "multiAggregate";
  const isPipelineCombineTemplate = rawOperation === "pipelineCombine";

  let autoRows = [
    ["자동화시트"],
    ["설정값을 바꾸면 아래 결과가 자동 계산됩니다."],
    [],
  ];

  if (isPivotTemplate) {
    const pivotColumns = result?.pivot?.columns || [];
    autoRows.push(["기준", ...pivotColumns]);

    uniqueValues.forEach((value) => {
      autoRows.push([value, ...pivotColumns.map(() => null)]);
    });
  } else if (isMultiAggregateTemplate) {
    autoRows.push(["기준", "평균", "최대값", "건수"]);
    uniqueValues.forEach((value) => autoRows.push([value, null, null, null]));
  } else if (isPipelineCombineTemplate) {
    autoRows.push(["상태", "설명"]);
    autoRows.push([
      "미지원",
      "pipelineCombine은 다중 기준/다중 지표 템플릿 분기가 필요합니다.",
    ]);
  } else if (isListTemplate) {
    autoRows.push(["순위", "항목", "값"]);
    for (let i = 1; i <= Math.min(10, resultRows.length || 10); i += 1) {
      autoRows.push([i, null, null]);
    }
  } else if (isCumulativeTemplate) {
    autoRows.push(["기준", "값", "누적값"]);
    uniqueValues.forEach((value) => autoRows.push([value, null, null]));
  } else if (isRollingTemplate) {
    autoRows.push(["기준", "값", "이동평균"]);
    uniqueValues.forEach((value) => autoRows.push([value, null, null]));
  } else if (isGrowthTemplate) {
    autoRows.push(["기준", "값", "증감률"]);
    uniqueValues.forEach((value) => autoRows.push([value, null, null]));
  } else {
    autoRows.push(["기준", "값"]);
    uniqueValues.forEach((value) => autoRows.push([value, null]));
  }

  const autoSheet = XLSX.utils.aoa_to_sheet(autoRows);

  if (isPivotTemplate) {
    const pivotColumns = result?.pivot?.columns || [];
    const rowGroupHeader =
      result?.pivot?.rowGroup?.header || result?.groupBy?.header || groupHeader;
    const colGroupHeader = result?.pivot?.columnGroup?.header || "";

    const rowCol = columns.find((c) => c.header === rowGroupHeader) || groupCol;
    const colCol = columns.find((c) => c.header === colGroupHeader) || {};

    const rowLetter =
      derivedGroupLetter ||
      resolveSourceDataColumnLetter(orderedColumns, rowCol, groupLetter);
    const colLetter = resolveSourceDataColumnLetter(
      orderedColumns,
      colCol,
      groupLetter,
    );

    const rowRange = buildColumnRange(primarySourceSheetName, rowLetter);
    const colRange = buildColumnRange(primarySourceSheetName, colLetter);

    for (let r = 0; r < uniqueValues.length; r += 1) {
      const rowNum = r + 5;

      for (let c = 0; c < pivotColumns.length; c += 1) {
        const colNum = c + 2;
        const cellAddr = XLSX.utils.encode_cell({
          r: rowNum - 1,
          c: colNum - 1,
        });
        const headerAddr = XLSX.utils.encode_cell({
          r: 3,
          c: colNum - 1,
        });

        autoSheet[cellAddr] = {
          t: "n",
          f: buildPivotAverageFormula({
            rowRange,
            rowCriteriaCell: `A${rowNum}`,
            colRange,
            colCriteriaCell: headerAddr,
            valueRange,
          }),
          v: 0,
        };
      }
    }
  } else if (isMultiAggregateTemplate) {
    for (let i = 0; i < uniqueValues.length; i += 1) {
      const rowNum = i + 5;

      autoSheet[`B${rowNum}`] = {
        t: "n",
        f: buildGroupAggregateFormula({
          operation: "average",
          sheetName: primarySourceSheetName,
          groupLetter,
          metricLetter,
          criteriaCell: `A${rowNum}`,
        }),
        v: 0,
      };

      autoSheet[`C${rowNum}`] = {
        t: "n",
        f: buildMaxIfFormula({
          groupRange: labelRange,
          criteriaCell: `A${rowNum}`,
          valueRange,
        }),
        v: 0,
      };

      autoSheet[`D${rowNum}`] = {
        t: "n",
        f: buildCountIfsFormula({
          criteriaRange: labelRange,
          criteriaCell: `A${rowNum}`,
        }),
        v: 0,
      };
    }
  } else if (isPipelineCombineTemplate) {
    // 현재 베타에서는 안내형 템플릿만 생성
  } else if (isListTemplate) {
    for (let i = 0; i < Math.min(10, resultRows.length || 10); i += 1) {
      const rowNum = i + 5;
      const rank = i + 1;

      autoSheet[`C${rowNum}`] = {
        t: "n",
        f: buildRankValueFormula({ valueRange, rank }),
        v: 0,
      };

      autoSheet[`B${rowNum}`] = {
        t: "s",
        f: buildRankLabelFormula({
          labelRange,
          valueRange,
          rankValueCell: `C${rowNum}`,
        }),
        v: "",
      };
    }
  } else {
    for (let i = 0; i < uniqueValues.length; i += 1) {
      const rowNum = i + 5;

      const baseFormula = buildGroupAggregateFormula({
        operation,
        sheetName: primarySourceSheetName,
        groupLetter,
        metricLetter,
        criteriaCell: `A${rowNum}`,
      });

      autoSheet[`B${rowNum}`] = {
        t: "n",
        f: baseFormula,
        v: 0,
      };

      if (isCumulativeTemplate) {
        autoSheet[`C${rowNum}`] = {
          t: "n",
          f:
            rowNum === 5
              ? `IFERROR(B${rowNum},"")`
              : buildRunningSumFormula({
                  valueCell: `B${rowNum}`,
                  previousCell: `C${rowNum - 1}`,
                }),
          v: 0,
        };
      }

      if (isRollingTemplate) {
        const startRow = Math.max(5, rowNum - 2);
        autoSheet[`C${rowNum}`] = {
          t: "n",
          f: `IFERROR(AVERAGE(B${startRow}:B${rowNum}),"")`,
          v: 0,
        };
      }

      if (isGrowthTemplate) {
        autoSheet[`C${rowNum}`] = {
          t: "n",
          f:
            rowNum === 5
              ? `""`
              : buildGrowthRateFormula({
                  currentCell: `B${rowNum}`,
                  previousCell: `B${rowNum - 1}`,
                }),
          v: 0,
        };
      }
    }
  }

  autoSheet["!ref"] = `A1:C${Math.max(5, autoRows.length)}`;

  XLSX.utils.book_append_sheet(wb, autoSheet, SHEET_NAMES.AUTOMATION_TEMPLATE);

  const sheetsToWrite = sourceTables.length ? sourceTables : [primaryTable];
  sheetsToWrite.forEach((table, index) => {
    const sourceData =
      table === primaryTable ? cleanSourceData : buildCleanSourceDataAoa(table);
    const tableSourceRows = sourceData.sourceRows?.length
      ? sourceData.sourceRows
      : [["데이터 없음"]];
    const sourceDataSheet = XLSX.utils.aoa_to_sheet(tableSourceRows);
    setAoaColumnWidths(sourceDataSheet, tableSourceRows);
    styleHeaderRow(sourceDataSheet);
    applyDefaultSheetOptions(sourceDataSheet);
    appendSheetSafe(
      wb,
      sourceDataSheet,
      sourceSheetNameForTableIndex(index, sheetsToWrite.length),
    );
  });

  const summaryRowsAoa = buildSummaryRowsAoa(primaryTable);
  const summaryRowsSheet = XLSX.utils.aoa_to_sheet(summaryRowsAoa);
  setAoaColumnWidths(summaryRowsSheet, summaryRowsAoa);
  styleHeaderRow(summaryRowsSheet);
  applyDefaultSheetOptions(summaryRowsSheet);
  XLSX.utils.book_append_sheet(wb, summaryRowsSheet, SHEET_NAMES.SUMMARY_ROWS);

  const diagnosticsAoa = buildDiagnosticsAoa({
    fileName,
    message,
    primaryTable,
    tables,
    sourceRows,
  });
  const diagnosticsSheet = XLSX.utils.aoa_to_sheet(diagnosticsAoa);
  setAoaColumnWidths(diagnosticsSheet, diagnosticsAoa);
  styleHeaderRow(diagnosticsSheet);
  applyDefaultSheetOptions(diagnosticsSheet);
  XLSX.utils.book_append_sheet(wb, diagnosticsSheet, SHEET_NAMES.DIAGNOSTICS);

  const previewSheet = XLSX.utils.json_to_sheet(result?.rows || []);
  applyDefaultSheetOptions(previewSheet);
  XLSX.utils.book_append_sheet(wb, previewSheet, SHEET_NAMES.EXECUTION_PREVIEW);

  return wb;
}

module.exports = {
  buildSummaryWorkbook,
  buildAutomationTemplateWorkbook,
  workbookToBuffer,
  buildChartSpec,
  normalizeSummarySheetMode,
  buildFormulaEngineContext,
  buildSectionFormulaPlan,
  applySimpleAggregateFormulaSection,
};
