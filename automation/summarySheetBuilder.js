const XLSX = require("xlsx");
const zlib = require("zlib");
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
  buildSourceTablePolicy,
  summarizeSourceTablePolicy,
} = require("./sourceTablePolicy");
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
const {
  SUMMARY_SHEET_RECIPE_MANIFEST_VERSION,
  appendSummarySheetRecipeManifest,
} = require("./summarySheetRecipeManifest");

const AUTOMATION_SHEET_V2_VERSION = "automation_sheet_v2";
const WORKBOOK_RECALCULATION_FLAG_VERSION = "workbook_recalculation_flag_v1";
const WORKBOOK_XML_CALCPR_INJECTION_VERSION =
  "workbook_xml_calcpr_injection_v1";
const WORKBOOK_CALCPR_XML =
  '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';

function emitSummarySheetDiagnostic(
  diagnostic,
  stage,
  status = "INFO",
  meta = {},
) {
  try {
    if (diagnostic && typeof diagnostic.checkpoint === "function") {
      diagnostic.checkpoint(stage, status, meta);
    }
  } catch (error) {
    // Diagnostics must never break workbook generation.
  }
}

function summarizeSectionForWorkbookDiagnostic(section = {}, index = 0) {
  const sectionResult = section?.result || {};
  return {
    index,
    sectionId: section?.sectionId || "",
    title: section?.title || "",
    resultType: sectionResult?.resultType || "",
    operation: sectionResult?.operation || "",
    rowCount: Array.isArray(sectionResult?.rows)
      ? sectionResult.rows.length
      : 0,
    hasMetric: Boolean(sectionResult?.metric?.header),
    hasGroupBy: Boolean(sectionResult?.groupBy?.header),
  };
}

function isTopBottomLikeResult(result = {}) {
  const recipeType = String(
    result.recipeType || result.recipeId || "",
  ).toLowerCase();
  const operation = String(result.operation || "").toLowerCase();

  return (
    recipeType === "top_bottom" ||
    operation === "topbottom" ||
    operation === "top_bottom"
  );
}

function resolveGroupedResultLabel(result = {}, row = {}, groupHeader = "") {
  return row?.[groupHeader] ?? row?.label ?? row?.item ?? row?.name ?? "";
}

function resolveGroupedResultOperation(result = {}, row = {}) {
  return row?.operation ?? row?.type ?? result.operation ?? "";
}

function resolveGroupedResultMetric(result = {}, row = {}, metricHeader = "") {
  return (
    row?.metric ??
    result.metric?.displayHeader ??
    result.metric?.header ??
    metricHeader ??
    ""
  );
}

function resolveGroupedResultRowCount(result = {}, row = {}) {
  const direct = row?.rowCount ?? row?.count;
  if (direct != null && direct !== "") return direct;
  return isTopBottomLikeResult(result) ? 1 : "";
}

function runWorkbookDiagnosticStep(diagnostic, stage, fn, meta = {}) {
  emitSummarySheetDiagnostic(diagnostic, stage, "START", meta);
  const startedAt = Date.now();
  try {
    const value = fn();
    emitSummarySheetDiagnostic(diagnostic, stage, "OK", {
      stageElapsedMs: Date.now() - startedAt,
    });
    return value;
  } catch (error) {
    emitSummarySheetDiagnostic(diagnostic, stage, "ERROR", {
      stageElapsedMs: Date.now() - startedAt,
      error: error?.message || String(error),
      stack: error?.stack || "",
    });
    throw error;
  }
}

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
      [groupHeader]: resolveGroupedResultLabel(result, r, groupHeader),
      [metricHeader]: r.value,
      행수: resolveGroupedResultRowCount(result, r),
    }));
  }

  if (result.resultType === "pivot") {
    return result.rows || [];
  }

  return [];
}

const XLSX_MAX_SHEET_NAME_LENGTH = 31;
const SHEET_NAME_COLLISION_MAX_ATTEMPTS = 200;

function sanitizeSheetNameBase(name) {
  return (
    String(name || "Sheet")
      .replace(/[:\\/?*\[\]]/g, " ")
      .replace(/\s+/g, " ")
      .trim() || "Sheet"
  );
}

function truncateSheetNameBase(base, maxLength = XLSX_MAX_SHEET_NAME_LENGTH) {
  const safeMax = Math.max(1, Number(maxLength) || XLSX_MAX_SHEET_NAME_LENGTH);
  return (
    String(base || "Sheet")
      .slice(0, safeMax)
      .trim() || "Sheet"
  );
}

function sanitizeSheetName(name) {
  return truncateSheetNameBase(sanitizeSheetNameBase(name));
}

function normalizeSheetNameForCollision(name = "") {
  return String(name || "")
    .trim()
    .toLocaleLowerCase("ko-KR");
}

function buildSheetNameCollisionHash(value = "") {
  const text = String(value || "Sheet");
  let hash = 2166136261;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(36).slice(0, 6);
}

function resolveUniqueSheetName(wb, requestedName = "Sheet") {
  const existingNames = Array.isArray(wb?.SheetNames) ? wb.SheetNames : [];
  const existingNormalized = new Set(
    existingNames.map((sheetName) => normalizeSheetNameForCollision(sheetName)),
  );
  const requestedBase = sanitizeSheetNameBase(requestedName);
  const initialName = truncateSheetNameBase(requestedBase);
  const initialKey = normalizeSheetNameForCollision(initialName);

  if (!existingNormalized.has(initialKey)) {
    return {
      requestedSheetName: String(requestedName || ""),
      sanitizedBaseName: requestedBase,
      safeSheetName: initialName,
      collisionCount: 0,
      collisionGuardApplied: false,
      truncated: requestedBase !== initialName,
      existingSheetCount: existingNames.length,
    };
  }

  for (
    let attempt = 2;
    attempt <= SHEET_NAME_COLLISION_MAX_ATTEMPTS;
    attempt += 1
  ) {
    const suffix = `_${attempt}`;
    const baseLength = XLSX_MAX_SHEET_NAME_LENGTH - suffix.length;
    const candidate = `${truncateSheetNameBase(requestedBase, baseLength)}${suffix}`;
    const candidateKey = normalizeSheetNameForCollision(candidate);

    if (!existingNormalized.has(candidateKey)) {
      return {
        requestedSheetName: String(requestedName || ""),
        sanitizedBaseName: requestedBase,
        safeSheetName: candidate,
        collisionCount: attempt - 1,
        collisionGuardApplied: true,
        truncated: requestedBase !== candidate,
        existingSheetCount: existingNames.length,
      };
    }
  }

  const hashSuffix = `_${buildSheetNameCollisionHash(
    `${requestedBase}|${existingNames.join("|")}`,
  )}`;
  const hashedName = `${truncateSheetNameBase(
    requestedBase,
    XLSX_MAX_SHEET_NAME_LENGTH - hashSuffix.length,
  )}${hashSuffix}`;
  let candidate = hashedName;
  let guardIndex = 2;

  while (
    existingNormalized.has(normalizeSheetNameForCollision(candidate)) &&
    guardIndex <= 50
  ) {
    const suffix = `_${buildSheetNameCollisionHash(`${hashedName}|${guardIndex}`)}`;
    candidate = `${truncateSheetNameBase(
      requestedBase,
      XLSX_MAX_SHEET_NAME_LENGTH - suffix.length,
    )}${suffix}`;
    guardIndex += 1;
  }

  if (existingNormalized.has(normalizeSheetNameForCollision(candidate))) {
    throw new Error(
      `Unable to resolve unique worksheet name after ${SHEET_NAME_COLLISION_MAX_ATTEMPTS} attempts: ${requestedBase}`,
    );
  }

  return {
    requestedSheetName: String(requestedName || ""),
    sanitizedBaseName: requestedBase,
    safeSheetName: candidate,
    collisionCount: SHEET_NAME_COLLISION_MAX_ATTEMPTS,
    collisionGuardApplied: true,
    hashFallbackApplied: true,
    truncated: requestedBase !== candidate,
    existingSheetCount: existingNames.length,
  };
}

function appendSheetSafe(wb, ws, name, options = {}) {
  const resolved = resolveUniqueSheetName(wb, name);

  if (options && typeof options.onResolved === "function") {
    try {
      options.onResolved(resolved);
    } catch (error) {
      // Worksheet diagnostics must never break workbook generation.
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, resolved.safeSheetName);
  return resolved;
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
        label: resolveGroupedResultLabel(result, r, groupHeader),
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

function hasBusinessTemplateRatioScale(result = {}) {
  const meta = result?.meta || {};
  const recipeType = String(
    result?.recipeType || result?.candidate?.recipeType || "",
  );
  const explicitScale = String(
    meta.rateValueScale ||
      meta.percentageValueScale ||
      result.rateValueScale ||
      "",
  ).toLowerCase();

  if (explicitScale === "ratio" || explicitScale === "fraction") return true;
  if (meta.salesReportVersion || meta.researchBudgetReportVersion) return true;
  if (
    /sales_report_v2_custom|research_budget_report_v2_custom/.test(recipeType)
  )
    return true;

  return false;
}

function inferNumberFormat(header = "", result = {}) {
  const h = String(header || "");

  if (
    h.includes("%") ||
    h.includes("률") ||
    h.includes("비율") ||
    (result.metric?.type === "rate" && h === "값")
  ) {
    const alreadyRatio = hasBusinessTemplateRatioScale(result);
    return {
      z: "0.00%",
      divideBy100: !alreadyRatio,
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
        [groupHeader]: resolveGroupedResultLabel(result, r, groupHeader),
        작업: resolveGroupedResultOperation(result, r),
        지표: resolveGroupedResultMetric(
          result,
          r,
          result.metric?.header || "값",
        ),
        값: r.value,
        행수: resolveGroupedResultRowCount(result, r),
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

function ensureWorkbookRecalculationOnOpen(
  wb,
  reason = "formula_or_dynamic_summary",
) {
  if (!wb || typeof wb !== "object") return wb;

  if (!wb.Workbook || typeof wb.Workbook !== "object") {
    wb.Workbook = {};
  }

  const calcPr = {
    ...(wb.CalcPr || {}),
    ...(wb.Workbook.CalcPr || {}),
    calcId: wb.Workbook.CalcPr?.calcId || wb.CalcPr?.calcId || 171027,
    calcMode: "auto",
    fullCalcOnLoad: "1",
    forceFullCalc: "1",
  };

  wb.CalcPr = calcPr;
  wb.Workbook.CalcPr = calcPr;
  wb.Workbook.WBProps = {
    ...(wb.Workbook.WBProps || {}),
  };

  wb["!beebeeWorkbookRecalculation"] = {
    version: WORKBOOK_RECALCULATION_FLAG_VERSION,
    applied: true,
    calcMode: "auto",
    fullCalcOnLoad: true,
    forceFullCalc: true,
    reason,
  };

  return wb;
}

function makeCrc32Table() {
  const table = new Array(256);
  for (let i = 0; i < 256; i += 1) {
    let crc = i;
    for (let j = 0; j < 8; j += 1) {
      crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    table[i] = crc >>> 0;
  }
  return table;
}

const CRC32_TABLE = makeCrc32Table();

function crc32(buffer) {
  const input = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer || "");
  let crc = 0xffffffff;
  for (let i = 0; i < input.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ input[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function findZipEndOfCentralDirectory(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 22) return -1;
  const min = Math.max(0, buffer.length - 0xffff - 22);
  for (let offset = buffer.length - 22; offset >= min; offset -= 1) {
    if (buffer.readUInt32LE(offset) === 0x06054b50) return offset;
  }
  return -1;
}

function parseZipEntries(buffer) {
  const eocdOffset = findZipEndOfCentralDirectory(buffer);
  if (eocdOffset < 0) {
    throw new Error("ZIP_EOCD_NOT_FOUND");
  }

  const entryCount = buffer.readUInt16LE(eocdOffset + 10);
  const centralDirectoryOffset = buffer.readUInt32LE(eocdOffset + 16);
  const commentLength = buffer.readUInt16LE(eocdOffset + 20);
  const eocdComment = buffer.slice(
    eocdOffset + 22,
    eocdOffset + 22 + commentLength,
  );

  const entries = [];
  let cursor = centralDirectoryOffset;
  for (let i = 0; i < entryCount; i += 1) {
    if (buffer.readUInt32LE(cursor) !== 0x02014b50) {
      throw new Error(`ZIP_CENTRAL_DIRECTORY_ENTRY_INVALID_${i}`);
    }

    const versionMadeBy = buffer.readUInt16LE(cursor + 4);
    const versionNeeded = buffer.readUInt16LE(cursor + 6);
    const flags = buffer.readUInt16LE(cursor + 8);
    const method = buffer.readUInt16LE(cursor + 10);
    const modTime = buffer.readUInt16LE(cursor + 12);
    const modDate = buffer.readUInt16LE(cursor + 14);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const uncompressedSize = buffer.readUInt32LE(cursor + 24);
    const fileNameLength = buffer.readUInt16LE(cursor + 28);
    const extraLength = buffer.readUInt16LE(cursor + 30);
    const commentLengthEntry = buffer.readUInt16LE(cursor + 32);
    const diskStart = buffer.readUInt16LE(cursor + 34);
    const internalAttrs = buffer.readUInt16LE(cursor + 36);
    const externalAttrs = buffer.readUInt32LE(cursor + 38);
    const localHeaderOffset = buffer.readUInt32LE(cursor + 42);
    const fileNameBuffer = buffer.slice(
      cursor + 46,
      cursor + 46 + fileNameLength,
    );
    const extra = buffer.slice(
      cursor + 46 + fileNameLength,
      cursor + 46 + fileNameLength + extraLength,
    );
    const comment = buffer.slice(
      cursor + 46 + fileNameLength + extraLength,
      cursor + 46 + fileNameLength + extraLength + commentLengthEntry,
    );
    const fileName = fileNameBuffer.toString("utf8");

    if (buffer.readUInt32LE(localHeaderOffset) !== 0x04034b50) {
      throw new Error(`ZIP_LOCAL_HEADER_INVALID_${fileName}`);
    }
    const localFileNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localHeaderOffset + 28);
    const dataOffset =
      localHeaderOffset + 30 + localFileNameLength + localExtraLength;
    const compressedData = buffer.slice(
      dataOffset,
      dataOffset + compressedSize,
    );

    entries.push({
      versionMadeBy,
      versionNeeded,
      flags,
      method,
      modTime,
      modDate,
      crc: buffer.readUInt32LE(cursor + 16),
      compressedSize,
      uncompressedSize,
      fileNameBuffer,
      fileName,
      extra,
      comment,
      diskStart,
      internalAttrs,
      externalAttrs,
      compressedData,
    });

    cursor += 46 + fileNameLength + extraLength + commentLengthEntry;
  }

  return { entries, eocdComment };
}

function inflateZipEntry(entry) {
  if (entry.method === 0) return Buffer.from(entry.compressedData);
  if (entry.method === 8) return zlib.inflateRawSync(entry.compressedData);
  throw new Error(
    `ZIP_UNSUPPORTED_COMPRESSION_METHOD_${entry.method}_${entry.fileName}`,
  );
}

function deflateZipEntry(entry, uncompressed) {
  if (entry.method === 0) return Buffer.from(uncompressed);
  if (entry.method === 8) return zlib.deflateRawSync(uncompressed);
  throw new Error(
    `ZIP_UNSUPPORTED_COMPRESSION_METHOD_${entry.method}_${entry.fileName}`,
  );
}

function injectCalcPrXml(workbookXml = "") {
  const xml = String(workbookXml || "");
  const calcPrPattern = /<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/i;
  if (calcPrPattern.test(xml)) {
    return xml.replace(calcPrPattern, WORKBOOK_CALCPR_XML);
  }
  if (/<\/workbook>/i.test(xml)) {
    return xml.replace(/<\/workbook>/i, `${WORKBOOK_CALCPR_XML}</workbook>`);
  }
  throw new Error("WORKBOOK_XML_CLOSING_TAG_NOT_FOUND");
}

function buildZipBuffer(entries, eocdComment = Buffer.alloc(0)) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  entries.forEach((entry) => {
    const fileNameBuffer = Buffer.from(
      entry.fileNameBuffer || Buffer.from(entry.fileName, "utf8"),
    );
    const extra = Buffer.from(entry.extra || Buffer.alloc(0));
    const comment = Buffer.from(entry.comment || Buffer.alloc(0));
    const compressedData = Buffer.from(entry.compressedData || Buffer.alloc(0));
    const localHeader = Buffer.alloc(30);

    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(entry.versionNeeded || 20, 4);
    localHeader.writeUInt16LE(entry.flags || 0, 6);
    localHeader.writeUInt16LE(entry.method || 0, 8);
    localHeader.writeUInt16LE(entry.modTime || 0, 10);
    localHeader.writeUInt16LE(entry.modDate || 0, 12);
    localHeader.writeUInt32LE(entry.crc >>> 0, 14);
    localHeader.writeUInt32LE(compressedData.length >>> 0, 18);
    localHeader.writeUInt32LE((entry.uncompressedSize || 0) >>> 0, 22);
    localHeader.writeUInt16LE(fileNameBuffer.length, 26);
    localHeader.writeUInt16LE(extra.length, 28);

    localParts.push(localHeader, fileNameBuffer, extra, compressedData);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(entry.versionMadeBy || 20, 4);
    centralHeader.writeUInt16LE(entry.versionNeeded || 20, 6);
    centralHeader.writeUInt16LE(entry.flags || 0, 8);
    centralHeader.writeUInt16LE(entry.method || 0, 10);
    centralHeader.writeUInt16LE(entry.modTime || 0, 12);
    centralHeader.writeUInt16LE(entry.modDate || 0, 14);
    centralHeader.writeUInt32LE(entry.crc >>> 0, 16);
    centralHeader.writeUInt32LE(compressedData.length >>> 0, 20);
    centralHeader.writeUInt32LE((entry.uncompressedSize || 0) >>> 0, 24);
    centralHeader.writeUInt16LE(fileNameBuffer.length, 28);
    centralHeader.writeUInt16LE(extra.length, 30);
    centralHeader.writeUInt16LE(comment.length, 32);
    centralHeader.writeUInt16LE(entry.diskStart || 0, 34);
    centralHeader.writeUInt16LE(entry.internalAttrs || 0, 36);
    centralHeader.writeUInt32LE((entry.externalAttrs || 0) >>> 0, 38);
    centralHeader.writeUInt32LE(offset >>> 0, 42);

    centralParts.push(centralHeader, fileNameBuffer, extra, comment);
    offset +=
      localHeader.length +
      fileNameBuffer.length +
      extra.length +
      compressedData.length;
  });

  const centralDirectoryOffset = offset;
  const centralDirectory = Buffer.concat(centralParts);
  const centralDirectorySize = centralDirectory.length;
  const comment = Buffer.from(eocdComment || Buffer.alloc(0));
  const eocd = Buffer.alloc(22);

  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralDirectorySize >>> 0, 12);
  eocd.writeUInt32LE(centralDirectoryOffset >>> 0, 16);
  eocd.writeUInt16LE(comment.length, 20);

  return Buffer.concat([...localParts, centralDirectory, eocd, comment]);
}

function injectWorkbookCalcPrIntoXlsxBuffer(buffer) {
  const result = {
    version: WORKBOOK_XML_CALCPR_INJECTION_VERSION,
    applied: false,
    workbookXmlPatched: false,
    workbookXmlPath: "xl/workbook.xml",
    error: "",
    buffer,
  };

  try {
    if (!Buffer.isBuffer(buffer)) {
      throw new Error("XLSX_BUFFER_NOT_BUFFER");
    }
    const { entries, eocdComment } = parseZipEntries(buffer);
    const workbookEntry = entries.find(
      (entry) => entry.fileName === result.workbookXmlPath,
    );
    if (!workbookEntry) {
      throw new Error("WORKBOOK_XML_ENTRY_NOT_FOUND");
    }

    const workbookXml = inflateZipEntry(workbookEntry).toString("utf8");
    const patchedXml = injectCalcPrXml(workbookXml);
    const patchedXmlBuffer = Buffer.from(patchedXml, "utf8");
    const patchedCompressedData = deflateZipEntry(
      workbookEntry,
      patchedXmlBuffer,
    );

    workbookEntry.compressedData = patchedCompressedData;
    workbookEntry.uncompressedSize = patchedXmlBuffer.length;
    workbookEntry.compressedSize = patchedCompressedData.length;
    workbookEntry.crc = crc32(patchedXmlBuffer);

    result.buffer = buildZipBuffer(entries, eocdComment);
    result.applied = true;
    result.workbookXmlPatched = true;
    result.error = "";
  } catch (error) {
    result.error = error?.message || String(error);
  }

  return result;
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
  sourceTablePolicy = null,
} = {}) {
  const dataRowCount = Math.max(0, sourceRows.length - 1);
  const primaryUsage = primaryTable.tableUsage || {};
  const primarySelection = primaryTable.primarySelection || {};
  const dataQuality = primaryTable.dataQuality || {};
  const policySummary = sourceTablePolicy
    ? summarizeSourceTablePolicy(sourceTablePolicy)
    : null;

  const overviewRows = [
    ["항목", "값", "비고"],
    ["원본 파일", fileName || "", "사용자가 업로드한 원본 파일명"],
    ["요청", message || "", "자동화 템플릿 생성 요청"],
    [
      "원본데이터 시트",
      policySummary?.sourceTables
        ?.map((entry) => entry.sourceSheetName)
        .join(", ") || SHEET_NAMES.SOURCE_DATA,
      "1행 헤더, 2행부터 정리된 분석 데이터",
    ],
    [
      "sourceTablePolicy.version",
      policySummary?.version || "",
      "원본데이터 시트 선택 정책",
    ],
    [
      "sourceTablePolicy.scope",
      policySummary?.sourceScope || "",
      "singleTable/multiTable",
    ],
    [
      "sourceTablePolicy.sourceTableCount",
      policySummary?.counts?.sourceTableCount ?? "",
      "원본데이터 시트로 출력된 physical table 수",
    ],
    [
      "sourceTablePolicy.primaryStrategy",
      policySummary?.primaryStrategy || "",
      "대표 테이블 선택 전략",
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
      "sourceTableId",
      "sourceSheetName",
      "sourceScope",
    ],
  ];

  const sourceEntryByTableId = new Map();
  for (const entry of sourceTablePolicy?.sourceTableEntries || []) {
    [entry.tableId, entry.sourceTableId].filter(Boolean).forEach((id) => {
      sourceEntryByTableId.set(id, entry);
    });
  }

  const tableRows = (Array.isArray(tables) ? tables : []).map((table) => {
    const usage = table.tableUsage || {};
    const selection = table.primarySelection || {};
    const quality = table.dataQuality || {};
    const sourceEntry =
      sourceEntryByTableId.get(table.tableId) ||
      sourceEntryByTableId.get(table.sourceTableId) ||
      null;

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
      sourceEntry?.sourceTableId || "",
      sourceEntry?.sourceSheetName || "",
      sourceEntry?.sourceScope || "",
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
  diagnostic = null,
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
  emitSummarySheetDiagnostic(diagnostic, "summary_workbook:start", "INFO", {
    summarySheetMode: normalizedSummarySheetMode,
    includeSourceDataSheet,
    sourceTableCount: Array.isArray(sourceTables) ? sourceTables.length : 0,
    sectionCount: Array.isArray(result?.sections) ? result.sections.length : 0,
    resultType: result?.resultType || "",
    operation: result?.operation || "",
  });

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
    const resolvedRecipeSections = [];

    const summaryRows = buildBusinessWorkbookSummaryRows({
      fileName,
      message,
      result,
      businessSections,
      summarySheetMode: normalizedSummarySheetMode,
      includeSourceDataSheet,
      sourceTables,
    });

    const wsSummary = runWorkbookDiagnosticStep(
      diagnostic,
      "business_workbook:summary_sheet_aoa",
      () => XLSX.utils.aoa_to_sheet(summaryRows),
      { rowCount: summaryRows.length },
    );
    setAoaColumnWidths(wsSummary, summaryRows);
    styleHeaderRow(wsSummary);
    applyDefaultSheetOptions(wsSummary);
    appendSheetSafe(wb, wsSummary, SHEET_NAMES.SUMMARY);

    runWorkbookDiagnosticStep(
      diagnostic,
      "business_workbook:append_source_data",
      () =>
        appendSourceDataSheets(wb, sourceTables, {
          includeSourceDataSheet,
          summarySheetMode: normalizedSummarySheetMode,
        }),
      {
        includeSourceDataSheet,
        sourceTableCount: Array.isArray(sourceTables) ? sourceTables.length : 0,
      },
    );

    businessSections.forEach((section, index) => {
      const sectionMeta = summarizeSectionForWorkbookDiagnostic(section, index);
      emitSummarySheetDiagnostic(
        diagnostic,
        "business_section:start",
        "INFO",
        sectionMeta,
      );
      const sectionResult = section.result || {};
      const rows = runWorkbookDiagnosticStep(
        diagnostic,
        "business_section:result_to_rows",
        () => resultToRows(sectionResult),
        sectionMeta,
      );
      const outputRows = rows.length
        ? rows
        : [
            {
              섹션: section.title || section.sectionId || `섹션_${index + 1}`,
              결과: "데이터 없음",
            },
          ];
      emitSummarySheetDiagnostic(
        diagnostic,
        "business_section:rows_ready",
        "INFO",
        {
          ...sectionMeta,
          rowCount: rows.length,
          outputRowCount: outputRows.length,
          columnCount:
            outputRows.length &&
            outputRows[0] &&
            typeof outputRows[0] === "object"
              ? Object.keys(outputRows[0]).length
              : 0,
        },
      );
      const ws = runWorkbookDiagnosticStep(
        diagnostic,
        "business_section:json_to_sheet",
        () => XLSX.utils.json_to_sheet(outputRows),
        { ...sectionMeta, outputRowCount: outputRows.length },
      );
      const formulaPlan = runWorkbookDiagnosticStep(
        diagnostic,
        "business_section:formula_plan",
        () =>
          buildSectionFormulaPlan({
            section,
            rows,
            formulaContext,
          }),
        sectionMeta,
      );

      runWorkbookDiagnosticStep(
        diagnostic,
        "business_section:format_sheet",
        () => {
          setColumnWidths(ws, outputRows);
          formatNumberCells(ws, sectionResult);
          styleHeaderRow(ws);
          applyDefaultSheetOptions(ws);
          ws["!freeze"] = { xSplit: 0, ySplit: 1 };
        },
        sectionMeta,
      );
      const formulaCount = runWorkbookDiagnosticStep(
        diagnostic,
        "business_section:apply_formulas",
        () =>
          applySimpleAggregateFormulaSection({
            ws,
            rows,
            formulaPlan,
            formulaContext,
          }),
        sectionMeta,
      );
      recordFormulaApplication(formulaEngineMeta, formulaCount);
      attachFormulaPlan(ws, formulaPlan);

      const requestedSheetName =
        section.title || section.sectionId || `섹션_${index + 1}`;
      runWorkbookDiagnosticStep(
        diagnostic,
        "business_section:append_sheet",
        () =>
          appendSheetSafe(wb, ws, requestedSheetName, {
            onResolved: (resolvedSheetName) => {
              resolvedRecipeSections.push({
                index,
                sectionId: section.sectionId || "",
                title: section.title || "",
                resolvedSheetName: resolvedSheetName.safeSheetName || "",
              });
              emitSummarySheetDiagnostic(
                diagnostic,
                "business_section:resolved_sheet_name",
                "INFO",
                {
                  ...sectionMeta,
                  formulaCount,
                  ...resolvedSheetName,
                },
              );
            },
          }),
        {
          ...sectionMeta,
          formulaCount,
          requestedSheetName: String(requestedSheetName || ""),
          existingSheetCount: Array.isArray(wb.SheetNames)
            ? wb.SheetNames.length
            : 0,
        },
      );
    });

    runWorkbookDiagnosticStep(
      diagnostic,
      "business_workbook:append_recipe_manifest",
      () =>
        appendSummarySheetRecipeManifest(wb, {
          result,
          businessSections,
          resolvedSections: resolvedRecipeSections,
        }),
      {
        manifestVersion: SUMMARY_SHEET_RECIPE_MANIFEST_VERSION,
        sectionCount: businessSections.length,
      },
    );

    runWorkbookDiagnosticStep(
      diagnostic,
      "business_workbook:finalize",
      () => {
        attachWorkbookFormulaEngineMeta(wb, formulaEngineMeta);
        ensureWorkbookRecalculationOnOpen(
          wb,
          "business_sections_formula_workbook",
        );
      },
      { sheetCount: wb.SheetNames?.length || 0 },
    );
    emitSummarySheetDiagnostic(diagnostic, "business_workbook:complete", "OK", {
      sheetCount: wb.SheetNames?.length || 0,
    });
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
  ensureWorkbookRecalculationOnOpen(wb, "summary_workbook");
  return wb;
}

function workbookToBuffer(workbook) {
  ensureWorkbookRecalculationOnOpen(workbook, "workbook_to_buffer_final_guard");
  const rawBuffer = XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
  });

  const injected = injectWorkbookCalcPrIntoXlsxBuffer(rawBuffer);
  const previousMeta = workbook?.["!beebeeWorkbookRecalculation"] || {};
  if (workbook && typeof workbook === "object") {
    workbook["!beebeeWorkbookRecalculation"] = {
      ...previousMeta,
      version: WORKBOOK_RECALCULATION_FLAG_VERSION,
      applied: injected.applied,
      calcMode: "auto",
      fullCalcOnLoad: true,
      forceFullCalc: true,
      reason: previousMeta.reason || "workbook_to_buffer_final_guard",
      injectionVersion: injected.version,
      workbookXmlPatched: injected.workbookXmlPatched,
      workbookXmlPath: injected.workbookXmlPath,
      injectionMethod: "xlsx_zip_workbook_xml_patch",
      injectionError: injected.error || "",
    };
  }

  return injected.buffer;
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

function asNonEmptyArray(value) {
  if (Array.isArray(value))
    return value.filter((item) => item != null && item !== "");
  if (value == null || value === "") return [];
  return [value];
}

function resolveAutomationCandidateTitle({
  result = {},
  candidate = null,
  templateCandidate = null,
} = {}) {
  return (
    candidate?.title ||
    candidate?.candidateTitle ||
    templateCandidate?.title ||
    templateCandidate?.templateTitle ||
    result?.title ||
    result?.templateTitle ||
    result?.operation ||
    "자동화 후보"
  );
}

function resolveAutomationCandidateType({
  result = {},
  candidate = null,
  templateCandidate = null,
} = {}) {
  return (
    candidate?.candidateType ||
    candidate?.recipeType ||
    candidate?.type ||
    templateCandidate?.templateId ||
    result?.recipeType ||
    result?.operation ||
    result?.resultType ||
    "sourceDataOnly"
  );
}

function resolvePreferredSourceTableId({
  result = {},
  intent = {},
  candidate = null,
  templateCandidate = null,
} = {}) {
  return (
    result?.sourceTableId ||
    result?.tableId ||
    result?.table?.tableId ||
    candidate?.sourceTableId ||
    asNonEmptyArray(candidate?.sourceTableIds)[0] ||
    candidate?.tableId ||
    templateCandidate?.sourceTableId ||
    asNonEmptyArray(templateCandidate?.sourceTableIds)[0] ||
    intent?.table?.tableId ||
    intent?.sourceTableId ||
    ""
  );
}

function buildCandidateSummaryAoa({
  result = {},
  candidate = null,
  templateCandidate = null,
  sourceTablePolicy = null,
} = {}) {
  const candidateTitle = resolveAutomationCandidateTitle({
    result,
    candidate,
    templateCandidate,
  });
  const candidateType = resolveAutomationCandidateType({
    result,
    candidate,
    templateCandidate,
  });
  const sourceSheetNames = asNonEmptyArray(
    candidate?.sourceSheetNames ||
      result?.sourceSheetName ||
      sourceTablePolicy?.primarySourceSheetName,
  );
  const sourceTableIds = asNonEmptyArray(
    candidate?.sourceTableIds ||
      result?.sourceTableId ||
      sourceTablePolicy?.primarySourceTableId,
  );

  return [
    ["항목", "값"],
    ["자동화시트버전", AUTOMATION_SHEET_V2_VERSION],
    ["후보명", candidateTitle],
    ["후보유형", candidateType],
    [
      "후보ID",
      candidate?.candidateId ||
        candidate?.id ||
        templateCandidate?.templateId ||
        "",
    ],
    [
      "recipeIds",
      asNonEmptyArray(candidate?.recipeIds || result?.recipeType).join(", "),
    ],
    [
      "sourceScope",
      candidate?.sourceScope || sourceTablePolicy?.sourceScope || "",
    ],
    ["sourceTableIds", sourceTableIds.join(", ")],
    ["sourceSheetNames", sourceSheetNames.join(", ")],
    ["rankScore", candidate?.rankScore ?? ""],
    ["rankingTier", candidate?.rankingTier || ""],
    ["reasonCodes", flattenReasonList(candidate?.reasonCodes || [])],
    [
      "outputTypes",
      asNonEmptyArray(candidate?.outputTypes || ["summarySheet"]).join(", "),
    ],
    ["resultType", result?.resultType || ""],
    ["operation", result?.operation || ""],
    [
      "rowCount",
      Array.isArray(result?.rows) ? result.rows.length : result?.rowCount || 0,
    ],
  ];
}

function buildSumIfsFormula({
  valueRange,
  criteriaRange1,
  criteriaCell1,
  criteriaRange2 = "",
  criteriaCell2 = "",
}) {
  if (criteriaRange2 && criteriaCell2) {
    return `IFERROR(SUMIFS(${valueRange},${criteriaRange1},${criteriaCell1},${criteriaRange2},${criteriaCell2}),0)`;
  }
  return `IFERROR(SUMIFS(${valueRange},${criteriaRange1},${criteriaCell1}),0)`;
}

function buildAverageIfsFormula({
  valueRange,
  criteriaRange1,
  criteriaCell1,
  criteriaRange2 = "",
  criteriaCell2 = "",
}) {
  if (criteriaRange2 && criteriaCell2) {
    return `IFERROR(AVERAGEIFS(${valueRange},${criteriaRange1},${criteriaCell1},${criteriaRange2},${criteriaCell2}),0)`;
  }
  return `IFERROR(AVERAGEIFS(${valueRange},${criteriaRange1},${criteriaCell1}),0)`;
}

function buildCountIfsFormulaV2({
  criteriaRange1,
  criteriaCell1,
  criteriaRange2 = "",
  criteriaCell2 = "",
}) {
  if (criteriaRange2 && criteriaCell2) {
    return `IFERROR(COUNTIFS(${criteriaRange1},${criteriaCell1},${criteriaRange2},${criteriaCell2}),0)`;
  }
  return buildCountIfsFormula({
    criteriaRange: criteriaRange1,
    criteriaCell: criteriaCell1,
  });
}

function resolveResultLabel(row = {}, result = {}, fallbackIndex = 0) {
  const groupHeader = result?.groupBy?.header || result?.date?.header || "";
  if (row.period != null) return row.period;
  if (groupHeader && row[groupHeader] != null) return row[groupHeader];
  if (row.label != null) return row.label;
  const firstValue = Object.values(row || {}).find(
    (value) => value != null && value !== "",
  );
  return firstValue ?? `항목_${fallbackIndex + 1}`;
}

function resolveResultLabel2(row = {}, result = {}) {
  const groupHeader2 = result?.groupBy2?.header || "";
  if (groupHeader2 && row[groupHeader2] != null) return row[groupHeader2];
  if (row.dimension2 != null) return row.dimension2;
  return "";
}

function uniqueResultLabels(resultRows = [], result = {}) {
  const seen = new Set();
  const values = [];
  resultRows.forEach((row, index) => {
    const value = resolveResultLabel(row, result, index);
    const key = String(value ?? "").trim();
    if (!key || seen.has(key)) return;
    seen.add(key);
    values.push(value);
  });
  return values;
}

function uniqueResultPairs(resultRows = [], result = {}) {
  const seen = new Set();
  const values = [];
  resultRows.forEach((row, index) => {
    const a = resolveResultLabel(row, result, index);
    const b = resolveResultLabel2(row, result);
    const key = `${a}||${b}`;
    if (!String(a ?? "").trim() || seen.has(key)) return;
    seen.add(key);
    values.push([a, b]);
  });
  return values;
}

function operationFormulaType(rawOperation = "") {
  const op = String(rawOperation || "").toLowerCase();
  if (op.includes("count")) return "count";
  if (op.includes("average") || op.includes("avg")) return "average";
  return "sum";
}

function buildAutomationSheetV2({
  result = {},
  candidate = null,
  templateCandidate = null,
  primarySourceSheetName = SHEET_NAMES.SOURCE_DATA,
  groupLetter = "A",
  metricLetter = "B",
  groupHeader = "",
  metricHeader = "",
  columns = [],
  orderedColumns = [],
  derivedGroupLetter = null,
} = {}) {
  const resultRows = Array.isArray(result?.rows) ? result.rows : [];
  const rawOperation =
    result?.operation || candidate?.recipeType || candidate?.type || "sum";
  const resultType = result?.resultType || "grouped";
  const candidateTitle = resolveAutomationCandidateTitle({
    result,
    candidate,
    templateCandidate,
  });
  const candidateType = resolveAutomationCandidateType({
    result,
    candidate,
    templateCandidate,
  });
  const valueHeader = metricHeader || result?.metric?.header || "값";
  const labelHeader =
    resultType === "timeSeries"
      ? "기간"
      : groupHeader || result?.groupBy?.header || "기준";
  const formulaGroupLetter = derivedGroupLetter || groupLetter;
  const labelRange = buildColumnRange(
    primarySourceSheetName,
    formulaGroupLetter,
  );
  const valueRange = buildColumnRange(primarySourceSheetName, metricLetter);
  const formulaType = operationFormulaType(rawOperation);

  const rows = [
    ["자동화시트 v2"],
    ["후보명", candidateTitle],
    ["후보유형", candidateType],
    ["원본시트", primarySourceSheetName],
    ["분석유형", `${resultType} / ${rawOperation}`],
    [],
  ];

  const opLower = String(rawOperation || "").toLowerCase();
  const recipeLower = String(
    candidate?.recipeType || candidate?.type || "",
  ).toLowerCase();
  const isTopBottom =
    opLower.includes("topbottom") || recipeLower.includes("top_bottom");
  const isComposition =
    opLower.includes("composition") ||
    recipeLower.includes("composition_ratio");
  const isCumulative = opLower.includes("cumulative");
  const isGrowth = opLower.includes("growth");
  const isCross = resultType === "crossTable" || opLower.startsWith("cross");
  const isTime = resultType === "timeSeries" || /time|trend|wide/.test(opLower);

  if (isTopBottom) {
    rows.push(["순위", "항목", valueHeader, "구분"]);
    const rowCount = Math.max(10, Math.min(20, resultRows.length || 10));
    for (let i = 1; i <= rowCount; i += 1)
      rows.push([i, null, null, i <= 5 ? "상위" : "하위"]);
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    for (let i = 1; i <= rowCount; i += 1) {
      const rowNum = 7 + i;
      const rank = i <= 5 ? i : Math.max(1, rowCount - i + 1);
      sheet[`C${rowNum}`] = {
        t: "n",
        f: buildRankValueFormula({ valueRange, rank }),
        v: 0,
      };
      sheet[`B${rowNum}`] = {
        t: "s",
        f: buildRankLabelFormula({
          labelRange,
          valueRange,
          rankValueCell: `C${rowNum}`,
        }),
        v: "",
      };
    }
    sheet["!ref"] = `A1:D${rows.length}`;
    return {
      sheet,
      rows,
      meta: {
        version: AUTOMATION_SHEET_V2_VERSION,
        layout: "topBottom",
        rowCount,
      },
    };
  }

  if (isCross) {
    const secondHeader =
      result?.groupBy2?.header || candidate?.columns?.dimension2 || "비교기준";
    const secondCol = columns.find((c) => c.header === secondHeader) || {};
    const secondLetter = resolveSourceDataColumnLetter(
      orderedColumns,
      secondCol,
      groupLetter,
    );
    const secondRange = buildColumnRange(primarySourceSheetName, secondLetter);
    const pairs = uniqueResultPairs(resultRows, result);
    rows.push([labelHeader, secondHeader, valueHeader]);
    pairs.forEach(([a, b]) => rows.push([a, b, null]));
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    pairs.forEach((_, index) => {
      const rowNum = 8 + index;
      const formula =
        formulaType === "count"
          ? buildCountIfsFormulaV2({
              criteriaRange1: labelRange,
              criteriaCell1: `A${rowNum}`,
              criteriaRange2: secondRange,
              criteriaCell2: `B${rowNum}`,
            })
          : buildSumIfsFormula({
              valueRange,
              criteriaRange1: labelRange,
              criteriaCell1: `A${rowNum}`,
              criteriaRange2: secondRange,
              criteriaCell2: `B${rowNum}`,
            });
      sheet[`C${rowNum}`] = { t: "n", f: formula, v: 0 };
    });
    sheet["!ref"] = `A1:C${Math.max(rows.length, 8)}`;
    return {
      sheet,
      rows,
      meta: {
        version: AUTOMATION_SHEET_V2_VERSION,
        layout: "crossTable",
        rowCount: pairs.length,
      },
    };
  }

  const labels = uniqueResultLabels(resultRows, result);
  const safeLabels = labels.length
    ? labels
    : Array.from({ length: 10 }, (_, index) => `항목_${index + 1}`);
  const headers = [labelHeader, valueHeader];
  if (isComposition) headers.push("구성비");
  if (isCumulative) headers.push("누적값");
  if (isGrowth) headers.push("증감률");
  if (isTime && !isCumulative && !isGrowth) headers.push("전기대비");
  const maxColumns = Math.max(headers.length, 4);
  rows.push(headers);
  safeLabels.forEach((value) =>
    rows.push([value, null, null, null].slice(0, maxColumns)),
  );

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  const firstDataRow = 8;
  safeLabels.forEach((_, index) => {
    const rowNum = firstDataRow + index;
    let formula = buildGroupAggregateFormula({
      operation: formulaType,
      sheetName: primarySourceSheetName,
      groupLetter: formulaGroupLetter,
      metricLetter,
      criteriaCell: `A${rowNum}`,
    });
    if (formulaType === "count") {
      formula = buildCountIfsFormula({
        criteriaRange: labelRange,
        criteriaCell: `A${rowNum}`,
      });
    }
    if (formulaType === "average") {
      formula = buildAverageIfsFormula({
        valueRange,
        criteriaRange1: labelRange,
        criteriaCell1: `A${rowNum}`,
      });
    }
    sheet[`B${rowNum}`] = { t: "n", f: formula, v: 0 };

    let colIndex = 2;
    if (isComposition) {
      const col = XLSX.utils.encode_col(colIndex);
      sheet[`${col}${rowNum}`] = {
        t: "n",
        f: `IFERROR(B${rowNum}/SUM(B${firstDataRow}:B${firstDataRow + safeLabels.length - 1}),"")`,
        v: 0,
      };
      colIndex += 1;
    }
    if (isCumulative) {
      const col = XLSX.utils.encode_col(colIndex);
      sheet[`${col}${rowNum}`] = {
        t: "n",
        f:
          rowNum === firstDataRow
            ? `IFERROR(B${rowNum},"")`
            : buildRunningSumFormula({
                valueCell: `B${rowNum}`,
                previousCell: `${col}${rowNum - 1}`,
              }),
        v: 0,
      };
      colIndex += 1;
    }
    if (isGrowth || (isTime && !isCumulative)) {
      const col = XLSX.utils.encode_col(colIndex);
      sheet[`${col}${rowNum}`] = {
        t: "n",
        f:
          rowNum === firstDataRow
            ? `""`
            : buildGrowthRateFormula({
                currentCell: `B${rowNum}`,
                previousCell: `B${rowNum - 1}`,
              }),
        v: 0,
      };
    }
  });

  sheet["!ref"] =
    `A1:${XLSX.utils.encode_col(maxColumns - 1)}${Math.max(rows.length, firstDataRow)}`;
  return {
    sheet,
    rows,
    meta: {
      version: AUTOMATION_SHEET_V2_VERSION,
      layout: isTime ? "timeSeries" : isComposition ? "composition" : "grouped",
      rowCount: safeLabels.length,
      formulaType,
      sourceSheetName: primarySourceSheetName,
    },
  };
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
  candidate = null,
  templateCandidate = null,
  tables = [],
}) {
  const wb = XLSX.utils.book_new();

  const allTables = Array.isArray(tables) ? tables.filter(Boolean) : [];
  const sourceTablePolicy = buildSourceTablePolicy({
    tables: allTables,
    preferredTableId: resolvePreferredSourceTableId({
      result,
      intent,
      candidate,
      templateCandidate,
    }),
    sourceSheetBaseName: SHEET_NAMES.SOURCE_DATA,
  });
  const sourceTables = sourceTablePolicy.sourceTables || [];
  const sourceTableEntries = sourceTablePolicy.sourceTableEntries || [];
  const primaryTable = sourceTablePolicy.primaryTable || {};
  const primarySourceSheetName =
    sourceTablePolicy.primarySourceSheetName || SHEET_NAMES.SOURCE_DATA;
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
      ["원본데이터 정책", sourceTablePolicy.version],
      ["원본데이터 범위", sourceTablePolicy.sourceScope],
      ["원본데이터 시트 수", sourceTablePolicy.counts?.sourceTableCount || 0],
      ["자동화시트 버전", AUTOMATION_SHEET_V2_VERSION],
      [
        "후보",
        resolveAutomationCandidateTitle({
          result,
          candidate,
          templateCandidate,
        }),
      ],
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
      [
        "원본데이터정책",
        sourceTablePolicy.version,
        "source table policy version",
      ],
      [
        "원본데이터범위",
        sourceTablePolicy.sourceScope,
        "singleTable/multiTable",
      ],
      [
        "원본데이터시트수",
        sourceTablePolicy.counts?.sourceTableCount || 0,
        "생성된 원본데이터 시트 수",
      ],
      [
        "자동화시트버전",
        AUTOMATION_SHEET_V2_VERSION,
        "candidate-aware automation sheet layout",
      ],
      [
        "후보명",
        resolveAutomationCandidateTitle({
          result,
          candidate,
          templateCandidate,
        }),
        "후보/템플릿 제목",
      ],
      [
        "후보유형",
        resolveAutomationCandidateType({
          result,
          candidate,
          templateCandidate,
        }),
        "candidateType/recipeType/templateId",
      ],
      [
        "후보ID",
        candidate?.candidateId ||
          candidate?.id ||
          templateCandidate?.templateId ||
          result?.recipeType ||
          "",
        "후보 식별자",
      ],
      ["rankScore", candidate?.rankScore ?? "", "Candidate Scoring v1 점수"],
      [
        "reasonCodes",
        flattenReasonList(candidate?.reasonCodes || []),
        "후보 추천 사유",
      ],
    ]),
    SHEET_NAMES.AUTOMATION_SETTINGS,
  );

  const automationSheetV2 = buildAutomationSheetV2({
    result,
    candidate,
    templateCandidate,
    primarySourceSheetName,
    groupLetter,
    metricLetter,
    groupHeader,
    metricHeader,
    columns,
    orderedColumns,
    groupCol,
    metricCol,
    derivedGroupLetter,
  });
  const autoSheet = automationSheetV2.sheet;
  setAoaColumnWidths(autoSheet, automationSheetV2.rows || []);
  styleHeaderRow(autoSheet);
  applyDefaultSheetOptions(autoSheet);
  XLSX.utils.book_append_sheet(wb, autoSheet, SHEET_NAMES.AUTOMATION_TEMPLATE);

  const candidateSummaryAoa = buildCandidateSummaryAoa({
    result,
    candidate,
    templateCandidate,
    sourceTablePolicy,
  });
  const candidateSummarySheet = XLSX.utils.aoa_to_sheet(candidateSummaryAoa);
  setAoaColumnWidths(candidateSummarySheet, candidateSummaryAoa);
  styleHeaderRow(candidateSummarySheet);
  applyDefaultSheetOptions(candidateSummarySheet);
  appendSheetSafe(wb, candidateSummarySheet, "후보요약");

  const sourceEntriesToWrite = sourceTableEntries.length
    ? sourceTableEntries
    : [
        {
          table: primaryTable,
          sourceSheetName: primarySourceSheetName,
          sourceTableIndex: 0,
        },
      ].filter((entry) => entry.table && Object.keys(entry.table).length);

  sourceEntriesToWrite.forEach((entry) => {
    const table = entry.table;
    const sourceData =
      table === primaryTable ? cleanSourceData : buildCleanSourceDataAoa(table);
    const tableSourceRows = sourceData.sourceRows?.length
      ? sourceData.sourceRows
      : [["데이터 없음"]];
    const sourceDataSheet = XLSX.utils.aoa_to_sheet(tableSourceRows);
    setAoaColumnWidths(sourceDataSheet, tableSourceRows);
    styleHeaderRow(sourceDataSheet);
    applyDefaultSheetOptions(sourceDataSheet);
    appendSheetSafe(wb, sourceDataSheet, entry.sourceSheetName);
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
    sourceTablePolicy,
  });
  const diagnosticsSheet = XLSX.utils.aoa_to_sheet(diagnosticsAoa);
  setAoaColumnWidths(diagnosticsSheet, diagnosticsAoa);
  styleHeaderRow(diagnosticsSheet);
  applyDefaultSheetOptions(diagnosticsSheet);
  XLSX.utils.book_append_sheet(wb, diagnosticsSheet, SHEET_NAMES.DIAGNOSTICS);

  const previewSheet = XLSX.utils.json_to_sheet(result?.rows || []);
  applyDefaultSheetOptions(previewSheet);
  XLSX.utils.book_append_sheet(wb, previewSheet, SHEET_NAMES.EXECUTION_PREVIEW);

  wb["!beebeeSourceTablePolicy"] =
    summarizeSourceTablePolicy(sourceTablePolicy);
  wb["!beebeeAutomationSheetV2"] = {
    version: AUTOMATION_SHEET_V2_VERSION,
    applied: true,
    layout: automationSheetV2.meta?.layout || "unknown",
    rowCount: automationSheetV2.meta?.rowCount || 0,
    sourceSheetName: primarySourceSheetName,
    candidateTitle: resolveAutomationCandidateTitle({
      result,
      candidate,
      templateCandidate,
    }),
    candidateType: resolveAutomationCandidateType({
      result,
      candidate,
      templateCandidate,
    }),
  };
  ensureWorkbookRecalculationOnOpen(wb, "automation_sheet_v2_formula_workbook");

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
  AUTOMATION_SHEET_V2_VERSION,
  SUMMARY_SHEET_RECIPE_MANIFEST_VERSION,
  WORKBOOK_RECALCULATION_FLAG_VERSION,
  ensureWorkbookRecalculationOnOpen,
  injectWorkbookCalcPrIntoXlsxBuffer,
};
