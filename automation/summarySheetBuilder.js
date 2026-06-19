const XLSX = require("xlsx");
const { buildNarrativeSections } = require("./reportNarrativeBuilder");
const { recommendChartSpec } = require("./chartRecommendationBuilder");
const { buildReportSections } = require("./reportSectionBuilder");
const {
  buildColumnRange,
  buildGroupAggregateFormula,
  buildRankValueFormula,
  buildRankLabelFormula,
  buildRunningSumFormula,
  buildGrowthRateFormula,
  buildMaxIfFormula,
  buildCountIfsFormula,
  buildPivotAverageFormula,
} = require("../builders/automationFormulaBuilder");

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

function appendSheetSafe(wb, ws, name) {
  let safeName = String(name || "Sheet").slice(0, 31);
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

function buildSummaryWorkbook({ fileName, message, intent, result }) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    ["요청", message || ""],
    ["원본 파일", fileName || ""],
    ["테이블", result?.table?.tableName || intent?.table?.tableName || ""],
    ["결과 유형", result?.resultType || ""],
    ["작업", result?.operation || intent?.operation || ""],
    ["지표", result?.metric?.header || intent?.metric?.header || ""],
    ["그룹 기준", result?.groupBy?.header || intent?.groupBy?.header || ""],
    ["결과 행 수", Array.isArray(result?.rows) ? result.rows.length : 0],
    ["생성일시", new Date().toISOString()],
    [],
  ];

  const businessSections = Array.isArray(result?.sections)
    ? result.sections
    : [];

  if (businessSections.length) {
    const wb = XLSX.utils.book_new();

    const summaryRows = [
      ["요청", message || ""],
      ["원본 파일", fileName || ""],
      ["템플릿", result.title || result.templateId || ""],
      ["결과 유형", result.resultType || ""],
      ["섹션 수", businessSections.length],
      ["생성일시", new Date().toISOString()],
    ];

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    setAoaColumnWidths(wsSummary, summaryRows);
    styleHeaderRow(wsSummary);
    applyDefaultSheetOptions(wsSummary);
    appendSheetSafe(wb, wsSummary, "요약");

    businessSections.forEach((section, index) => {
      const sectionResult = section.result || {};
      const rows = resultToRows(sectionResult);
      const ws = XLSX.utils.json_to_sheet(
        rows.length ? rows : [{ 결과: "데이터 없음" }],
      );

      setColumnWidths(ws, rows);
      formatNumberCells(ws, sectionResult);
      styleHeaderRow(ws);
      applyDefaultSheetOptions(ws);
      ws["!freeze"] = { xSplit: 0, ySplit: 1 };

      appendSheetSafe(
        wb,
        ws,
        section.title || section.sectionId || `섹션_${index + 1}`,
      );
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
  appendSheetSafe(wb, wsSummary, "요약");

  const wsResult = XLSX.utils.json_to_sheet(rows);
  setColumnWidths(wsResult, rows);
  formatNumberCells(wsResult, result);
  styleHeaderRow(wsResult);
  applyDefaultSheetOptions(wsResult);
  wsResult["!freeze"] = { xSplit: 0, ySplit: 1 };
  appendSheetSafe(wb, wsResult, "분석결과");

  if (result?.plan) {
    const planRows = objectToAoa(result.plan);
    const wsPlan = XLSX.utils.aoa_to_sheet([["항목", "값"], ...planRows]);
    setAoaColumnWidths(wsPlan, [["항목", "값"], ...planRows]);
    styleHeaderRow(wsPlan);
    applyDefaultSheetOptions(wsPlan);
    XLSX.utils.book_append_sheet(wb, wsPlan, "실행계획");
  }

  if (result?.executionMeta) {
    const metaRows = objectToAoa(result.executionMeta);
    const wsMeta = XLSX.utils.aoa_to_sheet([["항목", "값"], ...metaRows]);
    setAoaColumnWidths(wsMeta, [["항목", "값"], ...metaRows]);
    styleHeaderRow(wsMeta);
    applyDefaultSheetOptions(wsMeta);
    XLSX.utils.book_append_sheet(wb, wsMeta, "실행메타");
  }

  const chartDataRows = buildChartDataRows(result);
  if (chartDataRows.length) {
    const wsChartData = XLSX.utils.json_to_sheet(chartDataRows);
    setColumnWidths(wsChartData, chartDataRows);
    formatNumberCells(wsChartData, result);
    styleHeaderRow(wsChartData);
    applyDefaultSheetOptions(wsChartData);
    wsChartData["!freeze"] = { xSplit: 0, ySplit: 1 };
    appendSheetSafe(wb, wsChartData, "차트데이터");
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
    appendSheetSafe(wb, wsChartSpec, "차트설정");
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
    appendSheetSafe(wb, wsInsight, "인사이트");
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
    appendSheetSafe(wb, wsReportSections, "보고서구성");
  }

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

  const dateColIndex = (dateCol.columnIndex || 1) - 1;
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

function buildAutomationTemplateWorkbook({
  fileName = "",
  message = "",
  intent = null,
  result = null,
  tables = [],
}) {
  const wb = XLSX.utils.book_new();

  const primaryTable = tables.find((t) => t.isPrimary) || tables[0] || {};
  const columns = primaryTable.columns || [];
  const maxColIndex = Math.max(...columns.map((c) => c.columnIndex || 0), 1);

  let headers = Array.from({ length: maxColIndex }, (_, i) => {
    const col = columns.find((c) => c.columnIndex === i + 1);
    return col?.header || "";
  });

  let orderedColumns = Array.from({ length: maxColIndex }, (_, i) => {
    return columns.find((c) => c.columnIndex === i + 1) || null;
  });

  let sourceRows = [
    headers,
    ...(primaryTable.rows || []).slice(0, 200).map((row) =>
      orderedColumns.map((col) => {
        if (!col) return "";
        return getRowValueByColumn(row, col);
      }),
    ),
  ];

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
    "사용방법",
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
    derivedGroupLetter || groupCol.columnLetter || groupCol.letter || "A";
  const metricLetter = metricCol.columnLetter || metricCol.letter || "B";
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
      ["원본시트명", "원본데이터", "데이터가 들어있는 시트명"],
      ["기준열", groupLetter, "부서/월/분류 등 그룹 기준 열"],
      ["값열", metricLetter, "합계/평균 계산 대상 열"],
      ["집계방식", operation, "average, sum, count 중 선택"],
      ["요청문", message, "자동 생성 기준 요청"],
    ]),
    "자동화설정",
  );

  const resultRows = Array.isArray(result?.rows) ? result.rows : [];
  const autoGroupHeader = groupHeader || "기준";

  const uniqueValues = resultRows
    .map((row) => row[autoGroupHeader])
    .filter((v) => v !== undefined && v !== null && v !== "");

  const labelRange = buildColumnRange("원본데이터", groupLetter);
  const valueRange = buildColumnRange("원본데이터", metricLetter);

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
      derivedGroupLetter || rowCol.columnLetter || rowCol.letter || groupLetter;
    const colLetter = colCol.columnLetter || colCol.letter || groupLetter;

    const rowRange = buildColumnRange("원본데이터", rowLetter);
    const colRange = buildColumnRange("원본데이터", colLetter);

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
          sheetName: "원본데이터",
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
        sheetName: "원본데이터",
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

  XLSX.utils.book_append_sheet(wb, autoSheet, "자동화시트");

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet(sourceRows.length ? sourceRows : [["데이터 없음"]]),
    "원본데이터",
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result?.rows || []),
    "실행결과_미리보기",
  );

  return wb;
}

module.exports = {
  buildSummaryWorkbook,
  buildAutomationTemplateWorkbook,
  workbookToBuffer,
  buildChartSpec,
};
