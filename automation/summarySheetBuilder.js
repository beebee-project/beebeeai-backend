const XLSX = require("xlsx");

function buildChartDataRows(result = {}) {
  if (result.resultType !== "grouped") return [];

  const groupHeader = result.groupBy?.header || "그룹";
  const metricHeader = result.metric?.header || "값";

  return (result.rows || []).map((r) => ({
    [groupHeader]: r[groupHeader] ?? "",
    [metricHeader]: r.value,
    행수: r.rowCount,
  }));
}

function buildChartSpec(result = {}) {
  if (result.resultType !== "grouped") return null;

  const groupHeader = result.groupBy?.header || "그룹";
  const metricHeader = result.metric?.header || "값";

  return {
    version: "chart_spec_v1",
    recommendedType: "bar",
    title: `${groupHeader}별 ${metricHeader}`,
    categoryField: groupHeader,
    valueField: metricHeader,
    rowCount: Array.isArray(result.rows) ? result.rows.length : 0,
  };
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

function formatNumberCells(ws) {
  if (!ws["!ref"]) return;

  const range = XLSX.utils.decode_range(ws["!ref"]);

  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];

      if (!cell || cell.t !== "n") continue;

      if (Number.isInteger(cell.v)) {
        cell.z = "#,##0";
      } else {
        cell.z = "#,##0.00";
      }
    }
  }
}

function resultToRows(result = {}) {
  if (result.resultType === "grouped") {
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

  return result.rows || [];
}

function buildSummaryWorkbook({ fileName, message, intent, result }) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    ["요청", message || ""],
    ["원본 파일", fileName || ""],
    ["테이블", result?.table?.tableName || intent?.table?.tableName || ""],
    ["작업", result?.operation || intent?.operation || ""],
    ["생성일시", new Date().toISOString()],
    [],
  ];

  const rows = resultToRows(result);

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  setAoaColumnWidths(wsSummary, summaryRows);
  XLSX.utils.book_append_sheet(wb, wsSummary, "요약");

  const wsResult = XLSX.utils.json_to_sheet(rows);
  setColumnWidths(wsResult, rows);
  formatNumberCells(wsResult);
  wsResult["!freeze"] = { xSplit: 0, ySplit: 1 };
  XLSX.utils.book_append_sheet(wb, wsResult, "분석결과");

  const chartDataRows = buildChartDataRows(result);
  if (chartDataRows.length) {
    const wsChartData = XLSX.utils.json_to_sheet(chartDataRows);
    setColumnWidths(wsChartData, chartDataRows);
    formatNumberCells(wsChartData);
    wsChartData["!freeze"] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, wsChartData, "차트데이터");
  }

  return wb;
}

function workbookToBuffer(workbook) {
  return XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
  });
}

module.exports = {
  buildSummaryWorkbook,
  workbookToBuffer,
  buildChartSpec,
};
