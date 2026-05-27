const XLSX = require("xlsx");

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
    return (result.rows || []).map((r) => ({
      [result.groupBy?.header || "그룹"]: r[result.groupBy?.header] ?? "",
      작업: r.operation,
      지표: r.metric,
      값: r.value,
      행수: r.rowCount,
    }));
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
};
