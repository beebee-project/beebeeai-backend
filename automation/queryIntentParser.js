function normalizeText(v = "") {
  return String(v)
    .toLowerCase()
    .replace(/\([^)]*\)/g, "")
    .replace(/[^가-힣a-z0-9]/gi, "")
    .trim();
}

function includesAny(s, words = []) {
  return words.some((w) => s.includes(w));
}

function findColumnByText(columns = [], text = "") {
  const s = normalizeText(text);

  return columns.find((c) => {
    const h = normalizeText(c.header || c.originalHeader || c.name || "");

    const k = normalizeText(c.key || c.accessor || c.name || "");

    return s.includes(h) || h.includes(s) || s.includes(k) || k.includes(s);
  });
}

function detectOperation(message = "") {
  const s = String(message);

  if (includesAny(s, ["평균", "average", "avg"])) return "average";
  if (includesAny(s, ["합계", "총합", "sum", "total"])) return "sum";
  if (includesAny(s, ["개수", "몇 명", "몇개", "몇 개", "수"])) return "count";
  if (includesAny(s, ["최대", "최고", "max"])) return "max";
  if (includesAny(s, ["최소", "최저", "min"])) return "min";
  if (includesAny(s, ["목록", "리스트", "보여", "출력"])) return "list";

  return "list";
}

function detectMetricColumn(message = "", columns = [], operation = "") {
  if (operation === "count") return null;

  const direct = findColumnByText(columns, message);
  if (direct && ["number", "date"].includes(String(direct.type))) return direct;

  const numberCols = columns.filter((c) => String(c.type) === "number");
  if (numberCols.length === 1) return numberCols[0];

  return numberCols[0] || null;
}

function detectGroupBy(message = "", columns = []) {
  const m = String(message).match(/(.+?)별/);
  if (!m) return null;

  const hint = m[1].trim();
  return findColumnByText(columns, hint) || null;
}

function detectFilters(message = "", columns = []) {
  const filters = [];
  const s = String(message || "");

  for (const col of columns) {
    const header = String(col.header || "");
    if (!header || !s.includes(header)) continue;

    const escaped = header.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

    const numRe = new RegExp(
      `${escaped}\\s*(?:이|가)?\\s*(\\d+(?:\\.\\d+)?)\\s*(이상|초과|이하|미만)`,
    );
    const numMatch = s.match(numRe);

    if (numMatch) {
      const word = numMatch[2];
      let operator = ">=";
      if (word === "초과") operator = ">";
      if (word === "이하") operator = "<=";
      if (word === "미만") operator = "<";

      filters.push({
        columnKey: col.key,
        header: col.header,
        operator,
        value: Number(numMatch[1]),
        valueType: "number",
      });
      continue;
    }

    const textRe = new RegExp(`${escaped}\\s*(?:이|가)?\\s*([^\\s]+)`);
    const textMatch = s.match(textRe);

    if (textMatch && String(col.type) !== "number") {
      const value = textMatch[1]
        .replace(/인|인$|인\\s*$/g, "")
        .replace(/직원|학생|환자|데이터|목록|수|평균|합계/g, "")
        .trim();

      if (!value || value === "별" || value.endsWith("별")) continue;

      if (value) {
        filters.push({
          columnKey: col.key,
          header: col.header,
          operator: "=",
          value,
          valueType: "text",
        });
      }
    }
  }

  return filters;
}

function parseQueryIntent(message = "", queryTables = []) {
  const table =
    queryTables.find((t) => t.isPrimary) ||
    queryTables.sort(
      (a, b) => Number(b.confidence || 0) - Number(a.confidence || 0),
    )[0];

  if (!table) {
    return {
      ok: false,
      error: "분석 가능한 테이블이 없습니다.",
    };
  }

  const columns = table.columns || [];
  const operation = detectOperation(message);
  const metricColumn = detectMetricColumn(message, columns, operation);
  const groupBy = detectGroupBy(message, columns);
  const filters = detectFilters(message, columns);

  return {
    ok: true,
    version: "query_intent_v1",
    message,
    table: {
      tableId: table.tableId,
      tableName: table.tableName,
      sheetName: table.sheetName,
      confidence: table.confidence,
      isPrimary: !!table.isPrimary,
    },
    operation,
    metric: metricColumn
      ? {
          columnKey: metricColumn.key,
          header: metricColumn.header,
          type: metricColumn.type,
        }
      : null,
    groupBy: groupBy
      ? {
          columnKey: groupBy.key,
          header: groupBy.header,
          type: groupBy.type,
        }
      : null,
    filters,
  };
}

module.exports = { parseQueryIntent };
