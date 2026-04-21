function toAppsScriptFunctionName(intent = {}) {
  const map = {
    groupByAggregate: "groupByAggregateMacro",
    formatRange: "formatRangeMacro",
    setValue: "setValueMacro",
    copyRange: "copyRangeMacro",
    clearRange: "clearRangeMacro",
    moveRange: "moveRangeMacro",
    removeDuplicates: "removeDuplicatesMacro",
    sortRange: "sortRangeMacro",
    filterRange: "filterRangeMacro",
    insertRow: "insertRowMacro",
    deleteRow: "deleteRowMacro",
    insertColumn: "insertColumnMacro",
    deleteColumn: "deleteColumnMacro",
    createSheet: "createSheetMacro",
    duplicateSheet: "duplicateSheetMacro",
    renameSheet: "renameSheetMacro",
    deleteSheet: "deleteSheetMacro",
    activateSheet: "activateSheetMacro",
  };
  return map[intent?.type] || "runMacro";
}

function escapeJsString(value = "") {
  return JSON.stringify(String(value ?? "")).slice(1, -1);
}

function colLetterToPosition(letter) {
  if (!letter) return 1;
  let result = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i) - 64; // 'A' = 65 -> 1
    result = result * 26 + code;
  }
  return result; // 1-based
}

function getAppsScriptRangeExpr(intent) {
  const rangeRef = intent?.target?.range || null;
  if (!rangeRef || rangeRef === "__USED_RANGE__") {
    return "sheet.getDataRange()";
  }
  return `sheet.getRange("${rangeRef}")`;
}

function getColumnPosition(col) {
  if (!col) return 1;
  if (col.letter) return colLetterToPosition(col.letter);
  if (col.index) return col.index;
  return 1;
}

function buildAppsScript(intent) {
  if (!intent || !intent.type) {
    return fallbackScript("잘못된 intent");
  }

  switch (intent.type) {
    case "groupByAggregate":
      return buildGroupByAggregateScript(intent);
    case "formatRange":
      return buildFormatRangeScript(intent);
    case "setValue":
      return buildSetValueScript(intent);
    case "copyRange":
      return buildCopyRangeScript(intent);
    case "clearRange":
      return buildClearRangeScript(intent);
    case "moveRange":
      return buildMoveRangeScript(intent);
    case "removeDuplicates":
      return buildRemoveDuplicatesScript(intent);
    case "insertRow":
      return buildInsertRowScript(intent);
    case "deleteRow":
      return buildDeleteRowScript(intent);
    case "insertColumn":
      return buildInsertColumnScript(intent);
    case "deleteColumn":
      return buildDeleteColumnScript(intent);
    case "createSheet":
      return buildCreateSheetScript(intent);
    case "duplicateSheet":
      return buildDuplicateSheetScript(intent);
    case "renameSheet":
      return buildRenameSheetScript(intent);
    case "deleteSheet":
      return buildDeleteSheetScript(intent);
    case "activateSheet":
      return buildActivateSheetScript(intent);
    case "sortRange":
      return buildSortRangeScript(intent);
    case "filterRange":
      return buildFilterRangeScript(intent);
    default:
      return fallbackScript(intent.text || "");
  }
}

// ─────────────────────────────
// 0) 그룹 집계 (groupByAggregate)
// ─────────────────────────────
function buildGroupByAggregateScript(intent) {
  const fnName = toAppsScriptFunctionName(intent);
  const rangeExpr =
    intent.aggregateType === "count"
      ? getAppsScriptRangeExpr(intent)
      : "sheet.getDataRange()";
  const groupColPos = getColumnPosition(intent.groupByColumn);
  const valueColPos =
    intent.aggregateType === "count"
      ? null
      : getColumnPosition(intent.valueColumn || { index: 2 });
  const aggregateType = intent.aggregateType || "count";

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = ${rangeExpr};
  const values = range.getValues();
  const summary = {};

  values.forEach((row) => {
    const key = row[${groupColPos - 1}];
    if (key === "" || key == null) return;

    if (!summary[key]) {
      summary[key] = { count: 0, sum: 0 };
    }

    summary[key].count += 1;

    ${
      aggregateType === "count"
        ? ""
        : `const num = Number(row[${valueColPos - 1}]);
    if (!Number.isNaN(num)) {
      summary[key].sum += num;
    }`
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = ss.insertSheet("요약표");
  const output = [["기준값", "${aggregateType}"]];

  Object.keys(summary).forEach((key) => {
    let result = summary[key].count;
    if ("${aggregateType}" === "sum") {
      result = summary[key].sum;
    } else if ("${aggregateType}" === "average") {
      result = summary[key].count ? summary[key].sum / summary[key].count : 0;
    }
    output.push([key, result]);
  });

  target.getRange(1, 1, output.length, output[0].length).setValues(output);
}`;
}

// ─────────────────────────────
// 1) 범위 서식 (배경/글씨색/굵게/이탤릭/밑줄/정렬/테두리)
// intent: { type: "formatRange", target: { range }, style: { ... } }
// ─────────────────────────────
function buildFormatRangeScript(intent) {
  const rangeRef = (intent.target && intent.target.range) || "B:B";
  const s = intent.style || {};
  const lines = [];
  const fnName = toAppsScriptFunctionName(intent);

  // 배경색
  if (s.fillColor) {
    lines.push(`  range.setBackground("${s.fillColor}");`);
  }
  // 글씨색
  if (s.fontColor) {
    lines.push(`  range.setFontColor("${s.fontColor}");`);
  }
  // 굵게
  if (s.bold) {
    lines.push(`  range.setFontWeight("bold");`);
  }
  // 이탤릭
  if (s.italic) {
    lines.push(`  range.setFontStyle("italic");`);
  }
  // 밑줄
  if (s.underline) {
    lines.push(`  range.setFontLine("underline");`);
  }
  // 정렬
  if (s.horizontalAlign) {
    const align = s.horizontalAlign.toLowerCase(); // Center -> center
    lines.push(`  range.setHorizontalAlignment("${align}");`);
  }
  // 테두리
  if (s.border) {
    // 얇은 실선 테두리
    lines.push(
      `  range.setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);`,
    );
  }

  if (lines.length === 0) {
    lines.push(`  // 적용할 서식이 감지되지 않았습니다.`);
  }

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("${rangeRef}");
${lines.join("\n")}
}`;
}

// ─────────────────────────────
// 2) 값 입력 (setValue)
// intent: { type: "setValue", target: { range }, value }
// ─────────────────────────────
function buildSetValueScript(intent) {
  const rangeRef = (intent.target && intent.target.range) || "A1";
  const value = intent.value ?? "";
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("${rangeRef}");
  range.setValue(${JSON.stringify(value)});
}`;
}

// ─────────────────────────────
// 3) 복사 (copyRange)
// intent: { type: "copyRange", from, to }
// ─────────────────────────────
function buildCopyRangeScript(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const source = sheet.getRange("${from}");
  const target = sheet.getRange("${to}");
  // 값 + 서식까지 복사
  source.copyTo(target, { contentsOnly: false });
}`;
}

// ─────────────────────────────
// 4) 지우기 (clearRange)
// intent: { type: "clearRange", target: { range } }
// ─────────────────────────────
function buildClearRangeScript(intent) {
  const rangeRef = (intent.target && intent.target.range) || "A1:A10";
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("${rangeRef}");
  // 값 + 서식 전체 삭제
  range.clear();
}`;
}

// ─────────────────────────────
// 5) 이동 (moveRange)
// intent: { type: "moveRange", from, to }
// ─────────────────────────────
function buildMoveRangeScript(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const source = sheet.getRange("${from}");
  const dest = sheet.getRange("${to}");
  // 값+서식 복사 후 원본 지우기 = 이동
  source.copyTo(dest, { contentsOnly: false });
  source.clear();
}`;
}

// ─────────────────────────────
// 5-1) 중복 제거 (removeDuplicates)
// intent: { type: "removeDuplicates", target?: { range }, column?: { letter, index } }
// ─────────────────────────────
function buildRemoveDuplicatesScript(intent) {
  const fnName = toAppsScriptFunctionName(intent);
  const col = intent.column || { letter: null, index: 1 };

  let colPos = 1;
  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  const rangeExpr = getAppsScriptRangeExpr(intent);

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = ${rangeExpr};
  range.removeDuplicates([${colPos}]);
}`;
}

// ─────────────────────────────
// 6) 행 삽입 / 삭제
// insertRow: { type: "insertRow", rowIndex }
// deleteRow: { type: "deleteRow", rowIndex }
// ─────────────────────────────
function buildInsertRowScript(intent) {
  const rowIndex = intent.rowIndex || 1;
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRows(${rowIndex}, 1);
}`;
}

function buildDeleteRowScript(intent) {
  const rowIndex = intent.rowIndex || 1;
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.deleteRow(${rowIndex});
}`;
}

// ─────────────────────────────
// 7) 열 삽입 / 삭제
// insertColumn: { type: "insertColumn", column: { letter, index }, position }
// deleteColumn: { type: "deleteColumn", column: { letter, index } }
// ─────────────────────────────
function buildInsertColumnScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  const position = intent.position === "left" ? "left" : "right";
  const fnName = toAppsScriptFunctionName(intent);

  let colPos = 1;
  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  if (position === "left") {
    return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertColumnBefore(${colPos});
}`;
  } else {
    return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertColumnAfter(${colPos});
}`;
  }
}

function buildDeleteColumnScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  const fnName = toAppsScriptFunctionName(intent);

  let colPos = 1;
  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.deleteColumn(${colPos});
}`;
}

// ─────────────────────────────
// 8) 시트 생성 / 복사
// createSheet: { type: "createSheet", name }
// duplicateSheet: { type: "duplicateSheet", name }
// ─────────────────────────────
function buildCreateSheetScript(intent) {
  const name = escapeJsString(intent.name || "NewSheet");
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet("${name}");
}`;
}

function buildDuplicateSheetScript(intent) {
  const name = escapeJsString(intent.name || "Backup");
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const active = ss.getActiveSheet();
  const copied = active.copyTo(ss);
  copied.setName("${name}");
}`;
}

// ─────────────────────────────
// 9) 시트 이름 변경 / 삭제 / 이동
// renameSheet: { type: "renameSheet", fromName, toName }
// deleteSheet: { type: "deleteSheet", name }
// activateSheet: { type: "activateSheet", name }
// ─────────────────────────────
function buildRenameSheetScript(intent) {
  const fromName = intent.fromName ? escapeJsString(intent.fromName) : null;
  const toName = escapeJsString(intent.toName || "RenamedSheet");
  const fnName = toAppsScriptFunctionName(intent);

  if (fromName) {
    // 특정 시트 이름 변경
    return `function ${fnName}() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("${fromName}");
  if (sheet) {
    sheet.setName("${toName}");
  }
}`;
  } else {
    // 현재 시트 이름 변경
    return `function ${fnName}() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.setName("${toName}");
}`;
  }
}

function buildDeleteSheetScript(intent) {
  const name = intent.name ? escapeJsString(intent.name) : null;
  const fnName = toAppsScriptFunctionName(intent);

  if (name) {
    return `function ${fnName}() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("${name}");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
}`;
  } else {
    return `function ${fnName}() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  ss.deleteSheet(sheet);
}`;
  }
}

function buildActivateSheetScript(intent) {
  const name = escapeJsString(intent.name || "Sheet1");
  const fnName = toAppsScriptFunctionName(intent);

  return `function ${fnName}() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("${name}");
  if (sheet) {
    sheet.activate();
  }
}`;
}

// ─────────────────────────────
// 10) 정렬 (sortRange)
// intent: { type: "sortRange", column: { letter, index }, direction }
// ─────────────────────────────
function buildSortRangeScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  const fnName = toAppsScriptFunctionName(intent);
  let colPos = 1;

  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  const ascending = intent.direction !== "descending";

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = ${getAppsScriptRangeExpr(intent)};
  range.sort({ column: ${colPos}, ascending: ${ascending ? "true" : "false"} });
}`;
}

// ─────────────────────────────
// 11) 필터 (filterRange)
// intent: { type: "filterRange", column: { letter, index }, criteria }
// ─────────────────────────────
function buildFilterRangeScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  const fnName = toAppsScriptFunctionName(intent);
  let colPos = 1;

  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  const criteria = intent.criteria || "";

  return `function ${fnName}() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = ${getAppsScriptRangeExpr(intent)};

  let filter = range.getFilter();
  if (!filter) {
    filter = range.createFilter();
  }

  const crit = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(${JSON.stringify(criteria)})
    .build();

  filter.setColumnFilterCriteria(${colPos}, crit);
}`;
}

// ─────────────────────────────
// 12) fallback – 지원 안 되는 명령
// ─────────────────────────────
function fallbackScript(originalText) {
  const safe = (originalText || "").replace(/[\r\n]/g, " ");
  return `function runMacro() {
  // 지원하지 않는 작업입니다.
  // 입력: ${safe}
}`;
}

module.exports = {
  buildAppsScript,
};
