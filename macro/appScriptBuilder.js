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

function buildAppsScript(intent) {
  if (!intent || !intent.type) {
    return fallbackScript("잘못된 intent");
  }

  switch (intent.type) {
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
// 1) 범위 서식 (배경/글씨색/굵게/이탤릭/밑줄/정렬/테두리)
// intent: { type: "formatRange", target: { range }, style: { ... } }
// ─────────────────────────────
function buildFormatRangeScript(intent) {
  const rangeRef = (intent.target && intent.target.range) || "B:B";
  const s = intent.style || {};
  const lines = [];

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
    // 필요 시 TextStyle 기반으로 확장
    lines.push(`  // underline은 현재 버전에서 미지원(추후 보강)`);
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
      `  range.setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);`
    );
  }

  if (lines.length === 0) {
    lines.push(`  // 적용할 서식이 감지되지 않았습니다.`);
  }

  return `function main() {
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

  return `function main() {
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

  return `function main() {
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

  return `function main() {
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

  return `function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const source = sheet.getRange("${from}");
  const dest = sheet.getRange("${to}");
  // 값+서식 복사 후 원본 지우기 = 이동
  source.copyTo(dest, { contentsOnly: false });
  source.clear();
}`;
}

// ─────────────────────────────
// 6) 행 삽입 / 삭제
// insertRow: { type: "insertRow", rowIndex }
// deleteRow: { type: "deleteRow", rowIndex }
// ─────────────────────────────
function buildInsertRowScript(intent) {
  const rowIndex = intent.rowIndex || 1;

  return `function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRows(${rowIndex}, 1);
}`;
}

function buildDeleteRowScript(intent) {
  const rowIndex = intent.rowIndex || 1;

  return `function main() {
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

  let colPos = 1;
  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  if (position === "left") {
    return `function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertColumnBefore(${colPos});
}`;
  } else {
    return `function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertColumnAfter(${colPos});
}`;
  }
}

function buildDeleteColumnScript(intent) {
  const col = intent.column || { letter: null, index: 1 };

  let colPos = 1;
  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  return `function main() {
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
  const name = intent.name || "NewSheet";

  return `function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet("${name}");
}`;
}

function buildDuplicateSheetScript(intent) {
  const name = intent.name || "Backup";

  return `function main() {
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
  const fromName = intent.fromName || null;
  const toName = intent.toName || "RenamedSheet";

  if (fromName) {
    // 특정 시트 이름 변경
    return `function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("${fromName}");
  if (sheet) {
    sheet.setName("${toName}");
  }
}`;
  } else {
    // 현재 시트 이름 변경
    return `function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.setName("${toName}");
}`;
  }
}

function buildDeleteSheetScript(intent) {
  const name = intent.name || null;

  if (name) {
    return `function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("${name}");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
}`;
  } else {
    return `function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  ss.deleteSheet(sheet);
}`;
  }
}

function buildActivateSheetScript(intent) {
  const name = intent.name || "Sheet1";

  return `function main() {
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
  let colPos = 1;

  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  const ascending = intent.direction !== "descending";

  return `function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  range.sort({ column: ${colPos}, ascending: ${ascending ? "true" : "false"} });
}`;
}

// ─────────────────────────────
// 11) 필터 (filterRange)
// intent: { type: "filterRange", column: { letter, index }, criteria }
// ─────────────────────────────
function buildFilterRangeScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  let colPos = 1;

  if (col.letter) colPos = colLetterToPosition(col.letter);
  else if (col.index) colPos = col.index;

  const criteria = intent.criteria || "";

  return `function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();

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
  return `function main() {
  // 지원하지 않는 작업입니다.
  // 입력: ${safe}
}`;
}

module.exports = {
  buildAppsScript,
};
