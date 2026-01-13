function colLetterToIndex(letter) {
  if (!letter) return 0;
  let result = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i) - 64; // 'A' = 65 -> 1
    result = result * 26 + code;
  }
  return result - 1; // 0-based
}

function buildOfficeScript(intent) {
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
    case "sortRange":
      return buildSortRangeScript(intent);
    case "filterRange":
      return buildFilterRangeScript(intent);
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
    default:
      return fallbackScript(intent.text || "");
  }
}

// ─────────────────────────────
// 1) 범위 서식
// ─────────────────────────────
function buildFormatRangeScript(intent) {
  const rangeRef = (intent.target && intent.target.range) || "B:B";
  const s = intent.style || {};
  const lines = [];

  if (s.fillColor) {
    lines.push(`  range.getFormat().getFill().setColor("${s.fillColor}");`);
  }
  if (s.fontColor) {
    lines.push(`  range.getFormat().getFont().setColor("${s.fontColor}");`);
  }
  if (s.bold) {
    lines.push(`  range.getFormat().getFont().setBold(true);`);
  }
  if (s.italic) {
    lines.push(`  range.getFormat().getFont().setItalic(true);`);
  }
  if (s.underline) {
    lines.push(
      `  range.getFormat().getFont().setUnderline(ExcelScript.UnderlineStyle.single);`
    );
  }
  if (s.horizontalAlign) {
    lines.push(
      `  range.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.${s.horizontalAlign});`
    );
  }
  if (s.border) {
    lines.push(
      `  [`,
      `    ExcelScript.BorderIndex.edgeTop,`,
      `    ExcelScript.BorderIndex.edgeBottom,`,
      `    ExcelScript.BorderIndex.edgeLeft,`,
      `    ExcelScript.BorderIndex.edgeRight,`,
      `  ].forEach((edge) => {`,
      `    range.getFormat().getBorders().getItem(edge).setStyle(ExcelScript.BorderLineStyle.${s.border});`,
      `  });`
    );
  }

  if (lines.length === 0) {
    lines.push(`  // 적용할 서식이 감지되지 않았습니다.`);
  }

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange("${rangeRef}");
${lines.join("\n")}
}`;
}

// ─────────────────────────────
// 2) 값 입력
// ─────────────────────────────
function buildSetValueScript(intent) {
  const rangeRef = (intent.target && intent.target.range) || "A1";
  const value = intent.value ?? "";

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange("${rangeRef}");
  range.setValue(${JSON.stringify(value)});
}`;
}

// ─────────────────────────────
// 3) 복사
// ─────────────────────────────
function buildCopyRangeScript(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const source = sheet.getRange("${from}");
  const target = sheet.getRange("${to}");
  // source → target 복사 (값+서식 모두)
  target.copyFrom(source, ExcelScript.RangeCopyType.all, false, false);
}`;
}

// ─────────────────────────────
// 4) 지우기
// ─────────────────────────────
function buildClearRangeScript(intent) {
  const rangeRef = (intent.target && intent.target.range) || "A1:A10";

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange("${rangeRef}");
  // 값 + 서식 전체 삭제
  range.clear(ExcelScript.ClearApplyTo.all);
}`;
}

// ─────────────────────────────
// 5) 이동
// ─────────────────────────────
function buildMoveRangeScript(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const source = sheet.getRange("${from}");
  const dest = sheet.getRange("${to}");
  // 범위 전체 이동 (값+서식+수식)
  source.moveTo(dest);
}`;
}

// ─────────────────────────────
// 6) 행 삽입 / 삭제
// ─────────────────────────────
function buildInsertRowScript(intent) {
  const rowIndex = intent.rowIndex || 1;

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange("A${rowIndex}")
    .getEntireRow()
    .insert(ExcelScript.InsertShiftDirection.down);
}`;
}

function buildDeleteRowScript(intent) {
  const rowIndex = intent.rowIndex || 1;

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange("A${rowIndex}")
    .getEntireRow()
    .delete(ExcelScript.DeleteShiftDirection.up);
}`;
}

// ─────────────────────────────
// 7) 열 삽입 / 삭제
// ─────────────────────────────
function buildInsertColumnScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  const position = intent.position === "left" ? "left" : "right";

  if (col.letter) {
    const ref = `${col.letter}:${col.letter}`;
    const shiftDir = position === "left" ? "left" : "right";

    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange("${ref}");
  range.getEntireColumn()
    .insert(ExcelScript.InsertShiftDirection.${shiftDir});
}`;
  } else {
    const index = col.index || 1;
    const colIndex = index - 1;
    const shiftDir = position === "left" ? "left" : "right";

    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const totalRows = sheet.getRowCount();
  const colRange = sheet.getRangeByIndexes(0, ${colIndex}, totalRows, 1);
  colRange.getEntireColumn()
    .insert(ExcelScript.InsertShiftDirection.${shiftDir});
}`;
  }
}

// ─────────────────────────────
// 정렬 (sortRange)
// ─────────────────────────────
function buildSortRangeScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  let colIndex = 0;

  if (col.letter) {
    colIndex = colLetterToIndex(col.letter);
  } else if (col.index) {
    colIndex = Math.max(0, (col.index || 1) - 1);
  }

  const ascending = intent.direction !== "descending";

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return;
  }

  const sort = usedRange.getSort();
  sort.apply([
    {
      key: ${colIndex},
      ascending: ${ascending ? "true" : "false"}
    }
  ]);
}`;
}

// ─────────────────────────────
// 필터 (filterRange)
// ─────────────────────────────
function buildFilterRangeScript(intent) {
  const col = intent.column || { letter: null, index: 1 };
  let colIndex = 0;

  if (col.letter) {
    colIndex = colLetterToIndex(col.letter);
  } else if (col.index) {
    colIndex = Math.max(0, (col.index || 1) - 1);
  }

  const criteria = intent.criteria || "";

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return;
  }

  const autoFilter = sheet.getAutoFilter();
  const filterCriteria: ExcelScript.FilterCriteria = {
    filterOn: ExcelScript.FilterOn.custom,
    criterion1: ${JSON.stringify(criteria)}
  };

  autoFilter.apply(usedRange, ${colIndex}, filterCriteria);
}`;
}

function buildDeleteColumnScript(intent) {
  const col = intent.column || { letter: null, index: 1 };

  if (col.letter) {
    const ref = `${col.letter}:${col.letter}`;
    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange("${ref}");
  range.getEntireColumn()
    .delete(ExcelScript.DeleteShiftDirection.left);
}`;
  } else {
    const index = col.index || 1;
    const colIndex = index - 1;

    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const totalRows = sheet.getRowCount();
  const colRange = sheet.getRangeByIndexes(0, ${colIndex}, totalRows, 1);
  colRange.getEntireColumn()
    .delete(ExcelScript.DeleteShiftDirection.left);
}`;
  }
}

// ─────────────────────────────
// 8) 시트 생성 / 복사
// ─────────────────────────────
function buildCreateSheetScript(intent) {
  const name = intent.name || "NewSheet";

  return `function main(workbook: ExcelScript.Workbook) {
  workbook.addWorksheet("${name}");
}`;
}

function buildDuplicateSheetScript(intent) {
  const name = intent.name || "Backup";

  return `function main(workbook: ExcelScript.Workbook) {
  const active = workbook.getActiveWorksheet();
  active.copy("${name}");
}`;
}

// ─────────────────────────────
// 9) 시트 이름 변경 / 삭제 / 이동
// ─────────────────────────────
function buildRenameSheetScript(intent) {
  const fromName = intent.fromName || null;
  const toName = intent.toName || "RenamedSheet";

  if (fromName) {
    // "데이터" → "원본"
    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("${fromName}");
  if (sheet) {
    sheet.setName("${toName}");
  }
}`;
  } else {
    // 현재 시트 이름 변경
    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.setName("${toName}");
}`;
  }
}

function buildDeleteSheetScript(intent) {
  const name = intent.name || null;

  if (name) {
    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("${name}");
  if (sheet) {
    sheet.delete();
  }
}`;
  } else {
    // 이름이 없으면 현재 시트를 삭제 (조심!)
    return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.delete();
}`;
  }
}

function buildActivateSheetScript(intent) {
  const name = intent.name || "Sheet1";

  return `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("${name}");
  if (sheet) {
    sheet.activate();
  }
}`;
}

// ─────────────────────────────
// 10) fallback – 지원 안 되는 명령
// ─────────────────────────────
function fallbackScript(originalText) {
  const safe = (originalText || "").replace(/[\r\n]/g, " ");
  return `function main(workbook: ExcelScript.Workbook) {
  // 지원하지 않는 작업입니다.
  // 입력: ${safe}
}`;
}

module.exports = {
  buildOfficeScript,
};
