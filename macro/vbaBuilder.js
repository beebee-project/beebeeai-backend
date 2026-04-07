function toMacroProcedureName(intent = {}) {
  const map = {
    formatRange: "FormatRangeMacro",
    setValue: "SetValueMacro",
    copyRange: "CopyRangeMacro",
    clearRange: "ClearRangeMacro",
    moveRange: "MoveRangeMacro",
    sortRange: "SortRangeMacro",
    filterRange: "FilterRangeMacro",
    insertRow: "InsertRowMacro",
    deleteRow: "DeleteRowMacro",
    insertColumn: "InsertColumnMacro",
    deleteColumn: "DeleteColumnMacro",
    createSheet: "CreateSheetMacro",
    duplicateSheet: "DuplicateSheetMacro",
    renameSheet: "RenameSheetMacro",
    deleteSheet: "DeleteSheetMacro",
    activateSheet: "ActivateSheetMacro",
  };
  return map[intent?.type] || "RunMacro";
}

function escapeVbaString(value = "") {
  return String(value).replace(/"/g, '""');
}

function buildVbaScript(intent) {
  if (!intent || !intent.type) {
    return fallbackVba("잘못된 intent");
  }

  switch (intent.type) {
    case "formatRange":
      return buildFormatRangeVba(intent);
    case "setValue":
      return buildSetValueVba(intent);
    case "copyRange":
      return buildCopyRangeVba(intent);
    case "clearRange":
      return buildClearRangeVba(intent);
    case "moveRange":
      return buildMoveRangeVba(intent);
    case "sortRange":
      return buildSortRangeVba(intent);
    case "filterRange":
      return buildFilterRangeVba(intent);
    case "insertRow":
      return buildInsertRowVba(intent);
    case "deleteRow":
      return buildDeleteRowVba(intent);
    case "insertColumn":
      return buildInsertColumnVba(intent);
    case "deleteColumn":
      return buildDeleteColumnVba(intent);
    case "createSheet":
      return buildCreateSheetVba(intent);
    case "duplicateSheet":
      return buildDuplicateSheetVba(intent);
    case "renameSheet":
      return buildRenameSheetVba(intent);
    case "deleteSheet":
      return buildDeleteSheetVba(intent);
    case "activateSheet":
      return buildActivateSheetVba(intent);
    default:
      return fallbackVba(intent.text || "");
  }
}

/* =========================
 * 공통 헬퍼
 * =======================*/
function getRangeRef(intent, fallback = "A1") {
  return (intent?.target && intent.target.range) || fallback;
}

function getColumnLetterOrIndex(col) {
  if (!col) return { letter: null, index: 1 };
  if (col.letter)
    return { letter: String(col.letter).toUpperCase(), index: null };
  if (col.index) return { letter: null, index: Number(col.index) };
  return { letter: null, index: 1 };
}

function sortOrderToVba(direction = "ascending") {
  return direction === "descending" ? "xlDescending" : "xlAscending";
}

/* =========================
 * 1) 서식
 * =======================*/
function buildFormatRangeVba(intent) {
  const rangeRef = getRangeRef(intent, "B:B");
  const s = intent.style || {};
  const lines = [];

  if (s.fillColor) {
    lines.push(
      `    Range("${rangeRef}").Interior.Color = RGB(255, 255, 0) ' TODO: ${escapeVbaString(
        s.fillColor,
      )}`,
    );
  }
  if (s.fontColor) {
    lines.push(
      `    Range("${rangeRef}").Font.Color = RGB(0, 0, 0) ' TODO: ${escapeVbaString(
        s.fontColor,
      )}`,
    );
  }
  if (s.bold) {
    lines.push(`    Range("${rangeRef}").Font.Bold = True`);
  }
  if (s.italic) {
    lines.push(`    Range("${rangeRef}").Font.Italic = True`);
  }
  if (s.underline) {
    lines.push(
      `    Range("${rangeRef}").Font.Underline = xlUnderlineStyleSingle`,
    );
  }
  if (s.horizontalAlign) {
    const map = {
      Left: "xlLeft",
      Center: "xlCenter",
      Right: "xlRight",
    };
    lines.push(
      `    Range("${rangeRef}").HorizontalAlignment = ${map[s.horizontalAlign] || "xlGeneral"}`,
    );
  }
  if (s.border) {
    lines.push(`    Range("${rangeRef}").Borders.LineStyle = xlContinuous`);
  }

  if (!lines.length) {
    lines.push(`    ' 적용할 서식이 감지되지 않았습니다.`);
  }

  return `Sub Main()
${lines.join("\n")}
End Sub`;
}

/* =========================
 * 2) 값 입력
 * =======================*/
function buildSetValueVba(intent) {
  const rangeRef = getRangeRef(intent, "A1");
  const value = escapeVbaString(intent.value ?? "");

  return `Sub Main()
    Range("${rangeRef}").Value = "${value}"
End Sub`;
}

/* =========================
 * 3) 복사
 * =======================*/
function buildCopyRangeVba(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";

  return `Sub Main()
    Range("${from}").Copy Destination:=Range("${to}")
End Sub`;
}

/* =========================
 * 4) 지우기
 * =======================*/
function buildClearRangeVba(intent) {
  const rangeRef = getRangeRef(intent, "A1:A10");

  return `Sub Main()
    Range("${rangeRef}").Clear
End Sub`;
}

/* =========================
 * 5) 이동
 * =======================*/
function buildMoveRangeVba(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";

  return `Sub Main()
    Range("${from}").Cut Destination:=Range("${to}")
End Sub`;
}

/* =========================
 * 6) 정렬
 * =======================*/
function buildSortRangeVba(intent) {
  const col = getColumnLetterOrIndex(intent.column);
  const order = sortOrderToVba(intent.direction || "ascending");

  let keyRef = 'Range("A:A")';
  if (col.letter) {
    keyRef = `Range("${col.letter}:${col.letter}")`;
  } else if (col.index) {
    keyRef = `Columns(${col.index})`;
  }

  return `Sub Main()
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=${keyRef}, SortOn:=xlSortOnValues, Order:=${order}, DataOption:=xlSortNormal
        .SetRange ActiveSheet.UsedRange
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub`;
}

/* =========================
 * 7) 필터
 * =======================*/
function buildFilterRangeVba(intent) {
  const col = getColumnLetterOrIndex(intent.column);
  const criteria = escapeVbaString(intent.criteria || "");

  const fieldExpr =
    col.index || (col.letter ? `Range("${col.letter}1").Column` : 1);

  return `Sub Main()
    ActiveSheet.UsedRange.AutoFilter Field:=${fieldExpr}, Criteria1:="${criteria}"
End Sub`;
}

/* =========================
 * 8) 행 삽입/삭제
 * =======================*/
function buildInsertRowVba(intent) {
  const rowIndex = Number(intent.rowIndex || 1);

  return `Sub Main()
    Rows(${rowIndex}).Insert Shift:=xlDown
End Sub`;
}

function buildDeleteRowVba(intent) {
  const rowIndex = Number(intent.rowIndex || 1);
  const procName = toMacroProcedureName(intent);

  return `Sub ${procName}()
    Rows(${rowIndex}).Delete
End Sub`;
}

/* =========================
 * 9) 열 삽입/삭제
 * =======================*/
function buildInsertColumnVba(intent) {
  const col = getColumnLetterOrIndex(intent.column);

  if (col.letter) {
    return `Sub Main()
    Columns("${col.letter}:${col.letter}").Insert Shift:=xlToRight
End Sub`;
  }

  return `Sub Main()
    Columns(${col.index || 1}).Insert Shift:=xlToRight
End Sub`;
}

function buildDeleteColumnVba(intent) {
  const col = getColumnLetterOrIndex(intent.column);

  if (col.letter) {
    return `Sub Main()
    Columns("${col.letter}:${col.letter}").Delete
End Sub`;
  }

  return `Sub Main()
    Columns(${col.index || 1}).Delete
End Sub`;
}

/* =========================
 * 10) 시트
 * =======================*/
function buildCreateSheetVba(intent) {
  const name = escapeVbaString(intent.name || "NewSheet");
  const procName = toMacroProcedureName(intent);

  return `Sub ${procName}()
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "${name}"
End Sub`;
}

function buildDuplicateSheetVba(intent) {
  const name = escapeVbaString(intent.name || "Backup");

  return `Sub Main()
    ActiveSheet.Copy After:=ActiveSheet
    ActiveSheet.Name = "${name}"
End Sub`;
}

function buildRenameSheetVba(intent) {
  const fromName = intent.fromName ? escapeVbaString(intent.fromName) : null;
  const toName = escapeVbaString(intent.toName || "RenamedSheet");

  if (fromName) {
    return `Sub Main()
    Worksheets("${fromName}").Name = "${toName}"
End Sub`;
  }

  return `Sub Main()
    ActiveSheet.Name = "${toName}"
End Sub`;
}

function buildDeleteSheetVba(intent) {
  const name = intent.name ? escapeVbaString(intent.name) : null;

  if (name) {
    return `Sub Main()
    Application.DisplayAlerts = False
    Worksheets("${name}").Delete
    Application.DisplayAlerts = True
End Sub`;
  }

  return `Sub Main()
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End Sub`;
}

function buildActivateSheetVba(intent) {
  const name = escapeVbaString(intent.name || "Sheet1");

  return `Sub Main()
    Worksheets("${name}").Activate
End Sub`;
}

/* =========================
 * fallback
 * =======================*/
function fallbackVba(originalText) {
  const safe = escapeVbaString(originalText || "");
  return `Sub RunMacro()
    ' 지원하지 않는 작업입니다.
    ' 입력: ${safe}
End Sub`;
}

module.exports = {
  buildVbaScript,
};
