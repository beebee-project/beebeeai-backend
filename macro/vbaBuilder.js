function toMacroProcedureName(intent = {}) {
  const map = {
    groupByAggregate: "GroupByAggregateMacro",
    formatRange: "FormatRangeMacro",
    setValue: "SetValueMacro",
    copyRange: "CopyRangeMacro",
    clearRange: "ClearRangeMacro",
    moveRange: "MoveRangeMacro",
    removeDuplicates: "RemoveDuplicatesMacro",
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

function hexToRgbTuple(hex) {
  const normalized = String(hex || "")
    .trim()
    .replace(/^#/, "");
  if (!/^[0-9A-Fa-f]{6}$/.test(normalized)) return null;
  const r = parseInt(normalized.slice(0, 2), 16);
  const g = parseInt(normalized.slice(2, 4), 16);
  const b = parseInt(normalized.slice(4, 6), 16);
  return [r, g, b];
}

function colorToVbaRgb(hex, fallback = "RGB(255, 255, 0)") {
  const rgb = hexToRgbTuple(hex);
  if (!rgb) return fallback;
  return `RGB(${rgb[0]}, ${rgb[1]}, ${rgb[2]})`;
}

function buildVbaScript(intent) {
  if (!intent || !intent.type) {
    return fallbackVba("잘못된 intent");
  }

  switch (intent.type) {
    case "groupByAggregate":
      return buildGroupByAggregateVba(intent);
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
    case "removeDuplicates":
      return buildRemoveDuplicatesVba(intent);
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

function getRemoveDuplicatesRangeRef(intent) {
  return (intent?.target && intent.target.range) || "ActiveSheet.UsedRange";
}

function getVbaRangeExpr(intent, fallback = "ActiveSheet.UsedRange") {
  const rangeRef = intent?.target?.range || null;
  if (!rangeRef || rangeRef === "__USED_RANGE__") return fallback;
  return `Range("${rangeRef}")`;
}

function getVbaHeaderConst(intent) {
  if (intent?.hasHeader === true) return "xlYes";
  if (intent?.hasHeader === false) return "xlNo";
  return "xlGuess";
}

function getColumnNumberExpr(col) {
  const c = getColumnLetterOrIndex(col);
  if (c.index) return String(c.index);
  if (c.letter) return `Range("${c.letter}1").Column`;
  return "1";
}

/* =========================
 * 0) 그룹 집계
 * =======================*/
function buildGroupByAggregateVba(intent) {
  const procName = toMacroProcedureName(intent);
  const groupColExpr = getColumnNumberExpr(intent.groupByColumn);
  const valueColExpr =
    intent.aggregateType === "count"
      ? null
      : getColumnNumberExpr(intent.valueColumn || { index: 2 });

  const rangeExpr =
    intent.aggregateType === "count"
      ? getVbaRangeExpr(intent)
      : "ActiveSheet.UsedRange";
  const aggType = intent.aggregateType || "count";

  let formulaLine = `        ws.Cells(outRow, 2).Value = Application.WorksheetFunction.CountIf(${rangeExpr}.Columns(${groupColExpr}), key)`;
  if (aggType === "sum") {
    formulaLine = `        ws.Cells(outRow, 2).Value = Application.WorksheetFunction.SumIf(${rangeExpr}.Columns(${groupColExpr}), key, ${rangeExpr}.Columns(${valueColExpr}))`;
  } else if (aggType === "average") {
    formulaLine = `        ws.Cells(outRow, 2).Value = Application.WorksheetFunction.AverageIf(${rangeExpr}.Columns(${groupColExpr}), key, ${rangeExpr}.Columns(${valueColExpr}))`;
  }

  return `Sub ${procName}()
    Dim src As Range
    Dim ws As Worksheet
    Dim dict As Object
    Dim cell As Range
    Dim key As Variant
    Dim outRow As Long

    Set src = ${rangeExpr}
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "요약표"
    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In src.Columns(${groupColExpr}).Cells
        If Trim(CStr(cell.Value)) <> "" Then
            If Not dict.Exists(CStr(cell.Value)) Then
                dict.Add CStr(cell.Value), True
            End If
        End If
    Next cell

    ws.Cells(1, 1).Value = "기준값"
    ws.Cells(1, 2).Value = "${aggType}"
    outRow = 2

    For Each key In dict.Keys
        ws.Cells(outRow, 1).Value = key
${formulaLine}
        outRow = outRow + 1
    Next key
End Sub`;
}

/* =========================
 * 1) 서식
 * =======================*/
function buildFormatRangeVba(intent) {
  const rangeRef = getRangeRef(intent, "B:B");
  const s = intent.style || {};
  const lines = [];
  const procName = toMacroProcedureName(intent);

  if (s.fillColor) {
    lines.push(
      `    Range("${rangeRef}").Interior.Color = ${colorToVbaRgb(s.fillColor, "RGB(255, 255, 0)")}`,
    );
  }
  if (s.fontColor) {
    lines.push(
      `    Range("${rangeRef}").Font.Color = ${colorToVbaRgb(s.fontColor, "RGB(0, 0, 0)")}`,
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

  return `Sub ${procName}()
${lines.join("\n")}
End Sub`;
}

/* =========================
 * 2) 값 입력
 * =======================*/
function buildSetValueVba(intent) {
  const rangeRef = getRangeRef(intent, "A1");
  const procName = toMacroProcedureName(intent);
  const value =
    typeof intent.value === "number"
      ? String(intent.value)
      : intent.value === "__TODAY__"
        ? "Date"
        : `"${escapeVbaString(intent.value ?? "")}"`;

  return `Sub ${procName}()
    Range("${rangeRef}").Value = ${value}
End Sub`;
}

/* =========================
 * 3) 복사
 * =======================*/
function buildCopyRangeVba(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";
  const procName = toMacroProcedureName(intent);

  return `Sub ${procName}()
    Range("${from}").Copy Destination:=Range("${to}")
End Sub`;
}

/* =========================
 * 4) 지우기
 * =======================*/
function buildClearRangeVba(intent) {
  const rangeRef = getRangeRef(intent, "A1:A10");
  const procName = toMacroProcedureName(intent);
  const targetExpr =
    rangeRef === "__USED_RANGE__"
      ? "ActiveSheet.UsedRange"
      : `Range("${rangeRef}")`;

  return `Sub ${procName}()
    ${targetExpr}.Clear
End Sub`;
}

/* =========================
 * 5) 이동
 * =======================*/
function buildMoveRangeVba(intent) {
  const from = intent.from || "A1:A1";
  const to = intent.to || "B1:B1";
  const procName = toMacroProcedureName(intent);

  return `Sub ${procName}()
    Range("${from}").Cut Destination:=Range("${to}")
End Sub`;
}

/* =========================
+ * 5-1) 중복 제거
+ * =======================*/
function buildRemoveDuplicatesVba(intent) {
  const procName = toMacroProcedureName(intent);
  const rangeRef = getRemoveDuplicatesRangeRef(intent);
  const col = getColumnLetterOrIndex(intent.column);
  const headerConst = getVbaHeaderConst(intent);

  let columnsExpr = "Array(1)";
  if (col.index) {
    columnsExpr = `Array(${col.index})`;
  } else if (col.letter) {
    columnsExpr = `Array(Range("${col.letter}1").Column)`;
  }

  const targetExpr =
    rangeRef === "ActiveSheet.UsedRange" || rangeRef === "__USED_RANGE__"
      ? "ActiveSheet.UsedRange"
      : `Range("${rangeRef}")`;

  return `Sub ${procName}()
    ${targetExpr}.RemoveDuplicates Columns:=${columnsExpr}, Header:=${headerConst}
End Sub`;
}

/* =========================
 * 6) 정렬
 * =======================*/
function buildSortRangeVba(intent) {
  const col = getColumnLetterOrIndex(intent.column);
  const order = sortOrderToVba(intent.direction || "ascending");
  const procName = toMacroProcedureName(intent);
  const rangeExpr = getVbaRangeExpr(intent);
  const headerConst = getVbaHeaderConst(intent);

  let keyRef = 'Range("A:A")';
  if (col.letter) {
    keyRef = `Range("${col.letter}:${col.letter}")`;
  } else if (col.index) {
    keyRef = `Columns(${col.index})`;
  }

  return `Sub ${procName}()
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=${keyRef}, SortOn:=xlSortOnValues, Order:=${order}, DataOption:=xlSortNormal
        .SetRange ${rangeExpr}
        .Header = ${headerConst}
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
  const procName = toMacroProcedureName(intent);
  const rangeExpr = getVbaRangeExpr(intent);

  const fieldExpr =
    col.index || (col.letter ? `Range("${col.letter}1").Column` : 1);

  return `Sub ${procName}()
    ${rangeExpr}.AutoFilter Field:=${fieldExpr}, Criteria1:="${criteria}"
End Sub`;
}

/* =========================
 * 8) 행 삽입/삭제
 * =======================*/
function buildInsertRowVba(intent) {
  const rowIndex = Number(intent.rowIndex || 1);
  const procName = toMacroProcedureName(intent);

  return `Sub ${procName}()
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
  const procName = toMacroProcedureName(intent);

  if (col.letter) {
    return `Sub ${procName}()
    Columns("${col.letter}:${col.letter}").Insert Shift:=xlToRight
End Sub`;
  }

  return `Sub ${procName}()
    Columns(${col.index || 1}).Insert Shift:=xlToRight
End Sub`;
}

function buildDeleteColumnVba(intent) {
  const col = getColumnLetterOrIndex(intent.column);
  const procName = toMacroProcedureName(intent);

  if (col.letter) {
    return `Sub ${procName}()
    Columns("${col.letter}:${col.letter}").Delete
End Sub`;
  }

  return `Sub ${procName}()
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
  const procName = toMacroProcedureName(intent);

  return `Sub ${procName}()
    ActiveSheet.Copy After:=ActiveSheet
    ActiveSheet.Name = "${name}"
End Sub`;
}

function buildRenameSheetVba(intent) {
  const fromName = intent.fromName ? escapeVbaString(intent.fromName) : null;
  const toName = escapeVbaString(intent.toName || "RenamedSheet");
  const procName = toMacroProcedureName(intent);

  if (fromName) {
    return `Sub ${procName}()
    Worksheets("${fromName}").Name = "${toName}"
End Sub`;
  }

  return `Sub ${procName}()
    ActiveSheet.Name = "${toName}"
End Sub`;
}

function buildDeleteSheetVba(intent) {
  const name = intent.name ? escapeVbaString(intent.name) : null;
  const procName = toMacroProcedureName(intent);

  if (name) {
    return `Sub ${procName}()
    Application.DisplayAlerts = False
    Worksheets("${name}").Delete
    Application.DisplayAlerts = True
End Sub`;
  }

  return `Sub ${procName}()
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End Sub`;
}

function buildActivateSheetVba(intent) {
  const name = escapeVbaString(intent.name || "Sheet1");
  const procName = toMacroProcedureName(intent);

  return `Sub ${procName}()
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
