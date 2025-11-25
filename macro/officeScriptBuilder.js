module.exports.buildOfficeScript = function (intent) {
  switch (intent.type) {
    case "formatRange":
      return buildFormatRange(intent);

    case "insertRow":
      return buildInsertRow(intent);

    case "createSheet":
      return buildCreateSheet(intent);

    case "duplicateSheet":
      return buildDuplicateSheet(intent);

    default:
      return fallbackScript(intent.text);
  }
};

/* ---------- Builders ---------- */

function buildFormatRange({ target, style }) {
  const { range } = target;
  const lines = [
    "function main(workbook: ExcelScript.Workbook) {",
    "  const sheet = workbook.getActiveWorksheet();",
    `  const range = sheet.getRange("${range}");`,
  ];

  if (style.bold) lines.push(`  range.getFormat().getFont().setBold(true);`);
  if (style.fillColor)
    lines.push(`  range.getFormat().getFill().setColor("${style.fillColor}");`);

  lines.push("}");
  return lines.join("\n");
}

function buildInsertRow({ target }) {
  const { rowIndex } = target;

  return `
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange("A${rowIndex}").getEntireRow().insert(ExcelScript.InsertShiftDirection.down);
}`;
}

function buildCreateSheet({ name }) {
  return `
function main(workbook: ExcelScript.Workbook) {
  workbook.addWorksheet("${name}");
}`;
}

function buildDuplicateSheet({ newSheetName }) {
  return `
function main(workbook: ExcelScript.Workbook) {
  const active = workbook.getActiveWorksheet();
  active.copy("${newSheetName}");
}`;
}

function fallbackScript(text) {
  return `// 이해할 수 없는 명령이었습니다.
// 입력: ${text}`;
}
