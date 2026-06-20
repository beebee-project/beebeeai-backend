const XLSX = require("xlsx");

const DEFAULT_XLSX_READ_OPTIONS = {
  type: "buffer",
  cellDates: true,
  cellNF: true,
  cellText: false,
};

const MOJIBAKE_PATTERN =
  /[¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ¸¼½¾±ºµ]/;

function safeDecode(buffer, encoding) {
  try {
    return new TextDecoder(encoding, { fatal: false }).decode(buffer);
  } catch (error) {
    return null;
  }
}

function workbookToHeaderText(workbook) {
  if (!workbook?.SheetNames?.length) return "";

  const chunks = [];

  for (const sheetName of workbook.SheetNames.slice(0, 3)) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      blankrows: false,
    });

    rows.slice(0, 5).forEach((row) => {
      if (!Array.isArray(row)) return;
      chunks.push(row.join(" "));
    });
  }

  return chunks.join(" ");
}

function hasMojibakeText(text = "") {
  const value = String(text || "");
  if (!value) return false;

  const suspiciousCount = (value.match(MOJIBAKE_PATTERN) || []).length;
  return suspiciousCount >= 2;
}

function readWorkbookFromBuffer(buffer, options = {}) {
  const readOptions = {
    ...DEFAULT_XLSX_READ_OPTIONS,
    ...options,
  };

  const workbook = XLSX.read(buffer, readOptions);
  const headerText = workbookToHeaderText(workbook);

  if (!hasMojibakeText(headerText)) {
    return workbook;
  }

  const decodedText =
    safeDecode(buffer, "euc-kr") ||
    safeDecode(buffer, "windows-949") ||
    safeDecode(buffer, "cp949");

  if (!decodedText) {
    return workbook;
  }

  const decodedWorkbook = XLSX.read(decodedText, {
    type: "string",
    cellDates: true,
    cellNF: true,
    cellText: false,
  });

  const decodedHeaderText = workbookToHeaderText(decodedWorkbook);

  if (hasMojibakeText(decodedHeaderText)) {
    return workbook;
  }

  return decodedWorkbook;
}

module.exports = {
  readWorkbookFromBuffer,
};
