const XLSX = require("xlsx");
const iconv = require("iconv-lite");

const DEFAULT_XLSX_READ_OPTIONS = {
  type: "buffer",
  cellDates: true,
  cellNF: true,
  cellText: false,
};

const MOJIBAKE_PATTERN =
  /[¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ¸¼½¾±ºµ]/g;

const HANGUL_PATTERN = /[가-힣]/g;
const REPLACEMENT_PATTERN = /\uFFFD/g;

function isZipWorkbook(buffer) {
  return buffer?.[0] === 0x50 && buffer?.[1] === 0x4b;
}

function isOleWorkbook(buffer) {
  return (
    buffer?.[0] === 0xd0 &&
    buffer?.[1] === 0xcf &&
    buffer?.[2] === 0x11 &&
    buffer?.[3] === 0xe0
  );
}

function isProbablyTextBuffer(buffer) {
  if (!buffer || !buffer.length) return false;
  if (isZipWorkbook(buffer) || isOleWorkbook(buffer)) return false;

  const sample = Buffer.from(buffer).subarray(0, Math.min(buffer.length, 4096));
  let nullCount = 0;
  let delimiterCount = 0;

  for (const byte of sample) {
    if (byte === 0x00) nullCount += 1;
    if (
      byte === 0x2c || // ,
      byte === 0x09 || // tab
      byte === 0x3b || // ;
      byte === 0x0a || // \n
      byte === 0x0d // \r
    ) {
      delimiterCount += 1;
    }
  }

  const nullRatio = nullCount / sample.length;
  return nullRatio < 0.01 && delimiterCount >= 1;
}

function countMatches(text = "", pattern) {
  return (String(text || "").match(pattern) || []).length;
}

function hasMojibakeText(text = "") {
  return countMatches(text, MOJIBAKE_PATTERN) >= 2;
}

function safeDecode(buffer, encoding) {
  try {
    if (encoding === "utf8" || encoding === "utf-8") {
      return Buffer.from(buffer).toString("utf8");
    }

    return iconv.decode(Buffer.from(buffer), encoding);
  } catch (error) {
    return null;
  }
}

function scoreDecodedText(text = "") {
  const value = String(text || "");
  if (!value) return -Infinity;

  const hangulCount = countMatches(value, HANGUL_PATTERN);
  const mojibakeCount = countMatches(value, MOJIBAKE_PATTERN);
  const replacementCount = countMatches(value, REPLACEMENT_PATTERN);
  const delimiterCount = countMatches(value, /[,\t;\n\r]/g);

  return (
    hangulCount * 8 +
    delimiterCount * 0.1 -
    mojibakeCount * 15 -
    replacementCount * 30
  );
}

function getDecodedTextCandidates(buffer) {
  const encodings = ["utf8", "cp949", "euc-kr"];

  return encodings
    .map((encoding) => {
      const text = safeDecode(buffer, encoding);
      return {
        encoding,
        text,
        score: scoreDecodedText(text),
      };
    })
    .filter((item) => item.text && Number.isFinite(item.score))
    .sort((a, b) => b.score - a.score);
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

function readWorkbookFromText(text, options = {}) {
  const normalizedText = String(text || "").replace(/^\uFEFF/, "");

  return XLSX.read(normalizedText, {
    type: "string",
    raw: true,
    cellDates: true,
    cellNF: true,
    cellText: false,
    ...options,
  });
}

function readTextWorkbookFromBuffer(buffer, options = {}) {
  const candidates = getDecodedTextCandidates(buffer);

  let bestWorkbook = null;
  let bestScore = -Infinity;

  for (const candidate of candidates) {
    try {
      const workbook = readWorkbookFromText(candidate.text, options);
      const headerText = workbookToHeaderText(workbook);

      const workbookScore =
        scoreDecodedText(headerText) +
        candidate.score * 0.2 -
        (hasMojibakeText(headerText) ? 1000 : 0);

      if (workbookScore > bestScore) {
        bestScore = workbookScore;
        bestWorkbook = workbook;
      }

      if (!hasMojibakeText(headerText) && /[가-힣]/.test(headerText)) {
        return workbook;
      }
    } catch (error) {
      // 다음 인코딩 후보 시도
    }
  }

  return bestWorkbook;
}

function readWorkbookFromBuffer(buffer, options = {}) {
  const readOptions = {
    ...DEFAULT_XLSX_READ_OPTIONS,
    ...options,
  };

  if (isProbablyTextBuffer(buffer)) {
    const textWorkbook = readTextWorkbookFromBuffer(buffer, options);
    if (textWorkbook) {
      return textWorkbook;
    }
  }

  const workbook = XLSX.read(buffer, readOptions);
  const headerText = workbookToHeaderText(workbook);

  if (!hasMojibakeText(headerText)) {
    return workbook;
  }

  const fallbackWorkbook = readTextWorkbookFromBuffer(buffer, options);
  if (!fallbackWorkbook) {
    return workbook;
  }

  const fallbackHeaderText = workbookToHeaderText(fallbackWorkbook);

  if (hasMojibakeText(fallbackHeaderText)) {
    return workbook;
  }

  return fallbackWorkbook;
}

module.exports = {
  readWorkbookFromBuffer,
  hasMojibakeText,
};
