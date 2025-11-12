const XLSX = require("xlsx");

function indexToColumnLetter(idx) {
  let n = idx + 1,
    s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// 숫자 비슷한 값인지 (기존 isNumericLike 로직과 동일하게 유지)
function isNumericLike(v) {
  if (v === null || v === undefined) return false;
  if (v instanceof Date) return false;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "" || /[^\d.+\-eE]/.test(s)) return false;
  const n = Number(s);
  return Number.isFinite(n);
}

function isBooleanLike(v) {
  if (typeof v === "boolean") return true;
  const s = String(v).trim().toLowerCase();
  return ["true", "false", "yes", "no", "y", "n", "예", "아니오"].includes(s);
}

// 단순 날짜 패턴 (JS Date 객체이거나, 문자열 패턴 기반)
function isDateLike(v) {
  if (v instanceof Date) return true;
  if (v == null) return false;
  const s = String(v).trim();
  if (!s) return false;

  // YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD
  if (/^\d{4}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])/.test(s)) {
    return true;
  }

  // "2025-01-01 12:34" 같이 날짜+시간
  if (
    /^\d{4}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])\s+\d{1,2}:\d{2}/.test(
      s
    )
  ) {
    return true;
  }

  return false;
}

// 시간만 있는 패턴 (HH:MM 또는 HH:MM:SS)
function isTimeLike(v) {
  if (v == null) return false;
  const s = String(v).trim();
  if (!s) return false;

  // 0:00 ~ 23:59(:59)
  return /^([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/.test(s);
}

function isTimeLikeString(s) {
  // 09:30, 9:30, 09:30:15
  return /^([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/.test(s);
}

function isDateLikeString(s) {
  // 2024-01-01 / 2024.01.01 / 2024/01/01
  if (/^(19|20)\d{2}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])$/.test(s))
    return true;
  // 20240101 같은 8자리 숫자도 간단히 케이스 추가 가능
  if (/^(19|20)\d{6}$/.test(s)) return true;
  return false;
}

function isDateTimeLikeString(s) {
  // 2024-01-01 09:30 / 2024-01-01T09:30:00
  return /^(19|20)\d{2}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])[ T]([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/.test(
    s
  );
}

// 셀 하나의 타입 분류 - number / date / time / text / empty
function detectCellType(v) {
  if (v === null || v === undefined) return "empty";
  if (v instanceof Date) return "date";

  // 숫자 타입은 그냥 숫자로 본다 (엑셀에서 날짜를 Date로 주는 경우가 많음)
  if (typeof v === "number") {
    return "number";
  }

  const s = String(v).trim();
  if (!s) return "empty";

  if (isDateLike(s)) return "date";
  if (isTimeLike(s)) return "time";
  if (isNumericLike(s)) return "number";

  return "text";
}

function analyzeSamples(values) {
  let numeric = 0;
  let date = 0;
  let datetime = 0;
  let time = 0;
  let bool = 0;
  let text = 0;

  for (const v of values) {
    if (v == null) continue;

    if (v instanceof Date) {
      // 엑셀 날짜가 Date로 들어오는 경우
      datetime++;
      continue;
    }

    const s = String(v).trim();
    if (!s) continue;

    if (isBooleanLike(v)) {
      bool++;
    } else if (isDateTimeLikeString(s)) {
      datetime++;
    } else if (isDateLikeString(s)) {
      date++;
    } else if (isTimeLikeString(s)) {
      time++;
    } else if (isNumericLike(v)) {
      numeric++;
    } else {
      text++;
    }
  }

  const total = numeric + date + datetime + time + bool + text || 1;

  const ratios = {
    numericRatio: numeric / total,
    dateRatio: date / total,
    datetimeRatio: datetime / total,
    timeRatio: time / total,
    booleanRatio: bool / total,
    textRatio: text / total,
  };

  // 대표 타입(dominantType) 추론
  const entries = [
    ["number", ratios.numericRatio],
    ["date", ratios.dateRatio + ratios.datetimeRatio],
    ["time", ratios.timeRatio],
    ["boolean", ratios.booleanRatio],
    ["text", ratios.textRatio],
  ].sort((a, b) => b[1] - a[1]);

  const [topType, topRatio] = entries[0];
  const dominantType = topRatio >= 0.5 ? topType : "mixed";

  return {
    ...ratios,
    dominantType,
    sampleCount: total,
  };
}

function detectHeaderRowIndex(json, maxScanRows = 10) {
  const limit = Math.min(json.length, maxScanRows);
  let bestRow = 0;
  let bestScore = -1;

  for (let i = 0; i < limit; i++) {
    const row = json[i] || [];
    let nonEmpty = 0;

    for (const cell of row) {
      if (cell !== null && cell !== undefined && String(cell).trim() !== "") {
        nonEmpty++;
      }
    }

    if (nonEmpty > bestScore) {
      bestScore = nonEmpty;
      bestRow = i;
    }
  }

  // 시트 전체에서 "처음으로 값이 있는 행"을 헤더로 사용
  if (bestScore <= 0) {
    for (let i = 0; i < json.length; i++) {
      const row = json[i] || [];
      if (
        row.some(
          (c) => c !== null && c !== undefined && String(c).trim() !== ""
        )
      ) {
        return i; // 첫 non-empty 행
      }
    }
    return 0; // 그래도 없으면 0
  }

  return bestRow; // 0-based index
}

// workbook (XLSX.read 결과) → allSheetsData
function buildAllSheetsData(workbook) {
  const allSheetsData = {};

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws || !ws["!ref"]) continue;

    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rowCount = range.e.r + 1;

    // 2차원 배열 (각 원소가 행)
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (!json || json.length < 1) continue;

    // 1) 시트 전체에서 "값이 있는 첫 행 / 마지막 행" 찾기
    let firstNonEmpty = null;
    let lastNonEmpty = null;
    for (let i = 0; i < json.length; i++) {
      const row = json[i] || [];
      if (row.some((c) => c != null && String(c).trim() !== "")) {
        if (firstNonEmpty === null) firstNonEmpty = i;
        lastNonEmpty = i;
      }
    }
    if (firstNonEmpty === null) continue;

    // 2) 헤더처럼 보이는 모든 행 인덱스를 수집
    //    - nonEmpty >= 2
    //    - 그 중 60% 이상이 문자열인 행
    const headerRowIndexes = [];
    for (let i = firstNonEmpty; i <= lastNonEmpty; i++) {
      const row = json[i] || [];
      let nonEmpty = 0;
      let textLike = 0;

      for (const cell of row) {
        if (cell != null && String(cell).trim() !== "") {
          nonEmpty++;
          if (typeof cell === "string") textLike++;
        }
      }

      if (nonEmpty >= 2 && textLike / nonEmpty >= 0.6) {
        headerRowIndexes.push(i);
      }
    }

    if (headerRowIndexes.length === 0) continue;

    // 3) 헤더처럼 보이는 각 행에서 metaData 채우기
    const metaData = {};

    for (const headerIndex of headerRowIndexes) {
      const headers = json[headerIndex] || [];

      // 이 헤더 바로 아래에서 실제 데이터가 시작하는 행 (0-based index)
      let dataStart = headerIndex + 1;
      for (let r = dataStart; r < json.length; r++) {
        const row = json[r] || [];
        if (row.some((c) => c != null && String(c).trim() !== "")) {
          dataStart = r;
          break;
        }
      }

      headers.forEach((header, idx) => {
        const name = String(header || "").trim();
        if (!name) return;

        // 같은 이름의 헤더가 위에서 이미 등록되었으면 첫 번째 것만 사용
        if (metaData[name]) return;

        const values = [];
        for (
          let r = dataStart;
          r < Math.min(json.length, dataStart + 200);
          r++
        ) {
          const val = json[r]?.[idx];
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            values.push(val);
          }
        }

        const stats = analyzeSamples(values);

        metaData[name] = {
          columnLetter: indexToColumnLetter(idx),
          // ✅ 이 컬럼의 실제 데이터 시작/끝 행 (1-based)
          startRow: dataStart + 1,
          lastRow: lastNonEmpty + 1,
          ...stats,
        };
      });
    }

    if (Object.keys(metaData).length === 0) continue;

    // 4) 전체 범위 정보 (fallback 용도)
    const firstHeaderRow = Math.min(...headerRowIndexes);
    const startRow = firstHeaderRow + 2; // 0-based → 1-based
    const lastDataRow = lastNonEmpty + 1;

    allSheetsData[sheetName] = {
      rowCount,
      startRow,
      lastDataRow,
      metaData,
    };
  }

  return allSheetsData;
}

module.exports = { buildAllSheetsData, detectCellType };
