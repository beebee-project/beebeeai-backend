function detectFormulaCompatibility(formula = "") {
  const f = String(formula || "").toUpperCase();

  const excelOnly = [];
  const sheetsOnly = [];

  // 이 판정은 "Sheets 미지원 함수 탐지"용이다.
  // Sheets 공식 지원 함수는 여기 넣지 않는다.
  const excelOnlyFns = ["SORTBY("];

  const sheetsOnlyFns = [
    "IMPORTRANGE(",
    "GOOGLEFINANCE(",
    "REGEXEXTRACT(",
    "REGEXREPLACE(",
  ];

  for (const fn of excelOnlyFns) {
    if (f.includes(fn)) excelOnly.push(fn.replace("(", ""));
  }

  for (const fn of sheetsOnlyFns) {
    if (f.includes(fn)) sheetsOnly.push(fn.replace("(", ""));
  }

  if (excelOnly.length) {
    return {
      level: "excel_only",
      blockers: excelOnly,
    };
  }

  if (sheetsOnly.length) {
    return {
      level: "sheets_only",
      blockers: sheetsOnly,
    };
  }

  return {
    level: "common",
    blockers: [],
  };
}

const FALLBACK_BLOCKERS = ["SORTBY"];

function shouldAttemptCompatibilityFallback(compatibility = null) {
  const blockers = Array.isArray(compatibility?.blockers)
    ? compatibility.blockers.map((x) => String(x || "").toUpperCase())
    : [];

  return blockers.some((b) => FALLBACK_BLOCKERS.includes(b));
}

module.exports = {
  detectFormulaCompatibility,
  shouldAttemptCompatibilityFallback,
  FALLBACK_BLOCKERS,
};
