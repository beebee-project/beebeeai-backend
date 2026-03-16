function detectFormulaCompatibility(formula = "") {
  const f = String(formula || "").toUpperCase();

  const excelOnly = [];
  const sheetsOnly = [];

  const excelOnlyFns = [
    "TOCOL(",
    "TAKE(",
    "DROP(",
    "BYROW(",
    "MAP(",
    "HSTACK(",
    "VSTACK(",
  ];

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

module.exports = { detectFormulaCompatibility };
