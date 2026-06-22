function balanceCheck(str, openCh, closeCh) {
  let n = 0;
  for (const c of str) {
    if (c === openCh) n++;
    else if (c === closeCh) n--;
    if (n < 0) return false;
  }
  return n === 0;
}

function quoteBalance(str) {
  // Excel/Sheets는 보통 " 로 문자열 감쌈. ""는 escape.
  let q = 0;
  for (let i = 0; i < str.length; i++) {
    if (str[i] === '"') {
      if (str[i + 1] === '"')
        i++; // escaped quote
      else q++;
    }
  }
  return q % 2 === 0;
}

function validateFormula(output) {
  const issues = [];
  const t = String(output ?? "").trim();
  if (!t) issues.push("EMPTY_OUTPUT");
  if (
    t &&
    !t.startsWith("=") &&
    !/^(SELECT|WITH)\b/i.test(t) &&
    !t.includes("prop(")
  ) {
    issues.push("NOT_FORMULA_PREFIX");
  }
  if (t.startsWith("=")) {
    if (!quoteBalance(t)) issues.push("UNBALANCED_QUOTES");
    if (!balanceCheck(t, "(", ")")) issues.push("UNBALANCED_PARENS");
    if (/\bundefined\b|\bnull\b/i.test(t)) issues.push("CONTAINS_UNDEFINED");
    if (/=ERROR\s*\(/i.test(t)) issues.push("ERROR_FORMULA");
  }
  return { ok: issues.length === 0, kind: "formula", issues };
}

module.exports = {
  validateFormula,
};
