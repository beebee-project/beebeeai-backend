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

function validateOfficeScripts(code) {
  const issues = [];
  const t = String(code ?? "").trim();
  if (!t) issues.push("EMPTY_CODE");
  if (t && !/function\s+main\s*\(/i.test(t)) issues.push("MISSING_MAIN");
  if (t && !/ExcelScript\./.test(t)) issues.push("MISSING_EXCELSCRIPT_API");
  return { ok: issues.length === 0, kind: "officescripts", issues };
}

function validateAppScript(code) {
  const issues = [];
  const t = String(code ?? "").trim();
  if (!t) issues.push("EMPTY_CODE");
  if (t && !/SpreadsheetApp\./.test(t))
    issues.push("MISSING_SPREADSHEETAPP_API");
  if (t && !/function\s+\w+\s*\(/.test(t)) issues.push("MISSING_FUNCTION");
  return { ok: issues.length === 0, kind: "appscript", issues };
}

function validateMacroResult(target, result) {
  const t = String(target || "").toLowerCase();
  const code = result?.code ?? result?.script ?? result?.result ?? "";
  if (t.includes("office")) return validateOfficeScripts(code);
  if (t.includes("appscript") || t.includes("gas"))
    return validateAppScript(code);
  // target이 애매하면 “코드가 비었는지” 정도만
  const issues = [];
  if (!String(code ?? "").trim()) issues.push("EMPTY_CODE");
  return { ok: issues.length === 0, kind: "macro", issues };
}

module.exports = {
  validateFormula,
  validateMacroResult,
};
