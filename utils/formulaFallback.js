function splitTopLevelArgs(src = "") {
  const out = [];
  let cur = "";
  let paren = 0;
  let brace = 0;
  let bracket = 0;
  let quote = null;

  for (let i = 0; i < src.length; i += 1) {
    const ch = src[i];
    const prev = i > 0 ? src[i - 1] : "";

    if (quote) {
      cur += ch;
      if (ch === quote && prev !== "\\") quote = null;
      continue;
    }

    if (ch === '"' || ch === "'") {
      quote = ch;
      cur += ch;
      continue;
    }

    if (ch === "(") paren += 1;
    else if (ch === ")") paren -= 1;
    else if (ch === "{") brace += 1;
    else if (ch === "}") brace -= 1;
    else if (ch === "[") bracket += 1;
    else if (ch === "]") bracket -= 1;

    if (ch === "," && paren === 0 && brace === 0 && bracket === 0) {
      out.push(cur.trim());
      cur = "";
      continue;
    }

    cur += ch;
  }

  if (cur.trim()) out.push(cur.trim());
  return out;
}

function replaceFunctionCalls(input = "", fnName = "", replacer) {
  const src = String(input || "");
  const upper = src.toUpperCase();
  const token = `${String(fnName || "").toUpperCase()}(`;

  if (!token || token === "(") return src;

  let out = "";
  let i = 0;

  while (i < src.length) {
    const idx = upper.indexOf(token, i);
    if (idx < 0) {
      out += src.slice(i);
      break;
    }

    out += src.slice(i, idx);

    const argsStart = idx + token.length;
    let depth = 1;
    let j = argsStart;
    let quote = null;

    while (j < src.length && depth > 0) {
      const ch = src[j];
      const prev = j > 0 ? src[j - 1] : "";

      if (quote) {
        if (ch === quote && prev !== "\\") quote = null;
        j += 1;
        continue;
      }

      if (ch === '"' || ch === "'") {
        quote = ch;
        j += 1;
        continue;
      }

      if (ch === "(") depth += 1;
      else if (ch === ")") depth -= 1;
      j += 1;
    }

    if (depth !== 0) {
      out += src.slice(idx);
      break;
    }

    const argSource = src.slice(argsStart, j - 1);
    const replaced = replacer(argSource);
    out += replaced || src.slice(idx, j);
    i = j;
  }

  return out;
}

function fallbackHSTACK(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (args.length < 2) return null;
  return `{${args.join(",")}}`;
}

function fallbackTAKE(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (args.length < 2) return null;

  const arrayExpr = args[0];
  const rowsExpr = String(args[1] || "").trim();
  const colsExpr = String(args[2] || "").trim();

  if (!rowsExpr) return null;

  const trimmedArray = String(arrayExpr || "").trim();
  const looksSingleColumn =
    /,\s*,\s*1\)\s*$/.test(trimmedArray) || // INDEX(...,,1)
    /'[^']+'![A-Z]+\d+:[A-Z]+\d+\s*$/.test(trimmedArray) || // 단일열 range
    /^SORT\(\s*'[^']+'![A-Z]+\d+:[A-Z]+\d+\s*,\s*1\s*,\s*(TRUE|FALSE)\s*\)$/i.test(
      trimmedArray,
    ); // SORT(single-column-range,1,TRUE/FALSE)

  if (/^-?\d+$/.test(rowsExpr)) {
    const rowNum = Number(rowsExpr);
    if (rowNum > 0) {
      if (!colsExpr && looksSingleColumn) {
        return `INDEX(${arrayExpr},SEQUENCE(${rowNum}))`;
      }

      const colSeq =
        colsExpr && /^\d+$/.test(colsExpr)
          ? `SEQUENCE(1,${colsExpr})`
          : `SEQUENCE(1,COLUMNS(${arrayExpr}))`;
      return `INDEX(${arrayExpr},SEQUENCE(${rowNum}),${colSeq})`;
    }

    if (rowNum === -1 && !colsExpr) {
      if (looksSingleColumn) {
        return `INDEX(${arrayExpr},ROWS(${arrayExpr}))`;
      }
      return `INDEX(${arrayExpr},ROWS(${arrayExpr}),SEQUENCE(1,COLUMNS(${arrayExpr})))`;
    }
  }

  return null;
}

function fallbackSORTBY(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (args.length < 3) return null;

  const arrayExpr = args[0];
  const keyExpr = args[1];
  const orderExpr = String(args[2] || "").trim();
  const ascending = !String(orderExpr).startsWith("-");

  const normArray = String(arrayExpr || "").replace(/\s+/g, "");
  const normKey = String(keyExpr || "").replace(/\s+/g, "");

  // self-sort 단일열은 더 단순한 SORT로 축약
  if (normArray && normArray === normKey) {
    return `SORT(${arrayExpr},1,${ascending ? "TRUE" : "FALSE"})`;
  }

  return `INDEX(SORT({${arrayExpr},${keyExpr}},2,${ascending ? "TRUE" : "FALSE"}),,1)`;
}

function tryGenerateFallbackFormula(formula = "", compatibility = null) {
  const blockers = Array.isArray(compatibility?.blockers)
    ? compatibility.blockers.map((x) => String(x || "").toUpperCase())
    : [];

  let next = String(formula || "");
  const appliedFunctions = [];

  if (blockers.includes("SORTBY")) {
    const replaced = replaceFunctionCalls(next, "SORTBY", fallbackSORTBY);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("SORTBY");
    }
  }

  if (blockers.includes("TAKE")) {
    const replaced = replaceFunctionCalls(next, "TAKE", fallbackTAKE);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("TAKE");
    }
  }

  if (blockers.includes("HSTACK")) {
    const replaced = replaceFunctionCalls(next, "HSTACK", fallbackHSTACK);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("HSTACK");
    }
  }

  if (!appliedFunctions.length) return null;

  return {
    formula: next,
    appliedFunctions,
  };
}

module.exports = {
  tryGenerateFallbackFormula,
};
