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

  const isInt = (v) => /^-?\d+$/.test(String(v || "").trim());
  const hasCols = colsExpr.length > 0;
  const isSingleColumnExpr = /^\s*CHOOSECOLS\s*\(/i.test(arrayExpr);

  if (!isInt(rowsExpr)) return null;

  const rowNum = Number(rowsExpr);

  // TAKE(array, 1) / TAKE(array, n)
  if (rowNum > 0) {
    // 단일열 배열은 굳이 2차원 INDEX로 만들지 않아도 된다.
    if (!hasCols && isSingleColumnExpr && rowNum === 1) {
      return `INDEX(${arrayExpr},1)`;
    }

    if (!hasCols && rowNum === 1) {
      return `INDEX(${arrayExpr},1,SEQUENCE(1,COLUMNS(${arrayExpr})))`;
    }

    const colSeq =
      hasCols && isInt(colsExpr)
        ? Number(colsExpr) === 1
          ? "1"
          : `SEQUENCE(1,${colsExpr})`
        : `SEQUENCE(1,COLUMNS(${arrayExpr}))`;

    return `INDEX(${arrayExpr},SEQUENCE(${rowNum}),${colSeq})`;
  }

  // TAKE(array, -1) / TAKE(array, -n)
  if (rowNum < 0) {
    const absRows = Math.abs(rowNum);

    if (!hasCols && isSingleColumnExpr && absRows === 1) {
      return `INDEX(${arrayExpr},ROWS(${arrayExpr}))`;
    }

    if (!hasCols && absRows === 1) {
      return `INDEX(${arrayExpr},ROWS(${arrayExpr}),SEQUENCE(1,COLUMNS(${arrayExpr})))`;
    }

    const rowSeq =
      absRows === 1
        ? `ROWS(${arrayExpr})`
        : `SEQUENCE(${absRows},1,ROWS(${arrayExpr})-${absRows}+1)`;

    const colSeq =
      hasCols && isInt(colsExpr)
        ? Number(colsExpr) === 1
          ? "1"
          : `SEQUENCE(1,${colsExpr})`
        : `SEQUENCE(1,COLUMNS(${arrayExpr}))`;

    return `INDEX(${arrayExpr},${rowSeq},${colSeq})`;
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
