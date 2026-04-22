function fallbackVSTACK(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (args.length < 2) return null;
  return `{${args.join(";")}}`;
}

function fallbackTOCOL(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (!args.length) return null;

  const arrayExpr = String(args[0] || "").trim();
  if (!arrayExpr) return null;

  // 단일열이면 그대로 사용
  if (/^'[^']+'![A-Z]+\d+:[A-Z]+\d+$/i.test(arrayExpr)) return arrayExpr;
  if (/^INDEX\(.+,\s*,\s*1\)$/i.test(arrayExpr)) return arrayExpr;

  // 보수적 fallback: 첫 번째 열만 반환
  return `INDEX(${arrayExpr},,1)`;
}

function escapeRegExp(s = "") {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function fallbackLET(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (args.length < 3) return null;

  const body = String(args[args.length - 1] || "").trim();
  const bindings = args.slice(0, -1);
  if (bindings.length % 2 !== 0) return null;

  let out = body;

  for (let i = bindings.length - 2; i >= 0; i -= 2) {
    const name = String(bindings[i] || "").trim();
    const expr = String(bindings[i + 1] || "").trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return null;

    const re = new RegExp(`\\b${escapeRegExp(name)}\\b`, "g");
    out = out.replace(re, `(${expr})`);
  }

  return out;
}

function fallbackMAP(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (args.length !== 2) return null;

  const arrayExpr = String(args[0] || "").trim();
  const lambdaExpr = String(args[1] || "").trim();

  const m = lambdaExpr.match(
    /^LAMBDA\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*(.+)\)$/i,
  );
  if (!m) return null;

  return null;
}

function fallbackBYROW(argSource = "") {
  const args = splitTopLevelArgs(argSource);
  if (args.length !== 2) return null;

  const rangeExpr = String(args[0] || "").trim();
  const lambdaExpr = String(args[1] || "").trim();

  const m = lambdaExpr.match(
    /^LAMBDA\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*(.+)\)$/i,
  );
  if (!m) return null;

  const param = m[1];
  const body = m[2];

  // 최소 지원: BYROW(range, LAMBDA(r, SUM(r)))
  if (new RegExp(`\\bSUM\\(${escapeRegExp(param)}\\)`, "i").test(body)) {
    return `MMULT(N(${rangeExpr}), TRANSPOSE(COLUMN(${rangeExpr})^0))`;
  }

  return null;
}

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

function fallbackSortedGroupAverageTable(input = "") {
  const src = String(input || "").replace(/\s+/g, " ");

  // 패턴:
  // SORTBY(<groupAvgExpr>, CHOOSECOLS(<sameGroupAvgExpr>, 2), -1 or 1)
  const m = src.match(
    /^=?(?:IFERROR\()?SORTBY\(\s*(LET\(\s*keys\s*,\s*UNIQUE\([^)]+\)\s*,\s*HSTACK\([\s\S]*?AVERAGE\(\s*arr\s*\)[\s\S]*?\)\s*\)\s*)\s*,\s*CHOOSECOLS\(\s*\1\s*,\s*2\s*\)\s*,\s*(-?1)\s*\)(?:,\s*([\s\S]+?)\s*)?\)?$/i,
  );
  if (!m) return null;

  const groupExpr = String(m[1] || "").trim();
  const orderExpr = String(m[2] || "-1").trim();
  const avgTable = fallbackGroupAverageTable(groupExpr);
  if (!avgTable) return null;

  const avgTableExpr = String(avgTable).replace(/^=/, "").trim();

  const ascending = orderExpr === "1";
  const sorted = `SORT(${avgTableExpr}, 2, ${ascending ? "TRUE" : "FALSE"})`;

  // 원본이 IFERROR로 감싸져 있었으면 유지
  if (/^=?(?:IFERROR\()/i.test(src)) {
    const fallbackArg = m[3] != null ? String(m[3]).trim() : `""`;
    return `=IFERROR(${sorted}, ${fallbackArg})`;
  }
  return `=${sorted}`;
}

function fallbackGroupAverageTable(input = "") {
  const src = String(input || "").replace(/\s+/g, " ");

  // 패턴:
  // LET(keys, UNIQUE(<keyRange>), HSTACK(keys, MAP(keys, LAMBDA(k, LET(arr, IFERROR(FILTER(<valRange>, (<keyRange>=k)), ""), IF(COUNTA(arr)=0, "", AVERAGE(arr)))))))
  const m = src.match(
    /LET\(\s*keys\s*,\s*UNIQUE\(([^)]+)\)\s*,\s*HSTACK\(\s*keys\s*,\s*MAP\(\s*keys\s*,\s*LAMBDA\(\s*k\s*,[\s\S]*?AVERAGE\(\s*arr\s*\)[\s\S]*?\)\s*\)\s*\)\s*\)/i,
  );
  if (!m) return null;

  const keyRange = String(m[1] || "").trim();

  // FILTER(valRange, (keyRange=k)) 안의 valRange를 추출
  const fm = src.match(/FILTER\(\s*([^,]+)\s*,\s*\(([^=]+)=\s*k\)\s*\)/i);
  if (!fm) return null;

  const valRange = String(fm[1] || "").trim();
  const filterKeyRange = String(fm[2] || "").trim();

  // keyRange 일치 확인이 안 되면 보수적으로 중단
  const norm = (s) => String(s || "").replace(/\s+/g, "");
  if (norm(keyRange) !== norm(filterKeyRange)) return null;

  const keysExpr = `UNIQUE(${keyRange})`;
  return `={${keysExpr}, ARRAYFORMULA(IF(${keysExpr}="",,AVERAGEIF(${keyRange}, ${keysExpr}, ${valRange})))}`;
}

function tryGenerateFallbackFormula(formula = "", compatibility = null) {
  const blockers = Array.isArray(compatibility?.blockers)
    ? compatibility.blockers.map((x) => String(x || "").toUpperCase())
    : [];

  let next = String(formula || "");
  const appliedFunctions = [];
  const sortedGroupAvgFallback = fallbackSortedGroupAverageTable(next);
  if (sortedGroupAvgFallback && sortedGroupAvgFallback !== next) {
    next = sortedGroupAvgFallback;
    appliedFunctions.push("GROUP_AVERAGE_TABLE");
    appliedFunctions.push("SORTBY");
  } else {
    const groupAvgFallback = fallbackGroupAverageTable(next);
    if (groupAvgFallback && groupAvgFallback !== next) {
      next = groupAvgFallback;
      appliedFunctions.push("GROUP_AVERAGE_TABLE");
    }
  }

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

  if (blockers.includes("VSTACK")) {
    const replaced = replaceFunctionCalls(next, "VSTACK", fallbackVSTACK);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("VSTACK");
    }
  }

  if (blockers.includes("TOCOL")) {
    const replaced = replaceFunctionCalls(next, "TOCOL", fallbackTOCOL);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("TOCOL");
    }
  }

  if (blockers.includes("LET")) {
    const replaced = replaceFunctionCalls(next, "LET", fallbackLET);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("LET");
    }
  }

  if (blockers.includes("MAP")) {
    const replaced = replaceFunctionCalls(next, "MAP", fallbackMAP);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("MAP");
    }
  }

  if (blockers.includes("BYROW")) {
    const replaced = replaceFunctionCalls(next, "BYROW", fallbackBYROW);
    if (replaced !== next) {
      next = replaced;
      appliedFunctions.push("BYROW");
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
