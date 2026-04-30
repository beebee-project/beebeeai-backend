const { refFromHeaderSpec } = require("./builderHelpers");

function _q(s) {
  return `"${String(s ?? "").replace(/"/g, '""')}"`;
}

function _isNumericLiteral(v) {
  if (v == null) return false;
  const s = String(v).replace(/,/g, "").trim();
  return /^-?\d+(\.\d+)?$/.test(s);
}

function _isIsoDateLiteral(v) {
  return /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(String(v || "").trim());
}

function _normText(expr, caseSensitive = false) {
  const t = `TRIM(${expr}&"")`;
  return caseSensitive ? t : `LOWER(${t})`;
}

function _coerceNumber(expr) {
  return `IFERROR(VALUE(TRIM(${expr}&"")), ${expr})`;
}

function _coerceDate(expr) {
  return `IFERROR(DATEVALUE(TRIM(${expr}&"")), ${expr})`;
}

const ORDINAL_MAPS = {
  position: ["사원", "대리", "과장", "차장", "부장"],
};

function _formatScalar(
  v,
  formatValue,
  valueType = null,
  caseSensitive = false,
) {
  if (v == null) return _q("");

  // 셀 참조 / 범위는 그대로
  if (typeof v === "string" && valueType !== "text") {
    const s = v.trim();
    if (
      /^\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?$/i.test(s) ||
      /^([^!'\s]+|'[^']+')!\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?$/i.test(
        s,
      )
    ) {
      return s;
    }
  }

  if (valueType === "number" || _isNumericLiteral(v)) {
    return String(Number(String(v).replace(/,/g, "").trim()));
  }

  if (valueType === "cell") {
    return String(v).trim();
  }

  if (valueType === "date" || _isIsoDateLiteral(v)) {
    const iso = String(v).trim().replace(/[./]/g, "-");
    return `DATEVALUE(${_q(iso)})`;
  }

  let lit = typeof formatValue === "function" ? formatValue(v) : _q(v);
  if (typeof v === "string" && valueType === "text") {
    const s = String(lit || "").trim();
    const isQuoted = /^".*"$/.test(s);
    if (!isQuoted) lit = _q(v);
  }
  if (caseSensitive) return lit;
  return `LOWER(TRIM(${lit}&""))`;
}

function _buildLeafExpr(cond, ctx, formatValue) {
  const caseSensitive = !!ctx?.formatOptions?.case_sensitive;

  const header =
    cond?.ref?.header ||
    (typeof cond?.header === "string" ? cond.header : null) ||
    (typeof cond?.target === "string" ? cond.target : null) ||
    cond?.target?.header ||
    cond?.hint ||
    null;

  const ref =
    cond?.ref ||
    (header
      ? refFromHeaderSpec(ctx, {
          header,
          sheet: ctx?.resolved?.baseSheet || undefined,
        })
      : null) ||
    (header ? refFromHeaderSpec(ctx, header) : null);

  if (!ref?.range) return null;

  const range = ref.range;
  const op = String(cond?.operator || "=")
    .trim()
    .toLowerCase();
  const rawVal = cond?.value;
  const valueType = cond?.value_type || null;

  if (
    (op === "between" || op === "not_between") &&
    cond?.min != null &&
    cond?.max != null
  ) {
    const isDate =
      valueType === "date" ||
      _isIsoDateLiteral(cond.min) ||
      _isIsoDateLiteral(cond.max);
    const isNumber =
      valueType === "number" ||
      (_isNumericLiteral(cond.min) && _isNumericLiteral(cond.max));

    let left;
    let minExpr;
    let maxExpr;

    if (isDate) {
      left = _coerceDate(range);
      minExpr = _formatScalar(cond.min, formatValue, "date", caseSensitive);
      maxExpr = _formatScalar(cond.max, formatValue, "date", caseSensitive);
    } else if (isNumber) {
      left = _coerceNumber(range);
      minExpr = _formatScalar(cond.min, formatValue, "number", caseSensitive);
      maxExpr = _formatScalar(cond.max, formatValue, "number", caseSensitive);
    } else {
      left = _normText(range, caseSensitive);
      minExpr = _formatScalar(cond.min, formatValue, "text", caseSensitive);
      maxExpr = _formatScalar(cond.max, formatValue, "text", caseSensitive);
    }

    const expr = `((${left}>=${minExpr})*(${left}<=${maxExpr}))`;
    return op === "not_between" ? `(${expr}=0)` : expr;
  }

  if (valueType === "ordinal_text" || cond?.role === "ordinal_filter") {
    const headerText = String(header || "").trim();
    const order = /(직급|position|rank|title)/i.test(headerText)
      ? ORDINAL_MAPS.position
      : null;

    if (order) {
      const idx = order.indexOf(String(rawVal || "").trim());
      if (idx >= 0) {
        const rankExpr = `MATCH(TRIM(${range}&""), {"${order.join('","')}"}, 0)`;
        return `(${rankExpr}${op}${idx + 1})`;
      }
    }
  }

  if (valueType === "aggregate" || cond?.role === "aggregate_filter") {
    const agg = String(cond?.aggregate || "").toLowerCase();
    if (
      agg === "average" ||
      agg === "avg" ||
      agg === "mean" ||
      agg === "평균"
    ) {
      const left = _coerceNumber(range);
      const avgExpr = `AVERAGE(${left})`;

      if (op === ">=") return `(NOT(${left}<${avgExpr}))`;
      if (op === "<=") return `(NOT(${left}>${avgExpr}))`;
      if (op === ">") return `(${left}>${avgExpr})`;
      if (op === "<") return `(${left}<${avgExpr})`;

      return `(NOT(${left}<${avgExpr}))`;
    }
  }

  if (rawVal == null || (typeof rawVal === "string" && rawVal.trim() === "")) {
    return null;
  }

  // text operators
  if (op === "contains") {
    const needle = _formatScalar(rawVal, formatValue, "text", caseSensitive);
    return `ISNUMBER(SEARCH(${needle}, ${_normText(range, caseSensitive)}))`;
  }

  if (op === "starts_with" || op === "startswith") {
    const needle = _formatScalar(rawVal, formatValue, "text", caseSensitive);
    const col = _normText(range, caseSensitive);
    return `LEFT(${col}, LEN(${needle}))=${needle}`;
  }

  if (op === "ends_with" || op === "endswith") {
    const needle = _formatScalar(rawVal, formatValue, "text", caseSensitive);
    const col = _normText(range, caseSensitive);
    return `RIGHT(${col}, LEN(${needle}))=${needle}`;
  }

  // number compare
  if (valueType === "number" || _isNumericLiteral(rawVal)) {
    const left = _coerceNumber(range);
    const right = _formatScalar(rawVal, formatValue, "number", caseSensitive);
    return `(${left}${op}${right})`;
  }

  // date compare
  if (valueType === "date" || _isIsoDateLiteral(rawVal)) {
    const left = _coerceDate(range);
    const right = _formatScalar(rawVal, formatValue, "date", caseSensitive);
    return `(${left}${op}${right})`;
  }

  // default text compare
  const left = _normText(range, caseSensitive);
  const rightScalar = _formatScalar(
    rawVal,
    formatValue,
    valueType === "cell" ? "cell" : "text",
    caseSensitive,
  );
  const right =
    valueType === "cell" ? _normText(rightScalar, caseSensitive) : rightScalar;

  if (op === "=" || op === "==") return `(${left}=${right})`;
  if (op === "<>" || op === "!=") return `(${left}<>${right})`;
  if (op === ">" || op === ">=" || op === "<" || op === "<=") {
    return `(${left}${op}${right})`;
  }

  return `(${left}=${right})`;
}

function _buildGroupExpr(node, ctx, formatValue) {
  if (!node) return null;

  if (node.logical_operator && Array.isArray(node.conditions)) {
    const exprs = node.conditions
      .map((c) => buildSingleConditionExpr(c, ctx, formatValue))
      .filter(Boolean);

    if (!exprs.length) return null;

    const logical = String(node.logical_operator || "AND").toUpperCase();
    if (logical === "OR") return `((${exprs.join("+")})>0)`;
    return `(${exprs.join("*")})`;
  }

  return _buildLeafExpr(node, ctx, formatValue);
}

function buildSingleConditionExpr(cond, ctx, formatValue) {
  return _buildGroupExpr(cond, ctx, formatValue);
}

function buildConditionMask(ctx, formatValue) {
  const hasResolvedFilterColumns = Array.isArray(ctx?.resolved?.filterColumns);

  const raw = hasResolvedFilterColumns
    ? ctx.resolved.filterColumns
    : Array.isArray(ctx?.intent?.filters)
      ? ctx.intent.filters
      : Array.isArray(ctx?.intent?.conditions)
        ? ctx.intent.conditions
        : [];

  const sanitizedRaw = raw.filter((f) => {
    const rawText = String(ctx?.intent?.raw_message || ctx?.message || "");
    const value = String(f?.value ?? "").trim();

    if (
      /(존재하지\s*않는|존재하지않는|없는|없음)/.test(rawText) &&
      ["않", "않는", "없는", "없음", "존재하지", "존재하지않는"].includes(value)
    ) {
      return false;
    }

    return true;
  });

  const exprs = sanitizedRaw
    .map((c) => buildSingleConditionExpr(c, ctx, formatValue))
    .filter(Boolean);

  const _extractPrimaryRangeKey = (expr = "") => {
    const s = String(expr || "");

    // equality 기준: LEFT side 기준으로만 range 추출
    const leftEq = s.match(/LOWER\(TRIM\(([^)]*![A-Z]+\d+:[A-Z]+\d+)&""\)\)/);
    if (leftEq) {
      const m = leftEq[1].match(/!([A-Z]+)\d+:\1\d+/);
      if (m) return m[1];
    }

    // ordinal 기준
    const leftOrdinal = s.match(
      /MATCH\(TRIM\(([^)]*![A-Z]+\d+:[A-Z]+\d+)&""\)/,
    );
    if (leftOrdinal) {
      const m = leftOrdinal[1].match(/!([A-Z]+)\d+:\1\d+/);
      if (m) return m[1];
    }

    // fallback (최후)
    const m = s.match(/'[^']+'!([A-Z]+)\d+:\1\d+/);
    return m ? m[1] : null;
  };

  const _isOrdinalExpr = (expr = "") =>
    /MATCH\(TRIM\(.+?\&""\),\s*\{/.test(String(expr));

  const _isPlainTextEqualityExpr = (expr = "") =>
    /LOWER\(TRIM\(.+?\&""\)\)\s*=\s*LOWER\(TRIM\("/.test(String(expr));

  const _isAggregateExpr = (expr = "") => /AVERAGE\(/i.test(String(expr));

  const _isAggregateTextCompareExpr = (expr = "") =>
    /LOWER\(TRIM\(.+?\&""\)\)\s*(?:>=|<=|>|<|=)\s*LOWER\(TRIM\("평균"/.test(
      String(expr),
    );

  const ordinalRangeKeys = new Set(
    exprs.filter(_isOrdinalExpr).map(_extractPrimaryRangeKey).filter(Boolean),
  );

  const aggregateTextCompareKeys = new Set(
    exprs
      .filter(_isAggregateTextCompareExpr)
      .map(_extractPrimaryRangeKey)
      .filter(Boolean),
  );

  const aggregateFormulaKeys = new Set(
    exprs.filter(_isAggregateExpr).map(_extractPrimaryRangeKey).filter(Boolean),
  );

  const filteredExprs = exprs.filter((e) => {
    const key = _extractPrimaryRangeKey(e);

    // ordinal 조건 자체는 유지
    if (_isOrdinalExpr(e)) return true;

    // ordinal과 같은 열의 plain equality만 제거
    if (key && ordinalRangeKeys.has(key) && _isPlainTextEqualityExpr(e)) {
      return false;
    }

    // aggregate formula와 같은 열에 있는 "평균" 텍스트 비교만 제거
    // 단순히 같은 열이라는 이유로 숫자 비교/다른 조건을 제거하지 않는다.
    if (
      key &&
      aggregateFormulaKeys.has(key) &&
      aggregateTextCompareKeys.has(key) &&
      _isAggregateTextCompareExpr(e)
    ) {
      return false;
    }

    return true;
  });

  const _cellRefEqualityInfo = (expr = "") => {
    const s = String(expr || "");
    const m = s.match(
      /LOWER\(TRIM\(([^)]*![A-Z]+\d+:[A-Z]+\d+)&""\)\)\s*=\s*LOWER\(TRIM\(([A-Z]+\d+)&""\)\)/i,
    );
    if (!m) return null;
    return { range: m[1], cell: m[2].toUpperCase() };
  };

  const _quotedCellTextEqualityInfo = (expr = "") => {
    const s = String(expr || "");
    const m = s.match(
      /LOWER\(TRIM\(([^)]*![A-Z]+\d+:[A-Z]+\d+)&""\)\)\s*=\s*LOWER\(TRIM\("([A-Z]+\d+)"&""\)\)/i,
    );
    if (!m) return null;
    return { range: m[1], cell: m[2].toUpperCase() };
  };

  const cellRefEqualityKeys = new Set(
    filteredExprs
      .map(_cellRefEqualityInfo)
      .filter(Boolean)
      .map((x) => `${x.range}|${x.cell}`),
  );

  const cellRefSafeExprs = filteredExprs.filter((expr) => {
    const q = _quotedCellTextEqualityInfo(expr);
    if (!q) return true;
    return !cellRefEqualityKeys.has(`${q.range}|${q.cell}`);
  });

  const uniqExprs = [];
  const seen = new Set();
  for (const e of cellRefSafeExprs) {
    const key = String(e).replace(/\s+/g, "");
    if (seen.has(key)) continue;
    seen.add(key);
    uniqExprs.push(e);
  }

  const _extractEqualityKey = (expr = "") => {
    const s = String(expr || "");

    // LOWER(TRIM('나무'!C91:C177&""))=LOWER(TRIM("영업"&""))
    const m = s.match(
      /^\(?LOWER\(TRIM\(([^)]*![A-Z]+\d+:[A-Z]+\d+)&""\)\)\s*=\s*LOWER\(TRIM\("([^"]+)"&""\)\)\)?$/i,
    );

    if (!m) return null;

    const range = m[1];
    const value = m[2];

    const col = range.match(/!([A-Z]+)\d+:\1\d+/);
    if (!col) return null;

    return {
      colKey: col[1],
      value: String(value || "").trim(),
      expr: s,
    };
  };

  const equalityGroups = new Map();

  for (const expr of uniqExprs) {
    const parsed = _extractEqualityKey(expr);
    if (!parsed) continue;

    const arr = equalityGroups.get(parsed.colKey) || [];
    arr.push({ expr, value: parsed.value });
    equalityGroups.set(parsed.colKey, arr);
  }

  const orExprsByOriginal = new Map();

  for (const [_colKey, items] of equalityGroups) {
    const uniqueValues = [...new Set(items.map((x) => x.value))];

    // 같은 열에 서로 다른 equality 값이 2개 이상이면 OR로 묶음
    if (uniqueValues.length < 2) continue;

    const orExpr = `((${items.map((x) => x.expr).join("+")})>0)`;

    for (const item of items) {
      orExprsByOriginal.set(item.expr, orExpr);
    }
  }

  if (orExprsByOriginal.size) {
    const next = [];
    const pushedOr = new Set();

    for (const expr of uniqExprs) {
      const orExpr = orExprsByOriginal.get(expr);

      if (orExpr) {
        if (!pushedOr.has(orExpr)) {
          next.push(orExpr);
          pushedOr.add(orExpr);
        }
        continue;
      }

      next.push(expr);
    }

    uniqExprs.length = 0;
    uniqExprs.push(...next);
  }

  // window 조건 추가
  const win = ctx?.intent?.window;
  if (
    win &&
    String(win.type || "").toLowerCase() === "days" &&
    Number(win.size || 0) > 0
  ) {
    const hdr = win.date_header || "날짜";
    const rr =
      refFromHeaderSpec(ctx, {
        header: hdr,
        sheet: ctx?.resolved?.baseSheet || undefined,
      }) || refFromHeaderSpec(ctx, hdr);

    if (rr?.range) {
      uniqExprs.push(`(${_coerceDate(rr.range)}>=TODAY()-${Number(win.size)})`);
    }
  }

  if (!uniqExprs.length) return null;
  return uniqExprs.length === 1 ? uniqExprs[0] : `(${uniqExprs.join("*")})`;
}

module.exports = {
  buildConditionMask,
  buildSingleConditionExpr,
};
