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

function _normalizeCmpOp(op = "") {
  const s = String(op || "")
    .trim()
    .toLowerCase();
  if (s === "==") return "=";
  if (s === "!=") return "<>";
  if (s === "≤") return "<=";
  if (s === "≥") return ">=";
  return s || "=";
}

function _resolveCondRef(cond, ctx) {
  const header =
    cond?.ref?.header ||
    (typeof cond?.header === "string" ? cond.header : null) ||
    (typeof cond?.target === "string" ? cond.target : null) ||
    cond?.target?.header ||
    cond?.hint ||
    null;

  return (
    cond?.ref ||
    (header
      ? refFromHeaderSpec(ctx, {
          header,
          sheet: ctx?.resolved?.baseSheet || undefined,
        })
      : null) ||
    (header ? refFromHeaderSpec(ctx, header) : null)
  );
}

function _windowExprs(ctx) {
  const exprs = [];
  const pairs = [];

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
      const n = Number(win.size || 0);
      exprs.push(`(${_coerceDate(rr.range)}>=TODAY()-${n})`);
      pairs.push(rr.range, `">="&TODAY()-${n}`);
    }
  }

  return { exprs, pairs };
}

function _formatScalar(
  v,
  formatValue,
  valueType = null,
  caseSensitive = false,
) {
  if (v == null) return _q("");

  // 셀 참조 / 범위는 그대로
  if (typeof v === "string") {
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

  if (valueType === "date" || _isIsoDateLiteral(v)) {
    const iso = String(v).trim().replace(/[./]/g, "-");
    return `DATEVALUE(${_q(iso)})`;
  }

  const lit = typeof formatValue === "function" ? formatValue(v) : _q(v);
  if (caseSensitive) return lit;
  return `LOWER(TRIM(${lit}&""))`;
}

function _buildLeafExpr(cond, ctx, formatValue) {
  const caseSensitive = !!ctx?.formatOptions?.case_sensitive;
  const ref = _resolveCondRef(cond, ctx);

  if (!ref?.range) return null;

  const range = ref.range;
  const op = _normalizeCmpOp(cond?.operator || "=");
  const rawVal = cond?.value;
  const valueType = cond?.value_type || null;

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
  const right = _formatScalar(rawVal, formatValue, "text", caseSensitive);

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
  const raw = ctx?.resolved?.filterColumns?.length
    ? ctx.resolved.filterColumns
    : Array.isArray(ctx?.intent?.filters)
      ? ctx.intent.filters
      : Array.isArray(ctx?.intent?.conditions)
        ? ctx.intent.conditions
        : [];

  const exprs = raw
    .map((c) => buildSingleConditionExpr(c, ctx, formatValue))
    .filter(Boolean);

  const win = _windowExprs(ctx);
  if (win.exprs.length) exprs.push(...win.exprs);

  if (!exprs.length) return null;
  return exprs.length === 1 ? exprs[0] : `(${exprs.join("*")})`;
}

function _isLeafIfsSafe(cond = {}) {
  const op = _normalizeCmpOp(cond?.operator || "=");
  if (cond?.logical_operator && Array.isArray(cond?.conditions)) return false;
  if (
    ["contains", "starts_with", "startswith", "ends_with", "endswith"].includes(
      op,
    )
  ) {
    return false;
  }
  if (cond?.value == null) return false;
  return ["=", "<>", ">", ">=", "<", "<="].includes(op);
}

function _leafToIfsPair(cond, ctx, formatValue) {
  if (!_isLeafIfsSafe(cond)) return null;
  const ref = _resolveCondRef(cond, ctx);
  if (!ref?.range) return null;

  const op = _normalizeCmpOp(cond?.operator || "=");
  const rawVal = cond?.value;
  const valueType = cond?.value_type || null;

  let crit;
  if (valueType === "number" || _isNumericLiteral(rawVal)) {
    crit = ["=", "<>"].includes(op)
      ? String(Number(String(rawVal).replace(/,/g, "").trim()))
      : `"${op}${String(Number(String(rawVal).replace(/,/g, "").trim()))}"`;
    return [ref.range, crit];
  }

  if (valueType === "date" || _isIsoDateLiteral(rawVal)) {
    const rhs = _formatScalar(rawVal, formatValue, "date", false);
    crit = ["=", "<>"].includes(op) ? rhs : `"${op}"&${rhs}`;
    return [ref.range, crit];
  }

  const rhs =
    typeof formatValue === "function" ? formatValue(rawVal) : _q(rawVal);
  crit = ["=", "<>"].includes(op) ? rhs : `"${op}"&${rhs}`;
  return [ref.range, crit];
}

function canUseIfsPairs(ctx) {
  const raw = ctx?.resolved?.filterColumns?.length
    ? ctx.resolved.filterColumns
    : Array.isArray(ctx?.intent?.filters)
      ? ctx.intent.filters
      : Array.isArray(ctx?.intent?.conditions)
        ? ctx.intent.conditions
        : [];

  const ok = raw.every((c) => _isLeafIfsSafe(c));
  if (!ok) return false;

  // window(days)는 AND 단일 범위 비교이므로 허용
  return true;
}

function buildConditionPairs(ctx, formatValue) {
  if (!canUseIfsPairs(ctx)) return [];

  const raw = ctx?.resolved?.filterColumns?.length
    ? ctx.resolved.filterColumns
    : Array.isArray(ctx?.intent?.filters)
      ? ctx.intent.filters
      : Array.isArray(ctx?.intent?.conditions)
        ? ctx.intent.conditions
        : [];

  const out = [];
  for (const c of raw) {
    const pair = _leafToIfsPair(c, ctx, formatValue);
    if (!pair) return [];
    out.push(...pair);
  }

  const win = _windowExprs(ctx);
  if (win.pairs.length) out.push(...win.pairs);
  return out;
}

module.exports = {
  buildConditionMask,
  buildSingleConditionExpr,
  buildConditionPairs,
  canUseIfsPairs,
};
