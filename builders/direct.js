const { parseExplicitCellOrRange } = require("../utils/formulaUtils");

function getRawForExplicit(it = {}, ctx = {}) {
  return String(
    it?.raw_message ||
      ctx?.prompt ||
      ctx?.message ||
      ctx?.nl ||
      ctx?.text ||
      "",
  );
}

function hasExplicitRangeIntent(it = {}, ctx = {}) {
  const raw = getRawForExplicit(it, ctx);
  if (it?.range || it?.target_cell) return true;
  return !!parseExplicitCellOrRange(raw);
}

function hasHeaderDrivenIntent(it = {}) {
  return Boolean(
    it?.header_hint ||
    it?.return_hint ||
    it?.lookup_hint ||
    it?.group_by ||
    (Array.isArray(it?.return_fields) && it.return_fields.length) ||
    (Array.isArray(it?.filters) && it.filters.length) ||
    (Array.isArray(it?.conditions) && it.conditions.length) ||
    it?.lookup?.key_header,
  );
}

function shouldRejectDirectForUploadedSheet(ctx) {
  const it = ctx?.intent || {};
  const hasSheetsMeta =
    !!ctx?.allSheetsData && Object.keys(ctx.allSheetsData).length > 0;

  if (!hasSheetsMeta) return false;
  if (!hasExplicitRangeIntent(it, ctx)) return true;
  if (hasHeaderDrivenIntent(it)) return true;

  return false;
}

// === 유틸 ===
function isCellRef(v) {
  return /^[A-Z]{1,3}[0-9]{1,7}$/i.test(String(v || "").trim());
}

function getRangeOrCell(it, ctx = {}) {
  if (it?.range) return it.range;
  if (it?.target_cell) return it.target_cell;
  const guessed = parseExplicitCellOrRange(getRawForExplicit(it, ctx));
  return guessed || null;
}

function wrap(v) {
  if (v == null) return '""';
  if (typeof v === "number") return String(v);
  if (typeof v === "string") {
    if (isCellRef(v)) return v.trim();
    if (/^".*"$/.test(v)) return v;
    return `"${v}"`;
  }
  if (v?.cell) return v.cell;
  return '""';
}

function needRangeError(funcName) {
  return `=ERROR("${funcName.toUpperCase()}: 범위를 지정해 주세요 (예: A1:A10)")`;
}

// === 집계 공용 ===
function simpleAggregate(it, funcName, ctx) {
  if (ctx && shouldRejectDirectForUploadedSheet(ctx)) {
    return `=ERROR("${funcName.toUpperCase()}: 업로드 파일 기반 요청은 헤더/조건 엔진으로 처리해야 합니다.")`;
  }

  const r = getRangeOrCell(it, ctx);
  if (!r) return needRangeError(funcName);
  return `=${funcName.toUpperCase()}(${r})`;
}

function average(
  it,
  _formatValue,
  _buildConditionPairs,
  _buildConditionMask,
  ctx,
) {
  return simpleAggregate(it, "AVERAGE", ctx);
}

function sum(it, _formatValue, _buildConditionPairs, _buildConditionMask, ctx) {
  return simpleAggregate(it, "SUM", ctx);
}

function count(
  it,
  _formatValue,
  _buildConditionPairs,
  _buildConditionMask,
  ctx,
) {
  return simpleAggregate(it, "COUNT", ctx);
}

function minf(
  it,
  _formatValue,
  _buildConditionPairs,
  _buildConditionMask,
  ctx,
) {
  return simpleAggregate(it, "MIN", ctx);
}

function maxf(
  it,
  _formatValue,
  _buildConditionPairs,
  _buildConditionMask,
  ctx,
) {
  return simpleAggregate(it, "MAX", ctx);
}

function iff(it) {
  const cond = it?.condition || {};
  const target =
    cond.target_cell ||
    (cond.target && cond.target.cell) ||
    cond.target || // 문자열 "A1" 허용
    null;

  const op = (cond.operator || "=").trim();
  let val;
  if (cond.value && typeof cond.value === "object" && cond.value.cell) {
    val = cond.value.cell;
  } else if (typeof cond.value === "string") {
    val = wrap(cond.value);
  } else if (
    typeof cond.value === "number" ||
    typeof cond.value === "boolean"
  ) {
    val = wrap(cond.value);
  } else {
    // 값이 비었을 때도 허용: 공백 문자열
    val = '""';
  }

  const vt = wrap(it?.value_if_true);
  const vf = wrap(it?.value_if_false);

  if (!target) return `=ERROR("IF: 비교 대상 셀(예: A1)이 필요합니다")`;
  return `=IF(${target}${op}${val},${vt},${vf})`;
}

function textjoin(it) {
  const delim = it?.delimiter != null ? String(it.delimiter) : ",";
  const ignoreEmpty = !!it?.ignore_empty; // 기본 FALSE
  const args = (it?.values || [])
    .map((v) => (v && typeof v === "object" && v.cell ? v.cell : wrap(v)))
    .filter(Boolean);

  if (!args.length) return `=ERROR("TEXTJOIN: 인수가 없습니다")`;
  return `=TEXTJOIN(${wrap(delim)},${
    ignoreEmpty ? "TRUE" : "FALSE"
  },${args.join(",")})`;
}

function concat(it) {
  const a = it?.a?.cell || it?.a || null;
  const b = it?.b?.cell || it?.b || null;
  if (!a || !b) return `=ERROR("CONCAT: a, b가 필요합니다")`;
  return `=${it?.use_amp ? `${a}&${b}` : `CONCAT(${a},${b})`}`;
}

function left(it) {
  const c = it?.target_cell;
  const n = it?.num_chars ?? 1;
  if (!c) return `=ERROR("LEFT: 대상 셀")`;
  return `=LEFT(${c},${n})`;
}

function right(it) {
  const c = it?.target_cell;
  const n = it?.num_chars ?? 1;
  if (!c) return `=ERROR("RIGHT: 대상 셀")`;
  return `=RIGHT(${c},${n})`;
}

function len(it) {
  const c = it?.target_cell;
  if (!c) return `=ERROR("LEN: 대상 셀")`;
  return `=LEN(${c})`;
}

// === 날짜/시간 ===
function today() {
  return "=TODAY()";
}

function now() {
  return "=NOW()";
}

function year(it) {
  const c = it?.target_cell;
  if (!c) return `=ERROR("YEAR: 대상 셀")`;
  return `=YEAR(${c})`;
}

function month(it) {
  const c = it?.target_cell;
  if (!c) return `=ERROR("MONTH: 대상 셀")`;
  return `=MONTH(${c})`;
}

function day(it) {
  const c = it?.target_cell;
  if (!c) return `=ERROR("DAY: 대상 셀")`;
  return `=DAY(${c})`;
}

// === 숫자/반올림 ===
function round(it) {
  const c = it?.target_cell || it?.value?.cell || it?.value;
  const n = it?.num_digits ?? 0;
  if (!c) return `=ERROR("ROUND: 대상")`;
  return `=ROUND(${c},${n})`;
}

function roundup(it) {
  const c = it?.target_cell || it?.value?.cell || it?.value;
  const n = it?.num_digits ?? 0;
  if (!c) return `=ERROR("ROUNDUP: 대상")`;
  return `=ROUNDUP(${c},${n})`;
}

function rounddown(it) {
  const c = it?.target_cell || it?.value?.cell || it?.value;
  const n = it?.num_digits ?? 0;
  if (!c) return `=ERROR("ROUNDDOWN: 대상")`;
  return `=ROUNDDOWN(${c},${n})`;
}

function abs(it) {
  const c = it?.target_cell || it?.value?.cell || it?.value;
  if (!c) return `=ERROR("ABS: 대상")`;
  return `=ABS(${c})`;
}

function intf(it) {
  const c = it?.target_cell || it?.value?.cell || it?.value;
  if (!c) return `=ERROR("INT: 대상")`;
  return `=INT(${c})`;
}

function rand() {
  return "=RAND()";
}

function randbetween(it) {
  const a = it?.min ?? 0,
    b = it?.max ?? 100;
  return `=RANDBETWEEN(${a},${b})`;
}

// === 텍스트 보조 ===
function upper(it) {
  const c = it?.target_cell || it?.text?.cell || wrap(it?.text);
  if (!c) return `=ERROR("UPPER: 대상")`;
  return `=UPPER(${c})`;
}

function lower(it) {
  const c = it?.target_cell || it?.text?.cell || wrap(it?.text);
  if (!c) return `=ERROR("LOWER: 대상")`;
  return `=LOWER(${c})`;
}

function trimf(it) {
  const c = it?.target_cell || it?.text?.cell || wrap(it?.text);
  if (!c) return `=ERROR("TRIM: 대상")`;
  return `=TRIM(${c})`;
}

function mid(it) {
  const c = it?.target_cell;
  const start = it?.start || 1;
  const num = it?.num_chars || 1;
  if (!c) return `=ERROR("MID: 대상 셀")`;
  return `=MID(${c},${start},${num})`;
}

function substitute(it) {
  const c = it?.target_cell;
  const oldv = wrap(it?.old_text);
  const newv = wrap(it?.new_text);
  const inst = it?.instance_num;
  if (!c) return `=ERROR("SUBSTITUTE: 대상 셀")`;
  return inst
    ? `=SUBSTITUTE(${c},${oldv},${newv},${inst})`
    : `=SUBSTITUTE(${c},${oldv},${newv})`;
}

function replacef(it) {
  const c = it?.target_cell;
  const start = it?.start || 1;
  const num = it?.num_chars || 1;
  const newt = wrap(it?.new_text);
  if (!c) return `=ERROR("REPLACE: 대상 셀")`;
  return `=REPLACE(${c},${start},${num},${newt})`;
}

// === 찾기 ===
function findf(it) {
  const find = wrap(it?.find_text);
  const within = it?.within?.cell || it?.within;
  const start = it?.start_num || 1;
  if (!find || !within) return `=ERROR("FIND: 인수 부족")`;
  return `=FIND(${find},${within},${start})`;
}

function searchf(it) {
  const find = wrap(it?.find_text);
  const within = it?.within?.cell || it?.within;
  const start = it?.start_num || 1;
  if (!find || !within) return `=ERROR("SEARCH: 인수 부족")`;
  return `=SEARCH(${find},${within},${start})`;
}

function vstack(it, ctx) {
  const ranges =
    Array.isArray(it?.ranges) && it.ranges.length
      ? it.ranges
      : (() => {
          const raw = getRawForExplicit(it, ctx).toUpperCase();
          const tokens =
            raw.match(/[A-Z]+[0-9]+:[A-Z]+[0-9]+|[A-Z]+:[A-Z]+/g) || [];
          return tokens;
        })();

  if (!ranges || ranges.length < 2) {
    return `=ERROR("VSTACK: 두 개 이상의 범위가 필요합니다.")`;
  }
  return `=VSTACK(${ranges.join(", ")})`;
}

function tocol(it, ctx) {
  const r = getRangeOrCell(it, ctx);
  if (!r) return `=ERROR("TOCOL: 범위를 지정해 주세요 (예: A1:C3)")`;
  return `=TOCOL(${r}, 0, 0)`;
}

function byrow(it, ctx) {
  const r = getRangeOrCell(it, ctx);
  if (!r) return `=ERROR("BYROW: 범위를 지정해 주세요 (예: A1:C3)")`;
  const agg = String(it?.aggregate || "sum").toLowerCase();
  const body =
    agg === "average"
      ? "AVERAGE(r)"
      : agg === "max"
        ? "MAX(r)"
        : agg === "min"
          ? "MIN(r)"
          : "SUM(r)";
  return `=BYROW(${r}, LAMBDA(r, ${body}))`;
}

// === 디스패처 ===
const handlers = {
  average,
  sum,
  count,
  min: minf,
  max: maxf,

  textjoin,
  concat,
  left,
  right,
  len,

  if: iff,

  today,
  now,
  year,
  month,
  day,

  round: round,
  roundup: roundup,
  rounddown: rounddown,
  abs: abs,
  int: intf,
  rand: rand,
  randbetween: randbetween,

  upper: upper,
  lower: lower,
  trim: trimf,
  mid: mid,
  substitute: substitute,
  replace: replacef,

  find: findf,
  search: searchf,
  vstack,
  tocol,
  byrow,
};

function canHandleWithoutFile(intent) {
  const op = String(intent?.operation || "").toLowerCase();
  return !!handlers[op];
}

function buildFormula(intent, ctx = {}) {
  const op = String(intent?.operation || "").toLowerCase();
  const h = handlers[op];
  if (!h) return null;
  return h(intent, null, null, null, ctx);
}

function formula(ctx) {
  const it = ctx.intent || {};
  const raw =
    it.raw_formula || it.formula || it.expression || it.raw || it.text || "";
  const s = String(raw || "").trim();
  if (!s) return '=ERROR("DIRECT: 수식을 제공해 주세요")';
  const eq = s.startsWith("=") ? s : "=" + s;

  // 가벼운 안전망: 엔진-전용 함수가 반대 엔진에 들어왔을 때 알림
  if (
    ctx.engine === "excel" &&
    /\b(IMPORTRANGE|IMPORTHTML|IMPORTXML|IMPORTDATA|GOOGLEFINANCE|REGEXMATCH|REGEXEXTRACT|REGEXREPLACE)\b/i.test(
      eq,
    )
  ) {
    return '=ERROR("이 수식은 Google Sheets 전용 함수가 포함되어 Excel 엔진에서 지원되지 않습니다.")';
  }
  // 정책 래핑(NA/ERROR) 제거: 현재는 사용자 수식 그대로 전달
  return eq;
}

function shouldUseDirectBuilder(intent = {}, ctx = {}) {
  const it = intent || {};
  const hasSheetsMeta =
    !!ctx?.allSheetsData && Object.keys(ctx.allSheetsData).length > 0;

  const explicit = hasExplicitRangeIntent(it, ctx);
  const headerDriven = hasHeaderDrivenIntent(it);

  if (!explicit) return false;
  if (hasSheetsMeta && headerDriven) return false;
  return true;
}

module.exports = {
  shouldUseDirectBuilder,
  shouldRejectDirectForUploadedSheet,
  average: (ctx) => average(ctx.intent || {}, null, null, null, ctx),
  sum: (ctx) => sum(ctx.intent || {}, null, null, null, ctx),
  count: (ctx) => count(ctx.intent || {}, null, null, null, ctx),
  min: (ctx) => minf(ctx.intent || {}, null, null, null, ctx),
  max: (ctx) => maxf(ctx.intent || {}, null, null, null, ctx),
  if: (ctx) => iff(ctx.intent || {}),
  textjoin: (ctx) => textjoin(ctx.intent || {}),
  concat: (ctx) => concat(ctx.intent || {}),
  left: (ctx) => left(ctx.intent || {}),
  right: (ctx) => right(ctx.intent || {}),
  len: (ctx) => len(ctx.intent || {}),
  today,
  now,
  year: (ctx) => year(ctx.intent || {}),
  month: (ctx) => month(ctx.intent || {}),
  day: (ctx) => day(ctx.intent || {}),
  vstack: (ctx) => vstack(ctx.intent || {}, ctx),
  tocol: (ctx) => tocol(ctx.intent || {}, ctx),
  byrow: (ctx) => byrow(ctx.intent || {}, ctx),

  canHandleWithoutFile,
  buildFormula,
  formula,
};
