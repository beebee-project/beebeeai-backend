const {
  refFromHeaderSpec,
  evalSubIntentToScalar,
} = require("../utils/builderHelpers");

/* =========================================================
 * Date/Time 공통 헬퍼 (병합 고도화판)
 * - 헤더/시트 참조, 단일/벡터 라우팅, 서브-의도 스칼라
 * - 휴일 범위, row_selector 보조, week_start/주말 프리셋
 * - IFERROR/IFNA/ctx.util.wrapIfError 계열 전면 제거
 * =======================================================*/

// 기준일(+서브수식 가능)에서 ±N일 이동 표준화
// mode: "calendar" | "workday" | "eomonth"
// holidays: '시트'!A2:A or {header,sheet} or "시트!A2:A"
// weekend_mask: "0000011"(토/일) 등 INTL 마스크(선택)
function date_relative(
  ctx,
  base,
  offsetDays = 0,
  mode = "calendar",
  holidays,
  weekend_mask
) {
  const b = _scalarFromDate(base, ctx, (v) =>
    ctx.formulaBuilder._formatValue(v)
  );
  const n = Number(offsetDays || 0);
  if (!b) return `ERROR("date_relative: 기준일 해석 실패")`;

  if (mode === "workday") {
    const hol = _holidaysRange({ holidays }, ctx);
    const mask = _resolveWeekendMask(weekend_mask);
    if (mask) {
      return hol
        ? `WORKDAY.INTL(${b}, ${n}, "${mask}", ${hol})`
        : `WORKDAY.INTL(${b}, ${n}, "${mask}")`;
    }
    return hol ? `WORKDAY(${b}, ${n}, ${hol})` : `WORKDAY(${b}, ${n})`;
  }
  if (mode === "eomonth") {
    return `EOMONTH(${b}, ${n})`;
  }
  // calendar(영업일 고려 없음)
  if (n === 0) return `${b}`;
  return `(${b})+${n}`;
}

function _scalarFromDate(spec, ctx, formatValue = (x) => JSON.stringify(x)) {
  if (!spec) return null;
  if (typeof spec === "object") {
    if (spec.operation) return evalSubIntentToScalar(ctx, formatValue, spec);
    if (spec.header) {
      const r = refFromHeaderSpec(ctx, spec);
      return r ? r.cell : null;
    }
    if (spec.cell) return spec.cell;
  }
  const s = String(spec).trim();
  if (/^(TODAY|NOW)\(\)$/i.test(s)) return s.toUpperCase();
  if (
    /^(DATE|EOMONTH|EDATE|WORKDAY|WORKDAY\.INTL|NETWORKDAYS|NETWORKDAYS\.INTL)\(/i.test(
      s
    )
  )
    return s;
  // 로캘 독립 날짜: YYYY.MM.DD / YYYY/MM/DD / YYYY-MM-DD → DATEVALUE("…")
  if (/^\d{4}([\-./])\d{1,2}\1\d{1,2}$/.test(s)) {
    const iso = s.replace(/[./]/g, "-");
    return `DATEVALUE(${formatValue(iso)})`;
  }
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(s)) return formatValue(s);
  const r = refFromHeaderSpec(ctx, s);
  if (r) return r.cell;
  return formatValue(s);
}

/** 대상 셀/범위 결정: scope="all"이면 range, 아니면 단일 cell */
function _targetCellOrRange(ctx) {
  const it = ctx.intent || {};
  const hint = it.date_header || it.header_hint;
  if (it.scope === "all") {
    const r = hint
      ? refFromHeaderSpec(ctx, hint)
      : ctx.bestReturn
      ? refFromHeaderSpec(
          ctx,
          ctx.bestReturn.header ||
            ctx.bestReturn.header_hint ||
            ctx.bestReturn.columnLetter
        )
      : null;
    if (r) return { isRange: true, range: r.range, sheetName: r.sheetName };
  }
  if (hint) {
    const r = refFromHeaderSpec(ctx, hint);
    if (r) return { isRange: false, cell: r.cell, sheetName: r.sheetName };
  }
  if (ctx.bestReturn) {
    const br = ctx.bestReturn;
    return {
      isRange: false,
      cell: `'${br.sheetName}'!${br.columnLetter}${br.startRow}`,
      sheetName: br.sheetName,
    };
  }
  return null;
}

/** 휴일 범위: '시트'!A2:A / {header,sheet} / "시트!A2:A" 지원 */
function _holidaysRange(it, ctx) {
  if (!it || !it.holidays) return null;
  if (typeof it.holidays === "string") {
    if (/!/.test(it.holidays) && /:/.test(it.holidays)) return it.holidays;
    const r = refFromHeaderSpec(ctx, it.holidays);
    return r ? r.range : it.holidays;
  }
  if (typeof it.holidays === "object") {
    const r = refFromHeaderSpec(ctx, it.holidays);
    return r ? r.range : null;
  }
  return null;
}

/** row_selector 지원: 키열(힌트)에서 value 매칭 → 단일 셀 반환 */
function _selectCellByRowSelector(ctx, headerSpec) {
  const it = ctx.intent || {};
  if (!it.row_selector?.hint || it.row_selector.value == null) return null;
  const keyCol = refFromHeaderSpec(ctx, it.row_selector.hint);
  const tgtCol = refFromHeaderSpec(
    ctx,
    headerSpec || it.date_header || it.header_hint
  );
  if (!keyCol || !tgtCol) return null;
  const keyVal =
    typeof it.row_selector.value === "string"
      ? JSON.stringify(it.row_selector.value)
      : it.row_selector.value;
  return `XLOOKUP(${keyVal}, ${keyCol.range}, ${tgtCol.range})`;
}

function _routeDateUnary(t, innerFromD, rowPickExpr) {
  if (rowPickExpr) return `=${innerFromD(rowPickExpr)}`; // 특정 행 지정 시 단일 처리
  if (!t) return `=ERROR("날짜 열을 찾을 수 없습니다.")`;
  if (t.isRange) {
    const inner = innerFromD("d");
    return `=BYROW(${t.range}, LAMBDA(d, ${inner}))`;
  }
  return `=${innerFromD(t.cell)}`;
}

/* ===== Week/Weekend Presets & Mappers ===== */
function _mapWeekStartOptions(intent) {
  const raw = String(
    intent?.week_start || intent?.return_type || ""
  ).toLowerCase();
  if (raw === "iso")
    return { weekdayRT: 2, weeknumFN: "ISOWEEKNUM", weeknumRT: null };
  const idx = { sun: 1, mon: 2, tue: 3, wed: 4, thu: 5, fri: 6, sat: 7 }[raw];
  const weekdayRT =
    raw === "mon" ? 2 : raw === "sun" ? 1 : intent?.return_type ?? 2;
  let weeknumFN = "WEEKNUM";
  let weeknumRT = intent?.return_type ?? 21;
  if (idx) {
    weeknumRT =
      raw === "mon"
        ? 11
        : raw === "tue"
        ? 12
        : raw === "wed"
        ? 13
        : raw === "thu"
        ? 14
        : raw === "fri"
        ? 15
        : raw === "sat"
        ? 16
        : raw === "sun"
        ? 17
        : 21;
  }
  return { weekdayRT, weeknumFN, weeknumRT };
}

// INTL weekend mask 프리셋 (월화수목금토일; 1=휴무)
function _resolveWeekendMask(maskOrPreset) {
  if (!maskOrPreset) return null;
  const s = String(maskOrPreset).trim();
  if (/^[01]{7}$/.test(s)) return s;
  const key = s.toLowerCase();
  const presets = {
    sat_sun: "0000011",
    fri_sat: "0000110",
    thu_fri: "0001100",
  };
  return presets[key] || null;
}

/* ===== Row-pick 우선 앵커 선택 ===== */
function _anchorOrPick(
  it,
  ctx,
  formatValue,
  { fallbackToday = true, headerForPick = null } = {}
) {
  const a = _scalarFromDate(it?.anchor_date, ctx, formatValue);
  if (a) return a;
  const picked = _selectCellByRowSelector(
    ctx,
    headerForPick || it?.date_header || it?.header_hint
  );
  if (picked) return picked;
  const s = _scalarFromDate(it?.start_date, ctx, formatValue);
  if (s) return s;
  return fallbackToday ? "TODAY()" : null;
}

const dateFunctionBuilder = {
  today: () => `=TODAY()`,
  now: () => `=NOW()`,

  year(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    return _routeDateUnary(t, (d) => `YEAR(${d})`, pick);
  },
  month(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    return _routeDateUnary(t, (d) => `MONTH(${d})`, pick);
  },
  day(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    return _routeDateUnary(t, (d) => `DAY(${d})`, pick);
  },
  hour(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    return _routeDateUnary(t, (d) => `HOUR(${d})`, pick);
  },
  minute(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    return _routeDateUnary(t, (d) => `MINUTE(${d})`, pick);
  },
  second(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    return _routeDateUnary(t, (d) => `SECOND(${d})`, pick);
  },

  weekday(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    const { weekdayRT } = _mapWeekStartOptions(ctx.intent);
    const rt = weekdayRT ?? ctx.intent?.return_type ?? 2;
    return _routeDateUnary(t, (d) => `WEEKDAY(${d}, ${rt})`, pick);
  },
  weeknum(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    const { weeknumFN, weeknumRT } = _mapWeekStartOptions(ctx.intent);
    const fn = weeknumFN || "WEEKNUM";
    const rt = fn === "ISOWEEKNUM" ? "" : `, ${weeknumRT ?? 21}`;
    return _routeDateUnary(t, (d) => `${fn}(${d}${rt})`, pick);
  },

  edate(ctx, formatValue) {
    const it = ctx.intent || {};
    const anchor = _anchorOrPick(it, ctx, formatValue, {
      headerForPick: it.date_header,
    });
    const m = it.months ?? 1;
    return `=EDATE(${anchor}, ${m})`;
  },
  eomonth(ctx, formatValue) {
    const it = ctx.intent || {};
    const anchor = _anchorOrPick(it, ctx, formatValue, {
      headerForPick: it.date_header,
      fallbackToday: true,
    });
    const m = it.months ?? 0;
    const body = date_relative(ctx, anchor, m, "eomonth");
    return `=${body}`;
  },

  workday(ctx, formatValue) {
    const it = ctx.intent || {};
    const start = _anchorOrPick(it, ctx, formatValue, {
      headerForPick: it.date_header,
    });
    const days = it.workdays ?? it.days ?? 1;
    const mask =
      _resolveWeekendMask(it.weekend_preset || it.weekend_mask) ||
      it.weekend_mask;
    const hol = _holidaysRange(it, ctx);
    const body = date_relative(
      ctx,
      start,
      days,
      "workday",
      hol || it.holidays,
      mask
    );
    return `=${body}`;
  },
  networkdays(ctx, formatValue) {
    const it = ctx.intent || {};
    const s =
      _scalarFromDate(it.start_date, ctx, formatValue) ||
      _anchorOrPick({ ...it, anchor_date: it.start_date }, ctx, formatValue, {
        headerForPick: it.date_header,
      });
    const e =
      _scalarFromDate(it.end_date, ctx, formatValue) ||
      _anchorOrPick({ ...it, anchor_date: it.end_date }, ctx, formatValue, {
        headerForPick: it.date_header,
      });
    const mask =
      _resolveWeekendMask(it.weekend_preset || it.weekend_mask) ||
      it.weekend_mask;
    const intl = mask ? `, "${mask}"` : "";
    const hol = _holidaysRange(it, ctx);
    return mask
      ? `=NETWORKDAYS.INTL(${s}, ${e}${intl}${hol ? `, ${hol}` : ""})`
      : `=NETWORKDAYS(${s}, ${e}${hol ? `, ${hol}` : ""})`;
  },

  datedif(ctx, formatValue) {
    const it = ctx.intent || {};
    const s =
      _scalarFromDate(it.start_date, ctx, formatValue) ||
      _anchorOrPick({ ...it, anchor_date: it.start_date }, ctx, formatValue, {
        headerForPick: it.date_header,
      });
    const e =
      _scalarFromDate(it.end_date, ctx, formatValue) ||
      _anchorOrPick({ ...it, anchor_date: it.end_date }, ctx, formatValue, {
        headerForPick: it.date_header,
      });
    const unit = (it.unit || "D").toUpperCase();
    return `=DATEDIF(${s}, ${e}, "${unit}")`;
  },
  date(ctx) {
    const it = ctx.intent || {};
    const y = it.year ?? "YEAR(TODAY())";
    const m = it.month ?? "MONTH(TODAY())";
    const d = it.day ?? 1;
    return `=DATE(${y}, ${m}, ${d})`;
  },
  datevalue(ctx, formatValue) {
    const raw = ctx.intent?.text || "2025-01-01";
    const iso = String(raw).replace(/[./]/g, "-");
    return `=DATEVALUE(${formatValue(iso)})`;
  },
  time(ctx) {
    const it = ctx.intent || {};
    return `=TIME(${it.hour ?? 0}, ${it.minute ?? 0}, ${it.second ?? 0})`;
  },
  timevalue(ctx, formatValue) {
    const s = ctx.intent?.text || "09:00:00";
    return `=TIMEVALUE(${formatValue(s)})`;
  },

  week_of_month(ctx) {
    const t = _targetCellOrRange(ctx);
    const pick = _selectCellByRowSelector(ctx);
    const { weekdayRT } = _mapWeekStartOptions(ctx.intent);
    const weekdayBase = weekdayRT === 1 ? 1 : 2;
    const inner = (d) => {
      const firstOfMonth = `DATE(YEAR(${d}), MONTH(${d}), 1)`;
      const wd = `WEEKDAY(${firstOfMonth}, ${weekdayBase})`;
      return `ROUNDUP((DAY(${d}) + ${wd} - 1) / 7, 0)`;
    };
    return _routeDateUnary(t, inner, pick);
  },

  _date_relative: function (ctx, formatValue) {
    const it = ctx.intent || {};
    const { base, offset_days, mode, holidays, weekend_mask } = it;
    const expr = date_relative(
      ctx,
      base || "TODAY()",
      offset_days || 0,
      mode || "calendar",
      holidays,
      weekend_mask
    );
    const body = expr.startsWith("=") ? expr.slice(1) : expr;
    return `=${body}`;
  },
};

module.exports = dateFunctionBuilder;
module.exports.date_relative = date_relative;
