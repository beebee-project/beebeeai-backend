const {
  refFromHeaderSpec,
  rangeFromSpec,
  evalSubIntentToScalar,
} = require("../utils/builderHelpers");

/* 입력이 범위/스칼라인지에 따라 BYROW 또는 단일식으로 라우팅 */
function _asUnaryByRowOrScalar(ctx, formatValue, inputSpec, gen) {
  const { isRange, expr } = _toRangeOrScalarExpr(ctx, formatValue, inputSpec);
  if (!expr) {
    // IFERROR 사용 금지 → 명시적 오류만 반환
    return `=ERROR("텍스트 입력이 없습니다.")`;
  }
  return isRange ? `=BYROW(${expr}, LAMBDA(x, ${gen("x")}))` : `=${gen(expr)}`;
}

/* 서브-의도 → 범위 수식 (예: FILTER, 반환이 벡터/범위인 경우) */
function _evalSubExprToRange(ctx, formatValue, node) {
  if (!node || typeof node !== "object" || !node.operation) return null;
  const fb = ctx.formulaBuilder;
  if (fb && typeof fb[node.operation] === "function") {
    const res = fb[node.operation].call(
      fb,
      { ...ctx, intent: node },
      formatValue,
      fb._buildConditionPairs
    );
    if (typeof res === "string" && res.startsWith("=")) return res.slice(1);
    return res;
  }
  return null;
}

/* 대상 셀/범위 결정: scope="all"이면 range, 아니면 단일 cell */
function _targetCellOrRange(ctx, kind = "text") {
  const it = ctx.intent || {};
  const hint = it.target_header || it.header_hint;
  if (it.scope === "all") {
    const r = hint
      ? refFromHeaderSpec(hint, ctx)
      : ctx.bestReturn
      ? refFromHeaderSpec(
          ctx.bestReturn.header ||
            ctx.bestReturn.header_hint ||
            ctx.bestReturn.columnLetter,
          ctx
        )
      : null;
    if (r) return { isRange: true, range: r.range, sheetName: r.sheetName };
  }
  if (hint) {
    const r = refFromHeaderSpec(hint, ctx);
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
  if (it.target_cell)
    return { isRange: false, cell: it.target_cell, sheetName: null };
  return null;
}

/* row_selector: 키열/대상열 지정으로 한 셀 픽업 (에러래핑 없음) */
function _selectCellByRowSelector(ctx, headerSpec) {
  const it = ctx.intent || {};
  if (!it.row_selector?.hint || it.row_selector.value == null) return null;
  const keyCol = refFromHeaderSpec(it.row_selector.hint, ctx);
  const tgtCol = refFromHeaderSpec(
    headerSpec || it.target_header || it.header_hint,
    ctx
  );
  if (!keyCol || !tgtCol) return null;
  const keyVal =
    typeof it.row_selector.value === "string"
      ? JSON.stringify(it.row_selector.value)
      : it.row_selector.value;
  return `XLOOKUP(${keyVal}, ${keyCol.range}, ${tgtCol.range})`;
}

/* 단항 텍스트 공통 라우팅 */
function _routeTextUnary(
  t,
  innerFromX /* (xSym)=>"FUNC(xSym)" */,
  rowPickExpr /* optional */
) {
  if (rowPickExpr) return `=${innerFromX(rowPickExpr)}`;
  if (!t) return `=ERROR("텍스트 대상 열을 찾을 수 없습니다.")`;
  if (t.isRange) return `=BYROW(${t.range}, LAMBDA(x, ${innerFromX("x")}))`;
  return `=${innerFromX(t.cell)}`;
}

// 간단 정규식 우회 (정교한 정규식은 명시 에러로)
function _regexLiteFallbackExpr(kind, textExpr, patternExpr, formatValue) {
  const m = String(patternExpr).match(/^"(.*)"$/);
  if (!m) return null;
  const pat = m[1];

  // ^prefix
  if (pat.startsWith("^") && !pat.includes("|") && !pat.includes("[")) {
    const prefix = pat.slice(1);
    if (kind === "match")
      return `=LEFT(${textExpr}, LEN(${formatValue(prefix)})) = ${formatValue(
        prefix
      )}`;
    if (kind === "extract")
      return `=IF(LEFT(${textExpr}, LEN(${formatValue(prefix)}))=${formatValue(
        prefix
      )}, ${formatValue(prefix)}, "")`;
    if (kind === "replace")
      return `=IF(LEFT(${textExpr}, LEN(${formatValue(prefix)}))=${formatValue(
        prefix
      )}, "" & MID(${textExpr}, LEN(${formatValue(
        prefix
      )})+1, 10^6), ${textExpr})`;
  }

  // suffix$
  if (pat.endsWith("$") && !pat.includes("|") && !pat.includes("[")) {
    const suffix = pat.slice(0, -1);
    if (kind === "match")
      return `=RIGHT(${textExpr}, LEN(${formatValue(suffix)})) = ${formatValue(
        suffix
      )}`;
    if (kind === "extract")
      return `=IF(RIGHT(${textExpr}, LEN(${formatValue(suffix)}))=${formatValue(
        suffix
      )}, ${formatValue(suffix)}, "")`;
    if (kind === "replace")
      return `=IF(RIGHT(${textExpr}, LEN(${formatValue(suffix)}))=${formatValue(
        suffix
      )}, LEFT(${textExpr}, LEN(${textExpr})-LEN(${formatValue(
        suffix
      )})), ${textExpr})`;
  }

  // 단순 포함
  if (!/[.^$*+?()[\]{}\\|]/.test(pat)) {
    if (kind === "match")
      return `=ISNUMBER(SEARCH(${formatValue(pat)}, ${textExpr}))`;
    if (kind === "extract") {
      return `=IF(ISNUMBER(SEARCH(${formatValue(
        pat
      )}, ${textExpr})), ${formatValue(pat)}, "")`;
    }
    if (kind === "replace") {
      return `=SUBSTITUTE(${textExpr}, ${formatValue(pat)}, "")`;
    }
  }

  // 예시: 전체 숫자 ^\d+$ (간이)
  if (/^\\d\+\$/.test(pat) || /^\^\[0-9\]\+\$$/.test(pat)) {
    if (kind === "match") return `=ISNUMBER(--${textExpr})`;
    if (kind === "extract")
      return `=IF(ISNUMBER(--${textExpr}), ${textExpr}, "")`;
    if (kind === "replace")
      return `=IF(ISNUMBER(--${textExpr}), "", ${textExpr})`;
  }

  // 예시: 첫 하이픈 분리 ^(.+?)-(.*)$
  if (pat === "^(.+?)-(.*)$") {
    const pos = `FIND("-", ${textExpr})`;
    if (kind === "match") return `=ISNUMBER(${pos})`;
    if (kind === "extract")
      return `=IF(ISNUMBER(${pos}), LEFT(${textExpr}, ${pos}-1) & "-" & MID(${textExpr}, ${pos}+1, 10^6), "")`;
    if (kind === "replace")
      return `=IF(ISNUMBER(${pos}), LEFT(${textExpr}, ${pos}-1) & MID(${textExpr}, ${pos}+1, 10^6), ${textExpr})`;
  }

  return null;
}

// 다중 구분자 → Excel TEXTSPLIT용 배열 상수 {"-","/"} 생성
function _excelArrayConst(items, formatValue) {
  const arr = (Array.isArray(items) ? items : [items])
    .filter((v) => v != null)
    .map((v) => {
      if (typeof v === "object") return null;
      return String(v);
    })
    .filter((v) => v != null);
  if (arr.length === 0) return null;
  return `{${arr.map((s) => `"${s.replace(/"/g, '""')}"`).join(",")}}`;
}

// spec → 가능한 한 스칼라 수식으로
function _toScalarExpr(ctx, formatValue, spec) {
  if (!spec && spec !== 0) return null;
  if (typeof spec === "object" && spec.operation) {
    return evalSubIntentToScalar(ctx, formatValue, spec);
  }
  if (typeof spec === "string") {
    if (/!/.test(spec) && /:/.test(spec)) return `INDEX(${spec}, 1)`;
    const r = rangeFromSpec(ctx, spec);
    if (r) return `INDEX(${r}, 1)`;
    return formatValue(spec);
  }
  if (typeof spec === "object" && (spec.header || spec.sheet)) {
    const r = rangeFromSpec(ctx, spec);
    return r ? `INDEX(${r}, 1)` : null;
  }
  return formatValue(spec);
}

// spec → 범위면 그 범위, 아니면 스칼라
function _toRangeOrScalarExpr(ctx, formatValue, spec) {
  if (!spec && spec !== 0) return { isRange: false, expr: null };
  if (typeof spec === "object" && spec.operation) {
    const asRange = _evalSubExprToRange(ctx, formatValue, spec);
    if (asRange) return { isRange: true, expr: asRange };
    const asScalar = evalSubIntentToScalar(ctx, formatValue, spec);
    return { isRange: false, expr: asScalar };
  }
  if (typeof spec === "string") {
    if (/!/.test(spec) && /:/.test(spec)) return { isRange: true, expr: spec };
    const r = rangeFromSpec(ctx, spec);
    if (r) return { isRange: true, expr: r };
    return { isRange: false, expr: formatValue(spec) };
  }
  if (typeof spec === "object" && (spec.header || spec.sheet)) {
    const r = rangeFromSpec(ctx, spec);
    if (r) return { isRange: true, expr: r };
    return { isRange: false, expr: null };
  }
  return { isRange: false, expr: formatValue(spec) };
}

function _maybeTrim(expr, ctx) {
  return ctx.formatOptions?.trim_text
    ? `TRIM(SUBSTITUTE(${expr}, CHAR(160), " "))`
    : expr;
}

const textFunctionBuilder = {
  left(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    const n = it.num_chars != null ? it.num_chars : 1;
    return _routeTextUnary(t, (x) => `LEFT(${x}, ${n})`, pick);
  },
  right(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    const n = it.num_chars != null ? it.num_chars : 1;
    return _routeTextUnary(t, (x) => `RIGHT(${x}, ${n})`, pick);
  },
  mid(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    const s = it.start_num != null ? it.start_num : 1;
    const n = it.num_chars != null ? it.num_chars : 1;
    return _routeTextUnary(t, (x) => `MID(${x}, ${s}, ${n})`, pick);
  },
  len(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `LEN(${x})`, pick);
  },
  upper(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `UPPER(${x})`, pick);
  },
  lower(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `LOWER(${x})`, pick);
  },
  proper(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `PROPER(${x})`, pick);
  },
  trim(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `TRIM(${x})`, pick);
  },
  value(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `VALUE(${x})`, pick);
  },
  t(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `T(${x})`, pick);
  },
  text(ctx, formatValue) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    const fmt = it.format_text || "0";
    return _routeTextUnary(t, (x) => `TEXT(${x}, ${formatValue(fmt)})`, pick);
  },

  find(ctx, formatValue) {
    const it = ctx.intent || {};
    const ft =
      it.find_text && typeof it.find_text === "object" && it.find_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.find_text) ||
          formatValue("")
        : it.find_text != null
        ? formatValue(it.find_text)
        : formatValue("");
    const start = it.start_num != null ? it.start_num : 1;

    const withinRange = rangeFromSpec(ctx, it.within_text);
    if (withinRange)
      return `=BYROW(${withinRange}, LAMBDA(x, FIND(${ft}, x, ${start})))`;

    const withinScalar =
      it.within_text &&
      typeof it.within_text === "object" &&
      it.within_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.within_text)
        : null;
    if (withinScalar) return `=FIND(${ft}, ${withinScalar}, ${start})`;

    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `FIND(${ft}, ${x}, ${start})`, pick);
  },

  search(ctx, formatValue) {
    const it = ctx.intent || {};
    const ft =
      it.find_text && typeof it.find_text === "object" && it.find_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.find_text) ||
          formatValue("")
        : it.find_text != null
        ? formatValue(it.find_text)
        : formatValue("");
    const start = it.start_num != null ? it.start_num : 1;

    const withinRange = rangeFromSpec(ctx, it.within_text);
    if (withinRange)
      return `=BYROW(${withinRange}, LAMBDA(x, SEARCH(${ft}, x, ${start})))`;

    const withinScalar =
      it.within_text &&
      typeof it.within_text === "object" &&
      it.within_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.within_text)
        : null;
    if (withinScalar) return `=SEARCH(${ft}, ${withinScalar}, ${start})`;

    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `SEARCH(${ft}, ${x}, ${start})`, pick);
  },

  textjoin(ctx, formatValue) {
    const it = ctx.intent || {};
    // delimiter
    let delim = formatValue("");
    if (
      it.delimiter &&
      typeof it.delimiter === "object" &&
      it.delimiter.operation
    ) {
      delim =
        evalSubIntentToScalar(ctx, formatValue, it.delimiter) ||
        formatValue("");
    } else if (it.delimiter) {
      const asRange = rangeFromSpec(ctx, it.delimiter);
      delim = asRange
        ? `INDEX(${asRange}, 1)`
        : formatValue(String(it.delimiter));
    }
    const ignoreEmpty =
      typeof it.ignore_empty === "boolean" ? it.ignore_empty : false;

    const args = [];
    (Array.isArray(it.values) ? it.values : []).forEach((v) => {
      if (v && typeof v === "object" && v.operation) {
        const sub =
          _evalSubExprToRange(ctx, formatValue, v) ||
          evalSubIntentToScalar(ctx, formatValue, v);
        if (sub) args.push(sub);
      } else if (typeof v === "string") {
        const r = rangeFromSpec(ctx, v);
        if (r) args.push(r);
        else args.push(formatValue(v));
      } else if (v && typeof v === "object" && (v.header || v.sheet)) {
        const r = rangeFromSpec(ctx, v);
        if (r) args.push(r);
      } else if (v != null) {
        args.push(formatValue(v));
      }
    });

    if (args.length === 0) {
      const t = _targetCellOrRange(ctx, "text");
      const pick = _selectCellByRowSelector(
        ctx,
        it.target_header || it.header_hint
      );
      if (pick) args.push(pick);
      else if (t && t.isRange) args.push(t.range);
      else if (t && !t.isRange) args.push(t.cell);
    }
    if (args.length === 0) return `=ERROR("TEXTJOIN: 결합할 값이 없습니다.")`;
    return `=TEXTJOIN(${delim}, ${ignoreEmpty ? "TRUE" : "FALSE"}, ${args.join(
      ", "
    )})`;
  },

  textsplit(ctx, formatValue) {
    const it = ctx.intent || {};
    const isSheets =
      String(it.platform || it.engine || "").toLowerCase() === "sheets";

    // 대상 텍스트
    let textExpr = null;
    const withinRange = rangeFromSpec(ctx, it.within_text);
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        textExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      } else if (typeof it.within_text === "string") {
        const r = rangeFromSpec(it.within_text, ctx);
        if (!r) textExpr = formatValue(it.within_text);
      }
      if (!textExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        if (pick) textExpr = pick;
        else if (t && !t.isRange) textExpr = t.cell;
      }
    }

    // 옵션
    const delimScalar = (() => {
      if (
        it.delimiter &&
        typeof it.delimiter === "object" &&
        it.delimiter.operation
      ) {
        return (
          evalSubIntentToScalar(ctx, formatValue, it.delimiter) ||
          formatValue("")
        );
      } else if (it.delimiter) {
        const asRange = rangeFromSpec(ctx, it.delimiter);
        return asRange
          ? `INDEX(${asRange}, 1)`
          : formatValue(String(it.delimiter));
      }
      return formatValue("");
    })();

    const multi =
      it.delimiters && Array.isArray(it.delimiters) ? it.delimiters : null;

    const ignoreEmptyExcel =
      typeof it.ignore_empty === "boolean" ? it.ignore_empty : false;
    const removeEmptySheets =
      typeof it.remove_empty_text === "boolean" ? it.remove_empty_text : true;

    const splitByEach =
      typeof it.split_by_each === "boolean" ? it.split_by_each : false;

    const rowDelim =
      it.row_delimiter != null
        ? typeof it.row_delimiter === "string"
          ? formatValue(it.row_delimiter)
          : formatValue(String(it.row_delimiter))
        : null;
    const matchMode =
      typeof it.match_mode_num === "number" ? it.match_mode_num : 0;
    const padWith = it.pad_with != null ? formatValue(it.pad_with) : null;

    const targetRange =
      withinRange ||
      (() => {
        const t = _targetCellOrRange(ctx, "text");
        return t && t.isRange ? t.range : null;
      })();

    if (isSheets) {
      if (multi && multi.length > 0) {
        const pat = `(${multi
          .map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
          .join("|")})`;
        if (targetRange) {
          return `=BYROW(${targetRange}, LAMBDA(x, TEXTSPLIT(REGEXREPLACE(x, ${formatValue(
            pat
          )}, ${delimScalar}), ${delimScalar}, ${
            splitByEach ? "TRUE" : "FALSE"
          }, ${removeEmptySheets ? "TRUE" : "FALSE"})))`;
        }
        if (!textExpr)
          return `=ERROR("TEXTSPLIT: 분해할 텍스트를 찾지 못했습니다.")`;
        return `=TEXTSPLIT(REGEXREPLACE(${textExpr}, ${formatValue(
          pat
        )}, ${delimScalar}), ${delimScalar}, ${
          splitByEach ? "TRUE" : "FALSE"
        }, ${removeEmptySheets ? "TRUE" : "FALSE"})`;
      }

      if (targetRange) {
        return `=BYROW(${targetRange}, LAMBDA(x, TEXTSPLIT(x, ${delimScalar}, ${
          splitByEach ? "TRUE" : "FALSE"
        }, ${removeEmptySheets ? "TRUE" : "FALSE"})))`;
      }
      if (!textExpr)
        return `=ERROR("TEXTSPLIT: 분해할 텍스트를 찾지 못했습니다.")`;
      return `=TEXTSPLIT(${textExpr}, ${delimScalar}, ${
        splitByEach ? "TRUE" : "FALSE"
      }, ${removeEmptySheets ? "TRUE" : "FALSE"})`;
    }

    // Excel
    const excelColDelim =
      multi && multi.length > 0 ? _excelArrayConst(multi, formatValue) : null;
    const colArg = excelColDelim || delimScalar;
    const rowArg = rowDelim ? `, ${rowDelim}` : "";
    const ignArg = `, ${ignoreEmptyExcel ? "TRUE" : "FALSE"}`;
    const mmArg = `, ${matchMode}`;
    const padArg = padWith ? `, ${padWith}` : "";

    if (targetRange) {
      return `=BYROW(${targetRange}, LAMBDA(x, TEXTSPLIT(x, ${colArg}${rowArg}${ignArg}${mmArg}${padArg})))`;
    }
    if (!textExpr)
      return `=ERROR("TEXTSPLIT: 분해할 텍스트를 찾지 못했습니다.")`;
    return `=TEXTSPLIT(${textExpr}, ${colArg}${rowArg}${ignArg}${mmArg}${padArg})`;
  },

  substitute(ctx, formatValue) {
    const it = ctx.intent || {};
    let textExpr = null;
    const withinRange = rangeFromSpec(
      ctx,
      it.within_text || it.target_header || it.header_hint
    );
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        textExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      }
      if (!textExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        if (pick) textExpr = pick;
        else if (t && !t.isRange) textExpr = t.cell;
      }
    }

    const oldT =
      it.old_text && typeof it.old_text === "object" && it.old_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.old_text) ||
          formatValue("")
        : it.old_text != null
        ? formatValue(it.old_text)
        : formatValue("");
    const newT =
      it.new_text && typeof it.new_text === "object" && it.new_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.new_text) ||
          formatValue("")
        : it.new_text != null
        ? formatValue(it.new_text)
        : formatValue("");
    const inst = it.instance_num != null ? `, ${it.instance_num}` : "";

    if (withinRange) {
      return `=BYROW(${withinRange}, LAMBDA(x, SUBSTITUTE(x, ${oldT}, ${newT}${inst})))`;
    }
    if (!textExpr)
      return `=ERROR("SUBSTITUTE: 대상 텍스트를 찾지 못했습니다.")`;
    return `=SUBSTITUTE(${textExpr}, ${oldT}, ${newT}${inst})`;
  },

  replace(ctx, formatValue) {
    const it = ctx.intent || {};
    let oldExpr = null;
    const withinRange = rangeFromSpec(
      ctx,
      it.within_text || it.target_header || it.header_hint
    );
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        oldExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      }
      if (!oldExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        if (pick) oldExpr = pick;
        else if (t && !t.isRange) oldExpr = t.cell;
      }
    }

    const start = it.start_num != null ? it.start_num : 1;
    const n = it.num_chars != null ? it.num_chars : 0;
    const newT =
      it.new_text && typeof it.new_text === "object" && it.new_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.new_text) ||
          formatValue("")
        : it.new_text != null
        ? formatValue(it.new_text)
        : formatValue("");

    if (withinRange) {
      return `=BYROW(${withinRange}, LAMBDA(x, REPLACE(x, ${start}, ${n}, ${newT})))`;
    }
    if (!oldExpr) return `=ERROR("REPLACE: 대상 텍스트를 찾지 못했습니다.")`;
    return `=REPLACE(${oldExpr}, ${start}, ${n}, ${newT})`;
  },

  concat(ctx, formatValue) {
    const it = ctx.intent || {};
    const args = [];
    (Array.isArray(it.values) ? it.values : []).forEach((v) => {
      if (v && typeof v === "object" && v.operation) {
        const sub =
          _evalSubExprToRange(ctx, formatValue, v) ||
          evalSubIntentToScalar(ctx, formatValue, v);
        if (sub) args.push(sub);
      } else if (typeof v === "string") {
        const r = rangeFromSpec(ctx, v);
        if (r) args.push(r);
        else args.push(formatValue(v));
      } else if (v && typeof v === "object" && (v.header || v.sheet)) {
        const r = rangeFromSpec(ctx, v);
        if (r) args.push(r);
      } else if (v != null) {
        args.push(formatValue(v));
      }
    });
    if (args.length === 0) {
      const t = _targetCellOrRange(ctx, "text");
      const pick = _selectCellByRowSelector(
        ctx,
        it.target_header || it.header_hint
      );
      if (pick) args.push(pick);
      else if (t && t.isRange) args.push(t.range);
      else if (t && !t.isRange) args.push(t.cell);
    }
    if (args.length === 0) return `=ERROR("CONCAT: 연결할 값이 없습니다.")`;
    return `=TEXTJOIN("", FALSE, ${args.join(", ")})`;
  },

  clean(ctx) {
    const it = ctx.intent || {};
    const t = _targetCellOrRange(ctx, "text");
    const pick = _selectCellByRowSelector(
      ctx,
      it.target_header || it.header_hint
    );
    return _routeTextUnary(t, (x) => `CLEAN(${x})`, pick);
  },

  // values: [left, right] / left+right 지정 등
  exact(ctx, formatValue) {
    const it = ctx.intent || {};
    const aSpec =
      (Array.isArray(it.values) && it.values[0]) ||
      it.left ||
      it.within_text ||
      it.target_header ||
      it.header_hint;
    const bSpec =
      (Array.isArray(it.values) && it.values[1]) || it.right || it.find_text;

    const aRange = rangeFromSpec(aSpec, ctx);
    const bRange = rangeFromSpec(bSpec, ctx);

    const toScalar = (spec) => {
      if (spec && typeof spec === "object" && spec.operation) {
        return evalSubIntentToScalar(ctx, formatValue, spec);
      }
      if (typeof spec === "string") {
        const r = rangeFromSpec(ctx, spec);
        if (r) return r;
        return formatValue(spec);
      }
      if (spec && (spec.header || spec.sheet)) {
        const r = rangeFromSpec(ctx, spec);
        return r || null;
      }
      return formatValue(spec ?? "");
    };

    const nocase = String(it.match_mode || "").toLowerCase() === "nocase";
    const wrap = (expr) => (nocase ? `UPPER(${expr})` : expr);

    if (aRange || bRange) {
      const left = aRange || toScalar(aSpec) || `""`;
      const right = bRange || toScalar(bSpec) || `""`;
      return `=BYROW(${left}, LAMBDA(x, ${wrap("x")} = ${wrap(right)}))`;
    }

    let leftExpr = toScalar(aSpec);
    if (!leftExpr) {
      const t = _targetCellOrRange(ctx, "text");
      const pick = _selectCellByRowSelector(
        ctx,
        it.target_header || it.header_hint
      );
      leftExpr = pick || (t && !t.isRange ? t.cell : null) || formatValue("");
    }
    const rightExpr = toScalar(bSpec) || formatValue("");

    return `=${wrap(leftExpr)} = ${wrap(rightExpr)}`;
  },

  /** TEXTBEFORE: IFERROR 제거 → SEARCH 존재검사로 우회 */
  textbefore(ctx, formatValue) {
    const it = ctx.intent || {};
    const useSearch = String(it.match_mode || "").toLowerCase() === "nocase";
    const FINDER = useSearch ? "SEARCH" : "FIND";

    let textExpr = null;
    const withinRange = rangeFromSpec(
      ctx,
      it.within_text || it.target_header || it.header_hint
    );
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        textExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      }
      if (!textExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        textExpr = pick || (t && !t.isRange ? t.cell : null);
      }
    }

    // 구분자
    let delim = formatValue("");
    if (
      it.delimiter &&
      typeof it.delimiter === "object" &&
      it.delimiter.operation
    ) {
      delim =
        evalSubIntentToScalar(ctx, formatValue, it.delimiter) ||
        formatValue("");
    } else if (it.delimiter) {
      const asRange = rangeFromSpec(ctx, it.delimiter);
      delim = asRange
        ? `INDEX(${asRange}, 1)`
        : formatValue(String(it.delimiter));
    }

    const n = it.instance_num != null ? it.instance_num : 1;

    const core = (X) =>
      n <= 1
        ? // 존재 검사 → 있으면 LEFT, 없으면 ""
          `IF(ISNUMBER(SEARCH(${delim}, ${X})), LEFT(${X}, ${FINDER}(${delim}, ${X})-1), "")`
        : // N번째 구분자: SUBSTITUTE 트릭 (에러 없음)
          `LET(_x, ${X}, _t, "|", _y, SUBSTITUTE(_x, ${delim}, _t, ${n}), LEFT(_x, FIND(_t, _y)-1))`;

    if (withinRange) return `=BYROW(${withinRange}, LAMBDA(x, ${core("x")}))`;
    if (!textExpr)
      return `=ERROR("TEXTBEFORE: 대상 텍스트를 찾지 못했습니다.")`;
    return `=${core(textExpr)}`;
  },

  /** TEXTAFTER: IFERROR 제거 → SEARCH 존재검사로 우회 */
  textafter(ctx, formatValue) {
    const it = ctx.intent || {};
    const useSearch = String(it.match_mode || "").toLowerCase() === "nocase";
    const FINDER = useSearch ? "SEARCH" : "FIND";

    let textExpr = null;
    const withinRange = rangeFromSpec(
      ctx,
      it.within_text || it.target_header || it.header_hint
    );
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        textExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      }
      if (!textExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        textExpr = pick || (t && !t.isRange ? t.cell : null);
      }
    }

    let delim = formatValue("");
    if (
      it.delimiter &&
      typeof it.delimiter === "object" &&
      it.delimiter.operation
    ) {
      delim =
        evalSubIntentToScalar(ctx, formatValue, it.delimiter) ||
        formatValue("");
    } else if (it.delimiter) {
      const asRange = rangeFromSpec(ctx, it.delimiter);
      delim = asRange
        ? `INDEX(${asRange}, 1)`
        : formatValue(String(it.delimiter));
    }

    const n = it.instance_num != null ? it.instance_num : 1;

    const core = (X) =>
      n <= 1
        ? `IF(ISNUMBER(SEARCH(${delim}, ${X})), MID(${X}, ${FINDER}(${delim}, ${X})+LEN(${delim}), 10^6), "")`
        : `LET(_x, ${X}, _t, "|", _y, SUBSTITUTE(_x, ${delim}, _t, ${n}), MID(_x, FIND(_t, _y)+1, LEN(_x)))`;

    if (withinRange) return `=BYROW(${withinRange}, LAMBDA(x, ${core("x")}))`;
    if (!textExpr) return `=ERROR("TEXTAFTER: 대상 텍스트를 찾지 못했습니다.")`;
    return `=${core(textExpr)}`;
  },

  /** contains / startswith / endswith (고도화: case_sensitive + trim 옵션) */
  contains(ctx, formatValue) {
    const it = ctx.intent || {};
    const needle =
      _toRangeOrScalarExpr(ctx, formatValue, it.find_text || it.needle).expr ||
      formatValue("");
    const gen = (x) =>
      ctx.formatOptions?.case_sensitive
        ? `ISNUMBER(FIND(${needle}, ${x}))`
        : `ISNUMBER(SEARCH(${needle}, ${x}))`;
    return _asUnaryByRowOrScalar(
      ctx,
      formatValue,
      it.within_text || it.target_header || it.header_hint,
      (x) => gen(_maybeTrim(x, ctx))
    );
  },

  startswith(ctx, formatValue) {
    const it = ctx.intent || {};
    const needle =
      _toRangeOrScalarExpr(ctx, formatValue, it.find_text || it.needle).expr ||
      formatValue("");
    const gen = (x) =>
      ctx.formatOptions?.case_sensitive
        ? `EXACT(LEFT(${x}, LEN(${needle})), ${needle})`
        : `LOWER(LEFT(${x}, LEN(${needle})))=LOWER(${needle})`;
    return _asUnaryByRowOrScalar(
      ctx,
      formatValue,
      it.within_text || it.target_header || it.header_hint,
      (x) => gen(_maybeTrim(x, ctx))
    );
  },

  endswith(ctx, formatValue) {
    const it = ctx.intent || {};
    const needle =
      _toRangeOrScalarExpr(ctx, formatValue, it.find_text || it.needle).expr ||
      formatValue("");
    const gen = (x) =>
      ctx.formatOptions?.case_sensitive
        ? `EXACT(RIGHT(${x}, LEN(${needle})), ${needle})`
        : `LOWER(RIGHT(${x}, LEN(${needle})))=LOWER(${needle})`;
    return _asUnaryByRowOrScalar(
      ctx,
      formatValue,
      it.within_text || it.target_header || it.header_hint,
      (x) => gen(_maybeTrim(x, ctx))
    );
  },

  split_pick(ctx, formatValue) {
    const it = ctx.intent || {};
    const idx = typeof it.index === "number" && it.index >= 1 ? it.index : 1;

    let tExpr = null,
      tRange = rangeFromSpec(
        ctx,
        it.within_text || it.target_header || it.header_hint
      );
    if (!tRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      )
        tExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      if (!tExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        tExpr = pick || (t && !t.isRange ? t.cell : null);
      }
    }
    const delim =
      _toScalarExpr(ctx, formatValue, it.delimiter) || formatValue("");
    if (tRange)
      return `=BYROW(${tRange}, LAMBDA(x, INDEX(TEXTSPLIT(x, ${delim}), ${idx})))`;
    if (!tExpr) return `=ERROR("split_pick: 대상 텍스트를 찾지 못했습니다.")`;
    return `=INDEX(TEXTSPLIT(${tExpr}, ${delim}), ${idx})`;
  },

  coalesce(ctx, formatValue) {
    const it = ctx.intent || {};
    const list = Array.isArray(it.values) ? it.values : [];
    if (list.length === 0) return `=ERROR("coalesce: 후보 값이 없습니다.")`;

    const scalars = list.map((s) => _toScalarExpr(ctx, formatValue, s) || `""`);
    const pick = scalars.reduceRight(
      (acc, cur) => `IF(LEN(TRIM(${cur}))>0, ${cur}, ${acc})`,
      `""`
    );
    return `=${pick}`;
  },

  ifblank(ctx, formatValue) {
    const it = ctx.intent || {};
    const src = _toRangeOrScalarExpr(ctx, formatValue, it.within_text);
    const alt = _toScalarExpr(ctx, formatValue, it.alt) || `""`;

    if (src.isRange) {
      return `=BYROW(${src.expr}, LAMBDA(x, IF(LEN(TRIM(x))=0, ${alt}, x)))`;
    }
    const s = src.expr;
    if (!s) return `=ERROR("ifblank: 대상 텍스트를 찾지 못했습니다.")`;
    return `=IF(LEN(TRIM(${s}))=0, ${alt}, ${s})`;
  },

  slugify(ctx, formatValue) {
    const it = ctx.intent || {};
    const isSheets =
      String(it.platform || it.engine || "").toLowerCase() === "sheets";
    const sep = it.delimiter
      ? _toScalarExpr(ctx, formatValue, it.delimiter)
      : formatValue("-");
    const src = _toRangeOrScalarExpr(
      ctx,
      formatValue,
      it.within_text || it.target_header || it.header_hint
    );
    if (!src.expr) return `=ERROR("slugify: 대상 텍스트를 찾지 못했습니다.")`;

    if (isSheets) {
      const fx = (x) =>
        `LOWER(REGEXREPLACE(REGEXREPLACE(${x}, "[^0-9A-Za-z가-힣]+", ${sep}), ${sep}&"+", ${sep}))`;
      if (src.isRange) return `=BYROW(${src.expr}, LAMBDA(x, ${fx("x")}))`;
      return `=${fx(src.expr)}`;
    }
    const fx = (x) =>
      `LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE(${x}, "_", " "), "-", " ")), "  ", " "), " ", ${sep}), "--", ${sep}))`;
    if (src.isRange) return `=BYROW(${src.expr}, LAMBDA(x, ${fx("x")}))`;
    return `=${fx(src.expr)}`;
  },

  /** 숫자만 추출: Excel에서 IFERROR 없이 CODE() 범위 검사 */
  extract_number(ctx, formatValue) {
    const it = ctx.intent || {};
    const isSheets =
      String(it.platform || it.engine || "").toLowerCase() === "sheets";
    const src = _toRangeOrScalarExpr(
      ctx,
      formatValue,
      it.within_text || it.target_header || it.header_hint
    );
    if (!src.expr)
      return `=ERROR("extract_number: 대상 텍스트를 찾지 못했습니다.")`;

    if (isSheets) {
      const fx = (x) => `TEXTJOIN("", TRUE, REGEXREPLACE(${x}, "[^0-9]", ""))`;
      if (src.isRange) return `=BYROW(${src.expr}, LAMBDA(x, ${fx("x")}))`;
      return `=${fx(src.expr)}`;
    }
    // Excel: CODE 범위(48~57)로 숫자 판정
    const fx = (x) =>
      `TEXTJOIN("", TRUE, IF((CODE(MID(${x}, SEQUENCE(LEN(${x})), 1))>=48)*(CODE(MID(${x}, SEQUENCE(LEN(${x})), 1))<=57), MID(${x}, SEQUENCE(LEN(${x})), 1), ""))`;
    if (src.isRange) return `=BYROW(${src.expr}, LAMBDA(x, ${fx("x")}))`;
    return `=${fx(src.expr)}`;
  },

  pad_left(ctx, formatValue) {
    const it = ctx.intent || {};
    const len = typeof it.length === "number" ? it.length : 0;
    const ch = _toScalarExpr(ctx, formatValue, it.pad_char || " ");
    const src = _toRangeOrScalarExpr(
      ctx,
      formatValue,
      it.within_text || it.target_header || it.header_hint
    );
    if (!src.expr) return `=ERROR("pad_left: 대상 텍스트를 찾지 못했습니다.")`;
    const fx = (x) => `IF(LEN(${x})>=${len}, ${x}, REPT(${ch}, ${len}) & ${x})`;
    if (src.isRange) return `=BYROW(${src.expr}, LAMBDA(x, ${fx("x")}))`;
    return `=${fx(src.expr)}`;
  },

  pad_right(ctx, formatValue) {
    const it = ctx.intent || {};
    const len = typeof it.length === "number" ? it.length : 0;
    const ch = _toScalarExpr(ctx, formatValue, it.pad_char || " ");
    const src = _toRangeOrScalarExpr(
      ctx,
      formatValue,
      it.within_text || it.target_header || it.header_hint
    );
    if (!src.expr) return `=ERROR("pad_right: 대상 텍스트를 찾지 못했습니다.")`;
    const fx = (x) => `IF(LEN(${x})>=${len}, ${x}, ${x} & REPT(${ch}, ${len}))`;
    if (src.isRange) return `=BYROW(${src.expr}, LAMBDA(x, ${fx("x")}))`;
    return `=${fx(src.expr)}`;
  },

  if_contains(ctx, formatValue) {
    const it = ctx.intent || {};
    const condExpr = textFunctionBuilder.contains(ctx, formatValue).slice(1); // "=" 제거
    const vT = _toScalarExpr(ctx, formatValue, it.value_if_true) || `""`;
    const vF = _toScalarExpr(ctx, formatValue, it.value_if_false) || `""`;
    if (/^BYROW\(/.test(condExpr)) {
      return `=${condExpr.replace(
        /^BYROW\((.+)\)$/,
        `BYROW($1, LAMBDA(x, IF(x, ${vT}, ${vF})))`
      )}`;
    }
    return `=IF(${condExpr}, ${vT}, ${vF})`;
  },

  if_startswith(ctx, formatValue) {
    const it = ctx.intent || {};
    const condExpr = textFunctionBuilder.startswith(ctx, formatValue).slice(1);
    const vT = _toScalarExpr(ctx, formatValue, it.value_if_true) || `""`;
    const vF = _toScalarExpr(ctx, formatValue, it.value_if_false) || `""`;
    if (/^BYROW\(/.test(condExpr)) {
      return `=${condExpr.replace(
        /^BYROW\((.+)\)$/,
        `BYROW($1, LAMBDA(x, IF(x, ${vT}, ${vF})))`
      )}`;
    }
    return `=IF(${condExpr}, ${vT}, ${vF})`;
  },

  if_endswith(ctx, formatValue) {
    const it = ctx.intent || {};
    const condExpr = textFunctionBuilder.endswith(ctx, formatValue).slice(1);
    const vT = _toScalarExpr(ctx, formatValue, it.value_if_true) || `""`;
    const vF = _toScalarExpr(ctx, formatValue, it.value_if_false) || `""`;
    if (/^BYROW\(/.test(condExpr)) {
      return `=${condExpr.replace(
        /^BYROW\((.+)\)$/,
        `BYROW($1, LAMBDA(x, IF(x, ${vT}, ${vF})))`
      )}`;
    }
    return `=IF(${condExpr}, ${vT}, ${vF})`;
  },

  regexmatch(ctx, formatValue) {
    const it = ctx.intent || {};
    const isSheets =
      String(it.platform || it.engine || "").toLowerCase() === "sheets";

    let textExpr = null;
    const withinRange = rangeFromSpec(
      ctx,
      it.within_text || it.target_header || it.header_hint
    );
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        textExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      }
      if (!textExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        textExpr = pick || (t && !t.isRange ? t.cell : null);
      }
    }
    const pattern =
      it.find_text && typeof it.find_text === "object" && it.find_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.find_text) ||
          formatValue("")
        : it.find_text != null
        ? formatValue(it.find_text)
        : formatValue("");

    if (isSheets) {
      if (withinRange)
        return `=BYROW(${withinRange}, LAMBDA(x, REGEXMATCH(x, ${pattern})))`;
      if (!textExpr)
        return `=ERROR("REGEXMATCH: 대상 텍스트를 찾지 못했습니다.")`;
      return `=REGEXMATCH(${textExpr}, ${pattern})`;
    }

    if (withinRange) {
      const fb = _regexLiteFallbackExpr("match", "x", pattern, formatValue);
      if (fb)
        return `=BYROW(${withinRange}, LAMBDA(x, ${fb.replaceAll("x", "x")}))`;
      return `=ERROR("REGEXMATCH는 Sheets 전용입니다(간단 패턴만 우회 가능).")`;
    }
    if (!textExpr)
      return `=ERROR("REGEXMATCH: 대상 텍스트를 찾지 못했습니다.")`;

    const fb = _regexLiteFallbackExpr("match", textExpr, pattern, formatValue);
    return (
      fb || `=ERROR("REGEXMATCH는 Sheets 전용입니다(간단 패턴만 우회 가능).")`
    );
  },

  regexextract(ctx, formatValue) {
    const it = ctx.intent || {};
    const isSheets =
      String(it.platform || it.engine || "").toLowerCase() === "sheets";

    let textExpr = null;
    const withinRange = rangeFromSpec(
      ctx,
      it.within_text || it.target_header || it.header_hint
    );
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        textExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      }
      if (!textExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        textExpr = pick || (t && !t.isRange ? t.cell : null);
      }
    }
    const pattern =
      it.find_text && typeof it.find_text === "object" && it.find_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.find_text) ||
          formatValue("")
        : it.find_text != null
        ? formatValue(it.find_text)
        : formatValue("");

    if (isSheets) {
      if (withinRange)
        return `=BYROW(${withinRange}, LAMBDA(x, REGEXEXTRACT(x, ${pattern})))`;
      if (!textExpr)
        return `=ERROR("REGEXEXTRACT: 대상 텍스트를 찾지 못했습니다.")`;
      return `=REGEXEXTRACT(${textExpr}, ${pattern})`;
    }

    if (withinRange) {
      const fb = _regexLiteFallbackExpr("extract", "x", pattern, formatValue);
      if (fb)
        return `=BYROW(${withinRange}, LAMBDA(x, ${fb.replaceAll("x", "x")}))`;
      return `=ERROR("REGEXEXTRACT는 Sheets 전용입니다(간단 패턴만 우회 가능).")`;
    }
    if (!textExpr)
      return `=ERROR("REGEXEXTRACT: 대상 텍스트를 찾지 못했습니다.")`;

    const fb = _regexLiteFallbackExpr(
      "extract",
      textExpr,
      pattern,
      formatValue
    );
    return (
      fb || `=ERROR("REGEXEXTRACT는 Sheets 전용입니다(간단 패턴만 우회 가능).")`
    );
  },

  regexreplace(ctx, formatValue) {
    const it = ctx.intent || {};
    const isSheets =
      String(it.platform || it.engine || "").toLowerCase() === "sheets";

    let textExpr = null;
    const withinRange = rangeFromSpec(
      ctx,
      it.within_text || it.target_header || it.header_hint
    );
    if (!withinRange) {
      if (
        it.within_text &&
        typeof it.within_text === "object" &&
        it.within_text.operation
      ) {
        textExpr = evalSubIntentToScalar(ctx, formatValue, it.within_text);
      }
      if (!textExpr) {
        const t = _targetCellOrRange(ctx, "text");
        const pick = _selectCellByRowSelector(
          ctx,
          it.target_header || it.header_hint
        );
        textExpr = pick || (t && !t.isRange ? t.cell : null);
      }
    }
    const pattern =
      it.old_text && typeof it.old_text === "object" && it.old_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.old_text) ||
          formatValue("")
        : it.old_text != null
        ? formatValue(it.old_text)
        : formatValue("");
    const repl =
      it.new_text && typeof it.new_text === "object" && it.new_text.operation
        ? evalSubIntentToScalar(ctx, formatValue, it.new_text) ||
          formatValue("")
        : it.new_text != null
        ? formatValue(it.new_text)
        : formatValue("");

    if (isSheets) {
      if (withinRange)
        return `=BYROW(${withinRange}, LAMBDA(x, REGEXREPLACE(x, ${pattern}, ${repl})))`;
      if (!textExpr)
        return `=ERROR("REGEXREPLACE: 대상 텍스트를 찾지 못했습니다.")`;
      return `=REGEXREPLACE(${textExpr}, ${pattern}, ${repl})`;
    }

    if (withinRange) {
      const fb = _regexLiteFallbackExpr("replace", "x", pattern, formatValue);
      if (fb)
        return `=BYROW(${withinRange}, LAMBDA(x, ${fb.replaceAll("x", "x")}))`;
      return `=ERROR("REGEXREPLACE는 Sheets 전용입니다(간단 패턴만 우회 가능).")`;
    }
    if (!textExpr)
      return `=ERROR("REGEXREPLACE: 대상 텍스트를 찾지 못했습니다.")`;

    const fb = _regexLiteFallbackExpr(
      "replace",
      textExpr,
      pattern,
      formatValue
    );
    return (
      fb || `=ERROR("REGEXREPLACE는 Sheets 전용입니다(간단 패턴만 우회 가능).")`
    );
  },
};

textFunctionBuilder.concatenate = textFunctionBuilder.concat;
textFunctionBuilder.join = textFunctionBuilder.textjoin;

module.exports = textFunctionBuilder;
