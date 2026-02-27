const {
  refFromHeaderSpec,
  rangeFromSpec,
  evalSubIntentToScalar,
} = require("../utils/builderHelpers");

function _isSheets(ctx) {
  return String(ctx.engine || ctx.platform || "")
    .toLowerCase()
    .includes("sheet");
}

function _pickValueRange(ctx, it) {
  // metric/value 대상 열(최대/최소 기준)
  const h =
    it.value_header ||
    it.metric_header ||
    it.header_hint ||
    it.value_hint ||
    null;
  const r =
    (h ? refFromHeaderSpec(ctx, h) : null) ||
    (h
      ? refFromHeaderSpec(ctx, { header: h, sheet: ctx.bestReturn?.sheetName })
      : null) ||
    null;
  if (r) return r.range;
  if (ctx.bestReturn)
    return `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}:${ctx.bestReturn.columnLetter}${ctx.bestReturn.lastDataRow}`;
  return null;
}

function _pickReturnRange(ctx, it) {
  // 반환할 열(이름/항목 등)
  const h =
    it.return_header ||
    it.return_hint ||
    it.lookup_hint || // 사용자가 "누구/어떤 항목"을 물을 때 흔히 여기로 들어옴
    null;
  const r =
    (h ? refFromHeaderSpec(ctx, h) : null) ||
    (h
      ? refFromHeaderSpec(ctx, { header: h, sheet: ctx.bestReturn?.sheetName })
      : null) ||
    null;
  if (r) return r.range;
  if (ctx.bestLookup)
    return `'${ctx.bestLookup.sheetName}'!${ctx.bestLookup.columnLetter}${ctx.bestLookup.startRow}:${ctx.bestLookup.columnLetter}${ctx.bestLookup.lastDataRow}`;
  return null;
}

function _targetRangeFromBest(bestReturn) {
  return `'${bestReturn.sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}:${bestReturn.columnLetter}${bestReturn.lastDataRow}`;
}

function _ensureConditionPairs(ctx, buildConditionPairs) {
  if (!buildConditionPairs) return [];
  const pairs = buildConditionPairs(ctx) || [];
  return pairs.filter(Boolean);
}

function _buildFilterCall(targetRange, conditionPairs) {
  // ✅ 조건 없이 FILTER를 쓰면 '그럴듯한 오답'이 나올 수 있어 명시 실패
  if (!conditionPairs.length) return `=ERROR("FILTER: 조건이 비어 있습니다.")`;
  const clauses = [];
  for (let i = 0; i < conditionPairs.length; i += 2) {
    const rng = conditionPairs[i];
    const crit = conditionPairs[i + 1];
    clauses.push(`${rng}, ${crit}`);
  }
  return `FILTER(${targetRange}, ${clauses.join(", ")})`;
}

function _evalSubExprToRange(ctx, formatValue, node) {
  if (!node || typeof node !== "object" || !node.operation) return null;
  const op = String(node.operation).toLowerCase();

  if (op === "filter") {
    const src = rangeFromSpec(ctx, node.source_range);
    if (!src) return `ERROR("FILTER: source_range 없음")`;
    const crits = node.criteria || [];
    const clauses = [];
    for (const c of crits) {
      if (Array.isArray(c)) {
        const [rng, crit] = c;
        const rr = rangeFromSpec(ctx, rng);
        if (!rr) return `ERROR("FILTER: criteria range 없음")`;
        if (typeof crit === "object" && crit != null) {
          const op2 = _op(crit.op || "=");
          const v2 = _scalarFrom(crit.value, ctx, formatValue);
          clauses.push(`${rr}, "${op2.replace(/"/g, '""')}"&${v2}`);
        } else {
          clauses.push(
            `${rr}, ${
              typeof crit === "string"
                ? formatValue(crit)
                : formatValue(String(crit))
            }`,
          );
        }
      }
    }
    return `FILTER(${src}, ${clauses.join(", ")})`;
  }

  if (op === "xlookup") {
    const lv = _scalarFrom(node.lookup_value, ctx, formatValue);
    const la = rangeFromSpec(ctx, node.lookup_array);
    const ra = rangeFromSpec(ctx, node.return_array);
    if (!la || !ra) return `ERROR("XLOOKUP: lookup/return 범위 없음")`;
    const notFound =
      node.if_not_found != null ? `, ${formatValue(node.if_not_found)}` : "";
    return `XLOOKUP(${lv}, ${la}, ${ra}${notFound})`;
  }

  return `ERROR("range-subexpr '${op}' 미지원")`;
}

function _collectPairs(ctx, it, buildConditionPairs, formatValue) {
  let pairs = _ensureConditionPairs(ctx, buildConditionPairs);
  const winPairs = _injectWindowToConditionPairs(it?.window, ctx, formatValue);
  if (winPairs.length) pairs = pairs.concat(winPairs);
  return pairs;
}

function _injectWindowToConditionPairs(windowObj, ctx, _formatValue) {
  if (!windowObj || windowObj.type !== "days") return [];
  const size = Number(windowObj.size || 0);
  if (!size) return [];
  const hdr = windowObj.date_header || "날짜";
  const rr = refFromHeaderSpec(ctx, { header: hdr, sheet: windowObj.sheet });
  if (!rr) return [];
  return [rr.range, `">="&TODAY()-${size}`];
}

function _scalarFrom(spec, ctx, formatValue) {
  if (spec == null) return "0";
  if (typeof spec === "object" && spec.header) {
    const r = refFromHeaderSpec(ctx, spec);
    return r ? r.cell : "0";
  }
  if (typeof spec === "number") return String(spec);
  if (typeof spec === "string") {
    if (/^\s*(TODAY\(\)|DATE\(|WORKDAY\(|EOMONTH\(|NOW\(\))/.test(spec))
      return spec.trim();
    return formatValue(spec);
  }
  return formatValue(spec);
}

const mathStatsFunctionBuilder = {
  /* SUM / AVERAGE / COUNT (+ group_by) */
  sum: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const sumRange = _targetRangeFromBest(bestReturn);
    const conditionPairs = _collectPairs(
      ctx,
      intent,
      buildConditionPairs,
      formatValue,
    );
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      // ✅ Google Sheets에서는 group_by를 QUERY로 강제 (안정성 ↑)
      if (String(ctx.engine).toLowerCase().includes("sheet")) {
        const key = keyRef.range;
        const val = sumRange;
        return `=QUERY({${key},${val}},
"select Col1, sum(Col2) where Col1 is not null group by Col1 label sum(Col2) ''",
0)`;
      }
      const inner = (kSym) => {
        const pairsPlus = conditionPairs.length
          ? `${conditionPairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        return `SUMIFS(${sumRange}, ${pairsPlus})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    if (!conditionPairs.length) {
      return this._buildSimpleAggregate("sum", ctx);
    }
    return `=SUMIFS(${sumRange}, ${conditionPairs.join(", ")})`;
  },

  average: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    if (!intent.conditions || intent.conditions.length === 0) {
      return this._buildSimpleAggregate("average", ctx);
    }
    const avgRange = _targetRangeFromBest(bestReturn);
    const conditionPairs = _collectPairs(
      ctx,
      intent,
      buildConditionPairs,
      formatValue,
    );
    if (conditionPairs.length === 0) {
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    }
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      // ✅ Google Sheets에서는 group_by를 QUERY로 강제 (안정성 ↑)
      if (String(ctx.engine).toLowerCase().includes("sheet")) {
        const key = keyRef.range;
        const val = avgRange;
        return `=QUERY({${key},${val}},
"select Col1, sum(Col2) where Col1 is not null group by Col1 label sum(Col2) ''",
0)`;
      }
      const inner = (kSym) => {
        const pairsPlus = conditionPairs.length
          ? `${conditionPairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        return `AVERAGEIFS(${avgRange}, ${pairsPlus})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=AVERAGEIFS(${avgRange}, ${conditionPairs.join(", ")})`;
  },

  count: function (ctx, formatValue, buildConditionPairs) {
    const { intent } = ctx;
    if (!intent.conditions || intent.conditions.length === 0) {
      return this._buildSimpleAggregate("count", ctx);
    }
    const conditionPairs = _collectPairs(
      ctx,
      intent,
      buildConditionPairs,
      formatValue,
    );
    if (conditionPairs.length === 0) {
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    }
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      // ✅ Google Sheets에서는 group_by를 QUERY로 강제 (안정성 ↑)
      if (String(ctx.engine).toLowerCase().includes("sheet")) {
        const key = keyRef.range;
        const val = _targetRangeFromBest(ctx.bestReturn);
        return `=QUERY({${key},${val}},
"select Col1, sum(Col2) where Col1 is not null group by Col1 label sum(Col2) ''",
0)`;
      }
      const inner = (kSym) => {
        const pairsPlus = conditionPairs.length
          ? `${conditionPairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        return `COUNTIFS(${pairsPlus})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=COUNTIFS(${conditionPairs.join(", ")})`;
  },

  // ✅ 최대/최소 값이 있는 레코드 반환(예: 최고 연봉자의 이름)
  // intent:
  //  - value_header/header_hint: 기준 값 열(예: 연봉)
  //  - return_header/return_hint: 반환 열(예: 이름)
  argmax_record: function (ctx, formatValue, buildConditionPairs) {
    const it = ctx.intent || {};
    const v = refFromHeaderSpec(
      ctx,
      it.value_header || it.header_hint || "",
    )?.range;
    const r = refFromHeaderSpec(
      ctx,
      it.return_header || it.return_hint || "",
    )?.range;
    if (!v) return `=ERROR("argmax_record: value 열을 찾을 수 없습니다.")`;
    if (!r) return `=ERROR("argmax_record: return 열을 찾을 수 없습니다.")`;
    const pairs = (buildConditionPairs ? buildConditionPairs(ctx) : []) || [];
    const vf = pairs.length
      ? `FILTER(${v}, ${pairs
          .map((_, i) => (i % 2 === 0 ? `${pairs[i]}, ${pairs[i + 1]}` : null))
          .filter(Boolean)
          .join(", ")})`
      : v;
    const rf = pairs.length
      ? `FILTER(${r}, ${pairs
          .map((_, i) => (i % 2 === 0 ? `${pairs[i]}, ${pairs[i + 1]}` : null))
          .filter(Boolean)
          .join(", ")})`
      : r;
    return `=LET(_v,${vf},_r,${rf},_x,MAX(_v),XLOOKUP(_x,_v,_r))`;
  },
  argmin_record: function (ctx, formatValue, buildConditionPairs) {
    const it = ctx.intent || {};
    const v = refFromHeaderSpec(
      ctx,
      it.value_header || it.header_hint || "",
    )?.range;
    const r = refFromHeaderSpec(
      ctx,
      it.return_header || it.return_hint || "",
    )?.range;
    if (!v) return `=ERROR("argmin_record: value 열을 찾을 수 없습니다.")`;
    if (!r) return `=ERROR("argmin_record: return 열을 찾을 수 없습니다.")`;
    const pairs = (buildConditionPairs ? buildConditionPairs(ctx) : []) || [];
    const vf = pairs.length
      ? `FILTER(${v}, ${pairs
          .map((_, i) => (i % 2 === 0 ? `${pairs[i]}, ${pairs[i + 1]}` : null))
          .filter(Boolean)
          .join(", ")})`
      : v;
    const rf = pairs.length
      ? `FILTER(${r}, ${pairs
          .map((_, i) => (i % 2 === 0 ? `${pairs[i]}, ${pairs[i + 1]}` : null))
          .filter(Boolean)
          .join(", ")})`
      : r;
    return `=LET(_v,${vf},_r,${rf},_x,MIN(_v),XLOOKUP(_x,_v,_r))`;
  },

  counta: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions || intent.conditions.length === 0) {
      return `=COUNTA(${tgt})`;
    }
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (pairs.length === 0)
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const inner = (kSym) => {
        const base = pairs.length
          ? `${tgt}, "<>", ${pairs.join(", ")}`
          : `${tgt}, "<>"`;
        const pairsPlus = `${base}, ${keyRef.range}, ${kSym}`;
        return `COUNTIFS(${pairsPlus})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=COUNTIFS(${tgt}, "<>", ${pairs.join(", ")})`;
  },

  countblank: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions || intent.conditions.length === 0) {
      return `=COUNTBLANK(${tgt})`;
    }
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (pairs.length === 0)
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    return `=COUNTIFS(${tgt}, "", ${pairs.join(", ")})`;
  },

  min: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions || intent.conditions.length === 0) {
      return `=MIN(${tgt})`;
    }
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (pairs.length === 0)
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;

      // ✅ Google Sheets: group_by 집계는 QUERY 강제
      if (_isSheets(ctx)) {
        const key = keyRef.range;
        const val = sumRange;
        return `=QUERY({${key},${val}},"select Col1,sum(Col2) where Col1 is not null group by Col1 label sum(Col2) ''",0)`;
      }

      const inner = (kSym) => {
        const pairsPlus = pairs.length
          ? `${pairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        return `MINIFS(${tgt}, ${pairsPlus})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=MINIFS(${tgt}, ${pairs.join(", ")})`;
  },

  max: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions || intent.conditions.length === 0) {
      return `=MAX(${tgt})`;
    }
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (pairs.length === 0)
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      // ✅ Google Sheets: group_by 평균은 QUERY 강제
      if (_isSheets(ctx)) {
        const key = keyRef.range;
        const val = avgRange;
        return `=QUERY({${key},${val}},"select Col1,avg(Col2) where Col1 is not null group by Col1 label avg(Col2) ''",0)`;
      }
      const inner = (kSym) => {
        const pairsPlus = pairs.length
          ? `${pairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        return `MAXIFS(${tgt}, ${pairsPlus})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=MAXIFS(${tgt}, ${pairs.join(", ")})`;
  },

  median: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions || intent.conditions.length === 0) {
      return `=MEDIAN(${tgt})`;
    }
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (pairs.length === 0)
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    const filtered = _buildFilterCall(tgt, pairs);
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          tgt,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `MEDIAN(${filteredK})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=MEDIAN(${filtered})`;
  },

  stdev_s: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions || intent.conditions.length === 0) {
      return `=STDEV.S(${tgt})`;
    }
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (pairs.length === 0)
      return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    const filtered = _buildFilterCall(tgt, pairs);
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          tgt,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `STDEV.S(${filteredK})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=STDEV.S(${filtered})`;
  },

  stdev_p: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    const filtered = _buildFilterCall(tgt, pairs);

    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          tgt,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `STDEV.P(${filteredK})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=STDEV.P(${filtered})`;
  },

  var_s: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions?.length) return `=VAR.S(${tgt})`;
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (!pairs.length) return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          tgt,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `VAR.S(${filteredK})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=VAR.S(${_buildFilterCall(tgt, pairs)})`;
  },

  var_p: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions?.length) return `=VAR.P(${tgt})`;
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (!pairs.length) return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          tgt,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `VAR.P(${filteredK})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=VAR.P(${_buildFilterCall(tgt, pairs)})`;
  },

  percentile_inc: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const k = intent.k ?? 0.9;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions?.length) return `=PERCENTILE.INC(${tgt}, ${k})`;
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (!pairs.length) return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const K = ctx.intent.k ?? 0.9;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          tgt,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `PERCENTILE.INC(${filteredK}, ${K})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=PERCENTILE.INC(${_buildFilterCall(tgt, pairs)}, ${k})`;
  },

  quartile_inc: function (ctx, formatValue, buildConditionPairs) {
    const { intent, bestReturn } = ctx;
    const q = intent.quartile ?? 3;
    const tgt = _targetRangeFromBest(bestReturn);
    if (!intent.conditions?.length) return `=QUARTILE.INC(${tgt}, ${q})`;
    const pairs = _collectPairs(ctx, intent, buildConditionPairs, formatValue);
    if (!pairs.length) return `=ERROR("조건에 맞는 열을 찾을 수 없습니다.")`;
    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const Q = ctx.intent.quartile ?? 3;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          tgt,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `QUARTILE.INC(${filteredK}, ${Q})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=QUARTILE.INC(${_buildFilterCall(tgt, pairs)}, ${q})`;
  },

  _buildSimpleAggregate: (funcName, ctx) => {
    const { bestReturn } = ctx;
    const targetRange = `${bestReturn.columnLetter}${bestReturn.startRow}:${bestReturn.columnLetter}${bestReturn.lastDataRow}`;
    return `=${funcName.toUpperCase()}('${
      bestReturn.sheetName
    }'!${targetRange})`;
  },

  rank: function (ctx, formatValue) {
    const { intent, bestReturn } = ctx;
    const col = _targetRangeFromBest(bestReturn);
    const baseCell = `'${bestReturn.sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}`;
    if (intent.row_selector?.hint && intent.row_selector?.value != null) {
      const keyHeader = intent.row_selector.hint;
      const keyVal = formatValue(intent.row_selector.value);
      return `=RANK(XLOOKUP(${keyVal}, ${col}, ${col}), ${col})`;
    }
    return `=RANK(${baseCell}, ${col})`;
  },

  round: function (ctx, _formatValue) {
    const { intent, bestReturn } = ctx;
    const places = intent?.places != null ? intent.places : 0;
    const baseCell = `'${bestReturn.sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}`;
    return `=ROUND(${baseCell}, ${places})`;
  },

  abs: function (ctx) {
    const { bestReturn } = ctx;
    const baseCell = `'${bestReturn.sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}`;
    return `=ABS(${baseCell})`;
  },

  mod: function (ctx) {
    const { intent, bestReturn } = ctx;
    const divisor = intent?.divisor != null ? intent.divisor : 10;
    const baseCell = `'${bestReturn.sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}`;
    return `=MOD(${baseCell}, ${divisor})`;
  },

  sqrt: function (ctx) {
    const { bestReturn } = ctx;
    const baseCell = `'${bestReturn.sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}`;
    return `=SQRT(${baseCell})`;
  },

  randbetween: function (ctx) {
    const min = ctx?.intent?.min ?? 1;
    const max = ctx?.intent?.max ?? 100;
    return `=RANDBETWEEN(${min}, ${max})`;
  },

  sequence: function (ctx) {
    const rows = ctx?.intent?.rows ?? 10;
    const cols = ctx?.intent?.cols ?? 1;
    const start = ctx?.intent?.start ?? 1;
    const step = ctx?.intent?.step ?? 1;
    return `=SEQUENCE(${rows}, ${cols}, ${start}, ${step})`;
  },

  sumproduct: function (ctx) {
    const it = ctx.intent || {};
    const vr = refFromHeaderSpec(ctx, it.value_range);
    const wr = refFromHeaderSpec(ctx, it.weight_range);
    if (!vr || !wr)
      return `=ERROR("SUMPRODUCT: value_range / weight_range 를 찾을 수 없습니다.")`;

    const join = it.join_by || it.join_on;
    if (join) {
      const leftKey =
        refFromHeaderSpec(
          ctx,
          join?.left_key || { header: join, sheet: vr.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: vr.sheetName });
      const rightKey =
        refFromHeaderSpec(
          ctx,
          join?.right_key || { header: join, sheet: wr.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: wr.sheetName });
      if (leftKey && rightKey) {
        const wAligned = `XLOOKUP(${leftKey.range}, ${rightKey.range}, ${wr.range})`;
        return `=SUMPRODUCT(${vr.range}, ${wAligned})`;
      }
    }
    return `=SUMPRODUCT(${vr.range}, ${wr.range})`;
  },

  correl: function (ctx) {
    const it = ctx.intent || {};
    const xr = refFromHeaderSpec(ctx, it.x_range);
    const yr = refFromHeaderSpec(ctx, it.y_range);
    if (!xr || !yr)
      return `=ERROR("CORREL: x_range / y_range 를 찾을 수 없습니다.")`;

    const join = it.join_by || it.join_on;
    if (join) {
      const xKey =
        refFromHeaderSpec(
          ctx,
          join?.left_key || { header: join, sheet: xr.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: xr.sheetName });
      const yKey =
        refFromHeaderSpec(
          ctx,
          join?.right_key || { header: join, sheet: yr.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: yr.sheetName });
      if (xKey && yKey) {
        const yAligned = `XLOOKUP(${xKey.range}, ${yKey.range}, ${yr.range})`;
        return `=CORREL(${xr.range}, ${yAligned})`;
      }
    }
    return `=CORREL(${xr.range}, ${yr.range})`;
  },

  weighted_average: function (ctx, formatValue, buildConditionPairs) {
    const it = ctx.intent || {};
    let vExpr = null,
      wExpr = null;
    let vRef = null,
      wRef = null;

    if (it.value_range) {
      if (typeof it.value_range === "object" && it.value_range.operation) {
        const e = _evalSubExprToRange(ctx, formatValue, it.value_range);
        if (!e || /^ERROR/.test(e))
          return `=ERROR("weighted_average: value_range 해석 실패")`;
        vExpr = e;
      } else {
        vRef = refFromHeaderSpec(ctx, it.value_range);
        if (!vRef)
          return `=ERROR("weighted_average: value_range를 찾을 수 없습니다.")`;
        vExpr = vRef.range;
      }
    } else if (ctx.bestReturn) {
      vRef = ctx.bestReturn;
      vExpr = _targetRangeFromBest(vRef);
    } else {
      return `=ERROR("weighted_average: 대상 범위를 찾을 수 없습니다.")`;
    }

    if (it.weight_range) {
      if (typeof it.weight_range === "object" && it.weight_range.operation) {
        const e = _evalSubExprToRange(ctx, formatValue, it.weight_range);
        if (!e || /^ERROR/.test(e))
          return `=ERROR("weighted_average: weight_range 해석 실패")`;
        wExpr = e;
      } else {
        wRef = refFromHeaderSpec(ctx, it.weight_range);
        if (!wRef)
          return `=ERROR("weighted_average: weight_range를 찾을 수 없습니다.")`;
        wExpr = wRef.range;
      }
    } else {
      return `=ERROR("weighted_average: 가중치 범위가 필요합니다.")`;
    }

    const join = it.join_by || it.join_on;
    if (join) {
      const vSheet = (vRef && vRef.sheetName) || null;
      const wSheet = (wRef && wRef.sheetName) || null;

      const leftKey =
        refFromHeaderSpec(
          ctx,
          join?.left_key || { header: join, sheet: vSheet },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: vSheet });
      const rightKey =
        refFromHeaderSpec(
          ctx,
          join?.right_key || { header: join, sheet: wSheet },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: wSheet });

      if (leftKey && rightKey && wExpr) {
        wExpr = `XLOOKUP(${leftKey.range}, ${rightKey.range}, ${wExpr})`;
      }
    }

    const conditionPairs = _collectPairs(
      ctx,
      it,
      buildConditionPairs,
      formatValue,
    );
    if (conditionPairs.length) {
      const clauses = [];
      for (let i = 0; i < conditionPairs.length; i += 2) {
        const rng = conditionPairs[i];
        const crit = conditionPairs[i + 1];
        clauses.push(`${rng}, ${crit}`);
      }
      vExpr = `FILTER(${vExpr}, ${clauses.join(", ")})`;
      wExpr = `FILTER(${wExpr}, ${clauses.join(", ")})`;
    }

    const num = `SUMPRODUCT(${vExpr}, ${wExpr})`;
    const den = `SUM(${wExpr})`;

    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;

      const baseClauses = [];
      for (let i = 0; i < (conditionPairs || []).length; i += 2) {
        baseClauses.push(`${conditionPairs[i]}, ${conditionPairs[i + 1]}`);
      }

      const inner = (kSym) => {
        const allClauses = baseClauses.length
          ? `${baseClauses.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const vK = `FILTER(${vExpr}, ${allClauses})`;
        const wK = `FILTER(${wExpr}, ${allClauses})`;
        return `IF(SUM(${wK})=0, "", SUMPRODUCT(${vK}, ${wK}) / SUM(${wK}))`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }

    return `=IF(${den} = 0, "", ${num} / ${den})`;
  },

  percentrank_inc: function (ctx, formatValue, buildConditionPairs) {
    const it = ctx.intent || {};

    let srcRange;
    if (it.source_range) {
      const r = rangeFromSpec(ctx, it.source_range);
      if (!r) return `=ERROR("percentrank_inc: source_range 없음")`;
      srcRange = r;
    } else if (ctx.bestReturn) {
      srcRange = _targetRangeFromBest(ctx.bestReturn);
    } else {
      return `=ERROR("percentrank_inc: 대상 범위를 찾을 수 없습니다.")`;
    }

    const pairs = _collectPairs(ctx, it, buildConditionPairs, formatValue);
    const filtered = pairs.length
      ? _buildFilterCall(srcRange, pairs)
      : srcRange;

    let xExpr = null;
    if (it.x && typeof it.x === "object" && it.x.operation) {
      xExpr = evalSubIntentToScalar(ctx, formatValue, it.x);
    } else if (it.row_selector?.hint && it.row_selector?.value != null) {
      const keyRef =
        refFromHeaderSpec(ctx, {
          header: it.row_selector.hint,
          sheet: ctx.bestReturn?.sheetName,
        }) || refFromHeaderSpec(ctx, it.row_selector.hint);
      if (!keyRef)
        return `=ERROR("percentrank_inc: row_selector 키열을 찾을 수 없습니다.")`;
      const keyVal = formatValue(it.row_selector.value);
      xExpr = `XLOOKUP(${keyVal}, ${keyRef.range}, ${srcRange})`;
    } else if (it.x != null) {
      xExpr = _scalarFrom(it.x, ctx, formatValue);
    } else {
      xExpr = `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}`;
    }

    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const filteredK = _buildFilterCall(
          srcRange,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `PERCENTRANK.INC(${filteredK}, ${xExpr})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }
    return `=PERCENTRANK.INC(${filtered}, ${xExpr})`;
  },

  // ✅ "최대/최소값을 가진 레코드(이름/항목)" 반환
  // - 예: "가장 큰 매출을 기록한 고객명"
  // intent 예시:
  // { operation:"argmax_record", header_hint:"매출", return_hint:"고객명", conditions:[...] }
  argmax_record: function (ctx, formatValue, buildConditionPairs) {
    return _buildArgExtRecord(ctx, formatValue, buildConditionPairs, "MAX");
  },
  argmin_record: function (ctx, formatValue, buildConditionPairs) {
    return _buildArgExtRecord(ctx, formatValue, buildConditionPairs, "MIN");
  },

  slope: function (ctx, formatValue, buildConditionPairs) {
    const it = ctx.intent || {};
    let xExpr = null,
      yExpr = null;
    const xR = refFromHeaderSpec(ctx, it.x_range);
    const yR = refFromHeaderSpec(ctx, it.y_range);
    if (xR) xExpr = xR.range;
    else if (it.x_range && it.x_range.operation)
      xExpr = _evalSubExprToRange(ctx, formatValue, it.x_range);
    if (yR) yExpr = yR.range;
    else if (it.y_range && it.y_range.operation)
      yExpr = _evalSubExprToRange(ctx, formatValue, it.y_range);
    if (!xExpr || !yExpr)
      return `=ERROR("SLOPE: x_range / y_range 를 찾을 수 없습니다.")`;

    const join = it.join_by || it.join_on;
    if (join) {
      const xKey =
        refFromHeaderSpec(
          ctx,
          join?.left_key || { header: join, sheet: xR?.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: xR?.sheetName });
      const yKey =
        refFromHeaderSpec(
          ctx,
          join?.right_key || { header: join, sheet: yR?.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: yR?.sheetName });
      if (xKey && yKey) {
        yExpr = `XLOOKUP(${xKey.range}, ${yKey.range}, ${yExpr})`;
      }
    }

    const pairs = _collectPairs(ctx, it, buildConditionPairs, formatValue);
    const xF = pairs.length ? _buildFilterCall(xExpr, pairs) : xExpr;
    const yF = pairs.length ? _buildFilterCall(yExpr, pairs) : yExpr;

    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const xK = _buildFilterCall(
          xExpr,
          pairsPlus.split(", ").filter(Boolean),
        );
        const yK = _buildFilterCall(
          yExpr,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `SLOPE(${yK}, ${xK})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }

    return `=SLOPE(${yF}, ${xF})`;
  },

  intercept: function (ctx, formatValue, buildConditionPairs) {
    const it = ctx.intent || {};
    let xExpr = null,
      yExpr = null;
    const xR = refFromHeaderSpec(ctx, it.x_range);
    const yR = refFromHeaderSpec(ctx, it.y_range);
    if (xR) xExpr = xR.range;
    else if (it.x_range && it.x_range.operation)
      xExpr = _evalSubExprToRange(ctx, formatValue, it.x_range);
    if (yR) yExpr = yR.range;
    else if (it.y_range && it.y_range.operation)
      yExpr = _evalSubExprToRange(ctx, formatValue, it.y_range);
    if (!xExpr || !yExpr)
      return `=ERROR("INTERCEPT: x_range / y_range 를 찾을 수 없습니다.")`;

    const join = it.join_by || it.join_on;
    if (join) {
      const xKey =
        refFromHeaderSpec(
          ctx,
          join?.left_key || { header: join, sheet: xR?.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: xR?.sheetName });
      const yKey =
        refFromHeaderSpec(
          ctx,
          join?.right_key || { header: join, sheet: yR?.sheetName },
        ) || refFromHeaderSpec(ctx, { header: join, sheet: yR?.sheetName });
      if (xKey && yKey) {
        yExpr = `XLOOKUP(${xKey.range}, ${yKey.range}, ${yExpr})`;
      }
    }

    const pairs = _collectPairs(ctx, it, buildConditionPairs, formatValue);
    const xF = pairs.length ? _buildFilterCall(xExpr, pairs) : xExpr;
    const yF = pairs.length ? _buildFilterCall(yExpr, pairs) : yExpr;

    if (ctx.intent?.group_by) {
      const keyRef =
        refFromHeaderSpec(ctx, ctx.intent.group_by) ||
        refFromHeaderSpec(ctx, { header: ctx.intent.group_by });
      if (!keyRef) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
      const basePairs = pairs || [];
      const inner = (kSym) => {
        const pairsPlus = basePairs.length
          ? `${basePairs.join(", ")}, ${keyRef.range}, ${kSym}`
          : `${keyRef.range}, ${kSym}`;
        const xK = _buildFilterCall(
          xExpr,
          pairsPlus.split(", ").filter(Boolean),
        );
        const yK = _buildFilterCall(
          yExpr,
          pairsPlus.split(", ").filter(Boolean),
        );
        return `INTERCEPT(${yK}, ${xK})`;
      };
      return _wrapGroupByWithMaker(keyRef, inner);
    }

    return `=INTERCEPT(${yF}, ${xF})`;
  },
};

function _wrapGroupByWithMaker(keyRef, makeInnerWithK) {
  return `=MAP(UNIQUE(${keyRef.range}), LAMBDA(k, ${makeInnerWithK("k")}))`;
}
const _op = (o) => (o ? String(o) : "=");

module.exports = mathStatsFunctionBuilder;
