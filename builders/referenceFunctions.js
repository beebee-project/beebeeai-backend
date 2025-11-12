const formulaUtils = require("../utils/formulaUtils");
const {
  resolveHeaderInSheet,
  rangeFromSpec,
  evalSubIntentToScalar,
  asArray,
} = require("../utils/builderHelpers");

/* =========================
 * 힌트 기반 열 선택
 * ========================= */
function _bestColumnByHint(hint, ctx, role = "lookup") {
  if (!hint) return null;
  const terms = formulaUtils.expandTermsFromText
    ? formulaUtils.expandTermsFromText(hint)
    : new Set(String(hint).split(/\s+/));
  if (formulaUtils.findBestColumnAcrossSheets) {
    return formulaUtils.findBestColumnAcrossSheets(
      ctx.allSheetsData,
      terms,
      role
    );
  }
  return formulaUtils.findColumnInfo
    ? formulaUtils.findColumnInfo(hint, ctx.allSheetsData)
    : null;
}

/* =========================
 * Lookup 유틸
 * ========================= */
function toXMatchMode(m) {
  switch (m) {
    case "exact":
      return 0;
    case "nextSmaller":
      return -1;
    case "nextLarger":
      return 1;
    case "wildcard":
      return 2;
    default:
      return 0;
  }
}

function toMatchType(m) {
  switch (m) {
    case "exact":
      return 0;
    case "nextSmaller":
      return 1; // <= (오름차순)
    case "nextLarger":
      return -1; // >= (내림차순)
    case "wildcard":
      return 0;
    default:
      return 0;
  }
}

function keyWithWildcardIfNeeded(key, mode, env) {
  if (mode !== "wildcard") return key;
  return env === "excel" ? `"*"&${key}&"*"` : `CONCAT("*", ${key}, "*")`;
}

// 단일 조건을 숫자(True=1)로
function buildSingleCondition(range, key, mode, env) {
  if (mode === "wildcard") {
    const pat = keyWithWildcardIfNeeded(key, "wildcard", env);
    return `--ISNUMBER(SEARCH(${pat}, ${range}))`;
  }
  return `--(${range}=${key})`;
}

// 멀티 조건 AND 곱
function buildAndConditions(ranges, keys, modes, env) {
  const parts = ranges.map((rg, i) =>
    buildSingleCondition(rg, keys[i], modes[i] ?? modes[0] ?? "exact", env)
  );
  return parts.join(" * ");
}
const wrapFilter = (range, condProduct) => `FILTER(${range}, ${condProduct})`;

// 근사치 모드가 1개인지 확인
function findApproxPrimaryIndex(modes) {
  let idx = -1;
  modes.forEach((m, i) => {
    if (m === "nextSmaller" || m === "nextLarger") {
      if (idx === -1) idx = i;
      else idx = -2;
    }
  });
  return idx;
}

function buildSecondaryCondProduct(ranges, keys, modes, env, primaryIndex) {
  const parts = [];
  for (let i = 0; i < ranges.length; i++) {
    if (i === primaryIndex) continue;
    parts.push(
      buildSingleCondition(ranges[i], keys[i], modes[i] ?? "exact", env)
    );
  }
  return parts.join(" * ");
}

/* =========================
 * Not-Found 핸들링(무-IFERROR/IFNA)
 * ========================= */
// exact/wildcard 단일키: 존재성 검사 식(가능한 경우만)
function _existsExprSingle(lookupRange, keyArg, mode, env) {
  if (mode === "wildcard") {
    const pat = keyWithWildcardIfNeeded(keyArg, "wildcard", env);
    return `SUMPRODUCT(--ISNUMBER(SEARCH(${pat}, ${lookupRange})))>0`;
  }
  return `SUMPRODUCT(--(${lookupRange}=${keyArg}))>0`;
}
// 멀티키 condProduct: 존재성 검사 식
const _existsExprMulti = (condProduct) => `SUMPRODUCT(${condProduct})>0`;

// 감싸기(가능할 때만). 불가능/근사치 모드면 그냥 core 그대로 반환.
function _wrapIfNotFound(core, canCheck, existsExpr, ifNotFound) {
  if (!ifNotFound || !canCheck) return core;
  return `IF(${existsExpr}, ${core}, ${ifNotFound})`;
}

/* =========================
 * 마지막 매칭 인덱스 (뒤에서 첫번째)
 * ========================= */
function buildLastIndexExpr(ranges, keys, modes, env, orientation) {
  const condProduct =
    ranges.length > 1
      ? buildAndConditions(ranges, keys, modes, env)
      : buildSingleCondition(ranges[0], keys[0], modes[0] ?? "exact", env);

  if (orientation === "horizontal") {
    const relCols = `COLUMN(${ranges[0]})-COLUMN(INDEX(${ranges[0]},1,1))+1`;
    return `LOOKUP(2, 1/(${condProduct}), ${relCols})`;
  }
  const relRows = `ROW(${ranges[0]})-ROW(INDEX(${ranges[0]},1,1))+1`;
  return `LOOKUP(2, 1/(${condProduct}), ${relRows})`;
}

/* =========================
 * 멀티시트 스택(체인 제거 대체)
 * ========================= */
function _candidateSheetOrder(ctx) {
  const names = Object.keys(ctx.allSheetsData || {});
  const pri = [],
    used = new Set();
  const preferred = Array.isArray(ctx.intent?.sheet_priority)
    ? ctx.intent.sheet_priority
    : [];
  for (const s of preferred)
    if (names.includes(s) && !used.has(s)) {
      pri.push(s);
      used.add(s);
    }
  if (ctx.bestReturn?.sheetName && !used.has(ctx.bestReturn.sheetName)) {
    pri.push(ctx.bestReturn.sheetName);
    used.add(ctx.bestReturn.sheetName);
  }
  if (ctx.bestLookup?.sheetName && !used.has(ctx.bestLookup.sheetName)) {
    pri.push(ctx.bestLookup.sheetName);
    used.add(ctx.bestLookup.sheetName);
  }
  for (const n of names) if (!used.has(n)) pri.push(n);
  return pri;
}

function _bestColumnInSheetByHint(hint, sheetName, ctx) {
  if (!hint) return null;
  const info = (ctx.allSheetsData && ctx.allSheetsData[sheetName]) || null;
  if (!info || !info.metaData) return null;
  const m = info.metaData[String(hint)];
  if (m) {
    const col = m.columnLetter,
      sr = info.startRow,
      lr = info.lastDataRow;
    return {
      sheetName,
      header: hint,
      columnLetter: col,
      startRow: sr,
      lastDataRow: lr,
      range: `'${sheetName}'!${col}${sr}:${col}${lr}`,
    };
  }
  if (/^[A-Z]+$/i.test(String(hint))) {
    const col = String(hint).toUpperCase(),
      sr = info.startRow,
      lr = info.lastDataRow;
    return {
      sheetName,
      header: hint,
      columnLetter: col,
      startRow: sr,
      lastDataRow: lr,
      range: `'${sheetName}'!${col}${sr}:${col}${lr}`,
    };
  }
  return null;
}

function _pairsBySheetLoose(ctx, it) {
  const order = _candidateSheetOrder(ctx);
  const pairs = [];
  for (const s of order) {
    let l = null,
      r = null;
    if (it.lookup_hint) {
      const b = _bestColumnInSheetByHint(it.lookup_hint, s, ctx);
      if (b) l = { range: b.range };
    } else if (it.lookup_range) {
      const rr = rangeFromSpec(
        { ...ctx, bestReturn: { sheetName: s } },
        it.lookup_range
      );
      if (rr) l = { range: rr };
    }
    if (it.return_hint) {
      const b = _bestColumnInSheetByHint(it.return_hint, s, ctx);
      if (b) r = { range: b.range };
    } else if (it.return_range) {
      const rr = rangeFromSpec(
        { ...ctx, bestReturn: { sheetName: s } },
        it.return_range
      );
      if (rr) r = { range: rr };
    }
    if (l && r) {
      pairs.push({
        sheetName: s,
        lookupRange: l.range,
        returnRange: r.range,
        orientation: it.orientation || "vertical",
      });
    }
  }
  return pairs;
}

// 범위 스택(Excel: VSTACK, Sheets: {a;b;c})
function _stackRanges(ranges, env, supports) {
  const rs = ranges.filter(Boolean);
  if (!rs.length) return null;
  if (rs.length === 1) return rs[0];
  if (env === "excel" && supports?.VSTACK !== false) {
    return `VSTACK(${rs.join(", ")})`;
  }
  // Google Sheets or Excel 구버전: 배열 리터럴
  return `{${rs.join("; ")}}`;
}

/* =========================
 * 핵심: 1D Lookup (XLOOKUP 우선, INDEX/MATCH 폴백)
 * ========================= */
function buildLookupFormula(args) {
  const {
    key,
    lookupRange,
    returnRange,
    orientation,
    matchMode,
    searchMode,
    ifNotFound,
    nth,
    env,
    supports,
  } = args;

  const keys = asArray(key);
  const ranges = asArray(lookupRange);
  const modes = asArray(matchMode);
  const isMulti = keys.length > 1;
  const firstMode = modes[0] ?? "exact";

  // N번째 매칭(2+) : exact/wildcard만 안전 래핑
  const _nth = Number(nth || 0);
  const _hasApprox = modes.some(
    (m) => m === "nextSmaller" || m === "nextLarger"
  );
  if (_nth >= 2 && !_hasApprox) {
    const condProduct = isMulti
      ? buildAndConditions(ranges, keys, modes, env)
      : buildSingleCondition(ranges[0], keys[0], modes[0] ?? "exact", env);

    if (orientation === "horizontal") {
      const relCols = `COLUMN(${ranges[0]})-COLUMN(INDEX(${ranges[0]},1,1))+1`;
      const nthIdx = `INDEX(FILTER(${relCols}, ${condProduct}), ${_nth})`;
      const core = `INDEX(${returnRange}, 0, ${nthIdx})`;
      const exists = _existsExprMulti(condProduct);
      return _wrapIfNotFound(core, true, exists, ifNotFound);
    } else {
      const relRows = `ROW(${ranges[0]})-ROW(INDEX(${ranges[0]},1,1))+1`;
      const nthIdx = `INDEX(FILTER(${relRows}, ${condProduct}), ${_nth})`;
      const core = `INDEX(${returnRange}, ${nthIdx})`;
      const exists = _existsExprMulti(condProduct);
      return _wrapIfNotFound(core, true, exists, ifNotFound);
    }
  }

  // XLOOKUP 경로
  const has1DReturn = String(returnRange).includes(":");
  const canUseX = supports && supports.XLOOKUP === true && has1DReturn;
  if (canUseX) {
    let smode = "1";
    if (searchMode === "lastToFirst") smode = "-1";
    if (searchMode === "ascending") smode = "2";
    if (searchMode === "descending") smode = "-2";

    if (!isMulti) {
      const m0 = modes[0] ?? "exact";
      const keyArg = keyWithWildcardIfNeeded(keys[0], m0, env);
      const xMatch = toXMatchMode(m0);
      const nf = ifNotFound != null ? `, ${ifNotFound}` : "";
      return `XLOOKUP(${keyArg}, ${ranges[0]}, ${returnRange}${nf}, ${xMatch}, ${smode})`;
    }

    const approxIdx = findApproxPrimaryIndex(modes);
    if (approxIdx === -2) {
      const condProduct = buildAndConditions(ranges, keys, modes, env);
      const nf = ifNotFound != null ? `, ${ifNotFound}` : "";
      return `XLOOKUP(1, ${condProduct}, ${returnRange}${nf}, 0, ${smode})`;
    }

    if (approxIdx >= 0) {
      const sec = buildSecondaryCondProduct(
        ranges,
        keys,
        modes,
        env,
        approxIdx
      );
      const hasSecondary = sec.trim().length > 0;

      const primKey = keyWithWildcardIfNeeded(
        keys[approxIdx],
        modes[approxIdx],
        env
      );
      const xMatch = toXMatchMode(modes[approxIdx]);

      const lookFilt = hasSecondary
        ? wrapFilter(ranges[approxIdx], sec)
        : ranges[approxIdx];
      const retFilt = hasSecondary ? wrapFilter(returnRange, sec) : returnRange;

      const nf = ifNotFound != null ? `, ${ifNotFound}` : "";
      return `XLOOKUP(${primKey}, ${lookFilt}, ${retFilt}${nf}, ${xMatch}, ${smode})`;
    }

    const condProduct = buildAndConditions(ranges, keys, modes, env);
    const nf = ifNotFound != null ? `, ${ifNotFound}` : "";
    return `XLOOKUP(1, ${condProduct}, ${returnRange}${nf}, 0, ${smode})`;
  }

  // INDEX/MATCH 폴백
  if (isMulti) {
    const approxIdx = findApproxPrimaryIndex(modes);
    if (approxIdx === -1 && searchMode === "lastToFirst") {
      const lastIdx = buildLastIndexExpr(ranges, keys, modes, env, orientation);
      const core =
        orientation === "vertical"
          ? `INDEX(${returnRange}, ${lastIdx})`
          : `INDEX(${returnRange}, 0, ${lastIdx})`;
      // lastToFirst는 존재성 검사 어려움 → 그대로 반환
      return core;
    }
    if (approxIdx >= 0) {
      const sec = buildSecondaryCondProduct(
        ranges,
        keys,
        modes,
        env,
        approxIdx
      );
      const hasSecondary = sec.trim().length > 0;
      const lookFilt = hasSecondary
        ? wrapFilter(ranges[approxIdx], sec)
        : ranges[approxIdx];
      const retFilt = hasSecondary ? wrapFilter(returnRange, sec) : returnRange;

      const mType = toMatchType(modes[approxIdx]);
      const matchExpr = `MATCH(${keys[approxIdx]}, ${lookFilt}, ${mType})`;
      const core =
        orientation === "vertical"
          ? `INDEX(${retFilt}, ${matchExpr})`
          : `INDEX(${retFilt}, 0, ${matchExpr})`;
      // 근사치는 안전한 사전검사 어려움 → 그대로
      return core;
    }

    const condProduct = buildAndConditions(ranges, keys, modes, env);
    const matchExpr = `MATCH(1, ${condProduct}, 0)`;
    const core =
      orientation === "vertical"
        ? `INDEX(${returnRange}, ${matchExpr})`
        : `INDEX(${returnRange}, 0, ${matchExpr})`;
    const exists = _existsExprMulti(condProduct);
    return _wrapIfNotFound(core, true, exists, ifNotFound);
  } else {
    if (
      (firstMode === "exact" || firstMode === "wildcard") &&
      searchMode === "lastToFirst"
    ) {
      const lastIdx = buildLastIndexExpr(
        ranges,
        [keys[0]],
        [firstMode],
        env,
        orientation
      );
      const core =
        orientation === "vertical"
          ? `INDEX(${returnRange}, ${lastIdx})`
          : `INDEX(${returnRange}, 0, ${lastIdx})`;
      const exists = _existsExprSingle(ranges[0], keys[0], firstMode, env);
      return _wrapIfNotFound(core, true, exists, ifNotFound);
    }
  }

  // 일반 단일키
  const keyArg = keyWithWildcardIfNeeded(keys[0], firstMode, env);
  const mType = toMatchType(firstMode);
  const matchExpr = `MATCH(${keyArg}, ${ranges[0]}, ${mType})`;
  const core =
    orientation === "vertical"
      ? `INDEX(${returnRange}, ${matchExpr})`
      : `INDEX(${returnRange}, 0, ${matchExpr})`;
  const exists = _existsExprSingle(ranges[0], keyArg, firstMode, env);
  return _wrapIfNotFound(core, true, exists, ifNotFound);
}

/* =========================
 * 2-Way Lookup
 * ========================= */
function buildTwoWayLookupFormula(args) {
  const {
    rowKey,
    colKey,
    rowHeaderRange,
    colHeaderRange,
    dataRange,
    ifNotFound,
    supports,
  } = args;
  if (supports && supports.XLOOKUP) {
    const nf = ifNotFound != null ? `, ${ifNotFound}` : "";
    return `XLOOKUP(${colKey}, ${colHeaderRange}, XLOOKUP(${rowKey}, ${rowHeaderRange}, ${dataRange}${nf}, 0, 1)${nf}, 0, 1)`;
  }
  const rIdx = `MATCH(${rowKey}, ${rowHeaderRange}, 0)`;
  const cIdx = `MATCH(${colKey}, ${colHeaderRange}, 0)`;
  const core = `INDEX(${dataRange}, ${rIdx}, ${cIdx})`;
  // 안전 검사(가능 시)
  const exists = `AND(ISNUMBER(${rIdx}), ISNUMBER(${cIdx}))`;
  return ifNotFound ? `IF(${exists}, ${core}, ${ifNotFound})` : core;
}

/* =========================
 * 공개 빌더
 * ========================= */
const referenceFunctionBuilder = {
  _fv(val, opts = {}) {
    const s = String(val ?? "");
    if (opts.forceText === true) {
      if (s.startsWith('"') && s.endsWith('"')) return s;
      return `"${s.replace(/"/g, '""')}"`;
    }
    if (/^[A-Z]+\d+$/i.test(s)) return s; // A1
    if (/^-?\d+(\.\d+)?$/.test(s)) return s; // number
    if (s.startsWith('"') && s.endsWith('"')) return s; // already quoted
    return `"${s.replace(/"/g, '""')}"`;
  },

  lookup(ctx, formatValue) {
    const { bestReturn, bestLookup, intent } = ctx;
    const sheetName = bestReturn.sheetName;
    const returnRange = `'${sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}:${bestReturn.columnLetter}${bestReturn.lastDataRow}`;
    const lookupRange = `'${sheetName}'!${bestLookup.columnLetter}${bestLookup.startRow}:${bestLookup.columnLetter}${bestLookup.lastDataRow}`;
    const lookupValue = formatValue(intent.lookup_value);
    // 존재성 검사(정확일치 가정)
    const exists = `SUMPRODUCT(--(${lookupRange}=${lookupValue}))>0`;
    const core = `INDEX(${returnRange}, MATCH(${lookupValue}, ${lookupRange}, 0))`;
    return `=${_wrapIfNotFound(
      core,
      true,
      exists,
      intent.value_if_not_found
        ? JSON.stringify(String(intent.value_if_not_found))
        : null
    )}`;
  },

  xlookup(ctx, formatValue) {
    const it = ctx.intent || {};
    const FV = (v, opts) =>
      formatValue
        ? formatValue(v, opts)
        : referenceFunctionBuilder._fv(v, opts);

    // lookup_value
    let lookupValue =
      it.lookup_value &&
      typeof it.lookup_value === "object" &&
      it.lookup_value.operation
        ? evalSubIntentToScalar(ctx, FV, it.lookup_value)
        : it.lookup_value != null
        ? FV(it.lookup_value, { forceText: true })
        : null;
    if (!lookupValue) return `=ERROR("XLOOKUP: lookup_value가 없습니다.")`;

    // lookup/return 열(반환 시트 기준 정렬)
    const bestLookup = _bestColumnByHint(it.lookup_hint, ctx, "lookup");
    const bestReturn = _bestColumnByHint(it.return_hint, ctx, "return");
    if (!bestLookup || !bestReturn)
      return `=ERROR("XLOOKUP: lookup/return 열을 찾지 못했습니다.")`;

    const primarySheet = bestReturn.sheetName || bestLookup.sheetName;
    const lookCol =
      bestLookup.sheetName === primarySheet
        ? bestLookup
        : resolveHeaderInSheet(
            bestLookup.header || it.lookup_hint,
            primarySheet,
            ctx
          );
    if (!lookCol)
      return `=ERROR("XLOOKUP: return 시트에서 lookup 키 열을 찾지 못했습니다.")`;

    const lookupRange = `'${primarySheet}'!${lookCol.columnLetter}${lookCol.startRow}:${lookCol.columnLetter}${lookCol.lastDataRow}`;
    const returnRange = `'${primarySheet}'!${bestReturn.columnLetter}${bestReturn.startRow}:${bestReturn.columnLetter}${bestReturn.lastDataRow}`;

    // if_not_found, match/search
    const ifNF =
      it.value_if_not_found != null ? `, ${FV(it.value_if_not_found)}` : "";
    const mm = (() => {
      const v = it.match_mode;
      if (v == null) return null;
      if (typeof v === "number") return v;
      const s = String(v).toLowerCase();
      if (s === "exact") return 0;
      if (s === "approx" || s === "lte" || s === "lteq") return -1;
      if (s === "gte" || s === "gteq") return 1;
      if (s === "wildcard") return 2;
      return 0;
    })();
    const sm = (() => {
      const v = it.search_mode;
      if (v == null) return null;
      if (typeof v === "number") return v;
      const s = String(v).toLowerCase();
      if (s === "first") return 1;
      if (s === "last") return -1;
      if (s === "asc") return 2;
      if (s === "desc") return -2;
      return 1;
    })();

    // 멀티키 조합
    let lookupArray = lookupRange;
    if (Array.isArray(it.multi_keys) && it.multi_keys.length > 0) {
      const cols = it.multi_keys
        .map((k) => _bestColumnByHint(k.hint, ctx, "lookup"))
        .filter(Boolean)
        .map(
          (c) =>
            `'${c.sheetName}'!${c.columnLetter}${c.startRow}:${c.columnLetter}${c.lastDataRow}`
        );
      if (cols.length > 0) {
        lookupArray = cols.map((c) => `(${c})`).join(`&"|"&`);
        if (!it.lookup_value) {
          const keyVals = it.multi_keys
            .map((k) => (k.value != null ? FV(k.value) : `""`))
            .join(`&"|"&`);
          lookupValue = keyVals;
        }
      }
    }

    const mmArg = mm != null ? `, ${mm}` : "";
    const smArg = sm != null ? `, ${sm}` : "";
    return `=XLOOKUP(${lookupValue}, ${lookupArray}, ${returnRange}${ifNF}${mmArg}${smArg})`;
  },

  hlookup(ctx, formatValue) {
    const { intent } = ctx;
    return `=HLOOKUP(${formatValue(intent.lookup_value)}, ${
      intent.lookup_range
    }, ${intent.row_index}, FALSE)`;
  },

  vlookup(ctx) {
    const it = ctx.intent || {};
    const fmt = (v, opts) =>
      ctx.formulaBuilder && ctx.formulaBuilder._formatValue
        ? ctx.formulaBuilder._formatValue(v, opts)
        : referenceFunctionBuilder._fv(v, opts);

    const buildRangeFromBest = (best) =>
      `'${best.sheetName}'!${best.columnLetter}${best.startRow}:${best.columnLetter}${best.lastDataRow}`;

    const bestReturn = ctx.bestReturn;
    const bestLookup = ctx.bestLookup;

    const returnRange = bestReturn
      ? buildRangeFromBest(bestReturn)
      : it.return_range || it.return_array || it.returnRange;

    const lookupRange = bestLookup
      ? buildRangeFromBest(bestLookup)
      : it.lookup_range || it.lookup_array || it.lookupRange;

    // 2-way
    if (it.two_way) {
      const tw = it.two_way;
      const nf = it.value_if_not_found
        ? JSON.stringify(String(it.value_if_not_found))
        : undefined;
      const formula = buildTwoWayLookupFormula({
        rowKey: fmt(tw.row_value ?? tw.rowKey ?? ""),
        colKey: fmt(tw.col_value ?? tw.colKey ?? ""),
        rowHeaderRange: tw.row_range ?? tw.rowHeaderRange,
        colHeaderRange: tw.col_range ?? tw.colHeaderRange,
        dataRange: tw.data_range ?? tw.dataRange,
        ifNotFound: nf,
        supports: { XLOOKUP: ctx.supports?.XLOOKUP !== false },
      });
      return `=${formula}`;
    }

    // 멀티시트 스캔 → 스택 후 한 번에 조회
    const wantScan = it.multi_sheet_scan === true || it.scan_sheets === true;
    const canScanLoosely =
      !!ctx.allSheetsData &&
      (it.lookup_hint || it.return_hint || it.lookup_range || it.return_range);
    if (wantScan && canScanLoosely) {
      const pairs = _pairsBySheetLoose(ctx, it);
      const env =
        ctx.platform === "excel" || ctx.env === "excel" ? "excel" : "gsheets";
      const lookupStack = _stackRanges(
        pairs.map((p) => p.lookupRange),
        env,
        ctx.supports || {}
      );
      const returnStack = _stackRanges(
        pairs.map((p) => p.returnRange),
        env,
        ctx.supports || {}
      );
      if (!lookupStack || !returnStack)
        return `=ERROR("멀티시트 스캔에 사용할 범위를 찾지 못했습니다.")`;

      const matchMode =
        it.match_mode === "wildcard"
          ? "wildcard"
          : it.match_mode === "nextsmaller" || it.match_mode === -1
          ? "nextSmaller"
          : it.match_mode === "nextlarger" || it.match_mode === 1
          ? "nextLarger"
          : "exact";

      const keyRef =
        it.lookup_value_cell ||
        it.lookup_cell ||
        it.key_cell ||
        it.keyRef ||
        it.lookup_value ||
        it.key;
      const keyExpr = keyRef ? fmt(keyRef, { forceText: true }) : "";

      const inner = buildLookupFormula({
        key: keyExpr,
        lookupRange: lookupStack,
        returnRange: returnStack,
        nth: Number(it.nth ?? it.ordinal ?? 0),
        orientation: it.orientation || "vertical",
        matchMode: matchMode,
        searchMode:
          it.search_mode === "last"
            ? "lastToFirst"
            : it.search_mode === "asc"
            ? "ascending"
            : it.search_mode === "desc"
            ? "descending"
            : "firstToLast",
        env,
        supports: {
          XLOOKUP: ctx.supports?.XLOOKUP !== false,
          VSTACK: ctx.supports?.VSTACK !== false,
        },
        ifNotFound: it.value_if_not_found
          ? JSON.stringify(String(it.value_if_not_found))
          : null,
      });
      return `=${inner}`;
    }

    // 일반 경로(단일/멀티키)
    const keyRef =
      it.lookup_value_cell ||
      it.lookup_cell ||
      it.key_cell ||
      it.keyRef ||
      it.lookup_value ||
      it.key;
    const keyRefs = Array.isArray(it.keyRefs)
      ? it.keyRefs
      : it.multi_keys
      ? it.multi_keys.map((k) => fmt(k.value, { forceText: true }))
      : null;

    const lookupRanges = Array.isArray(it.lookupRanges)
      ? it.lookupRanges
      : it.multi_keys
      ? it.multi_keys.map((_k) => lookupRange)
      : null;

    const matchMode =
      it.match_mode === "wildcard"
        ? "wildcard"
        : it.match_mode === "nextsmaller" || it.match_mode === -1
        ? "nextSmaller"
        : it.match_mode === "nextlarger" || it.match_mode === 1
        ? "nextLarger"
        : "exact";

    const env =
      ctx.platform === "excel" || ctx.env === "excel" ? "excel" : "gsheets";
    const inner = buildLookupFormula({
      key: keyRefs?.length
        ? keyRefs
        : keyRef
        ? fmt(keyRef, { forceText: true })
        : "",
      lookupRange: lookupRanges?.length ? lookupRanges : lookupRange,
      returnRange,
      nth: Number(it.nth ?? it.ordinal ?? 0),
      orientation: it.orientation || "vertical",
      matchMode: keyRefs?.length
        ? Array.isArray(it.wildcard)
          ? it.wildcard.map((w) => (w ? "wildcard" : "exact"))
          : it.wildcard
          ? "wildcard"
          : "exact"
        : matchMode,
      searchMode: it.search_mode === "last" ? "lastToFirst" : "firstToLast",
      env,
      supports: { XLOOKUP: ctx.supports?.XLOOKUP !== false },
      ifNotFound: it.value_if_not_found
        ? JSON.stringify(String(it.value_if_not_found))
        : null,
    });
    return `=${inner}`;
  },

  offset(ctx) {
    const it = ctx.intent || {};
    const ref = it.reference || "A1";
    const rows = it.rows || 0,
      cols = it.cols || 0;
    const h = it.height || 1,
      w = it.width || 1;
    return `=OFFSET(${ref}, ${rows}, ${cols}, ${h}, ${w})`;
  },

  indirect(ctx) {
    const it = ctx.intent || {};
    if (it.sheet && it.cell)
      return `=INDIRECT("'"&${it.sheet}&"'!"&${it.cell})`;
    return `=INDIRECT(${it.target_cell || '"A1"'})`;
  },

  address(ctx) {
    const it = ctx.intent || {};
    const row = it.row || 1,
      col = it.column || 1;
    const abs = it.abs || 1;
    const sheet = it.sheet ? `, TRUE, "${it.sheet}"` : "";
    return `=ADDRESS(${row}, ${col}, ${abs}${sheet})`;
  },

  formulatext(ctx) {
    const ref = ctx.intent.target_cell || "A1";
    // 에러시 그대로 에러 표시 (IFERROR/IFNA 제거 요구)
    return `=FORMULATEXT(${ref})`;
  },

  row(ctx) {
    const ref = ctx.intent.target_cell || "";
    return `=ROW(${ref})`;
  },
  column(ctx) {
    const ref = ctx.intent.target_cell || "";
    return `=COLUMN(${ref})`;
  },
  rows(ctx) {
    const ref =
      ctx.intent.target_range ||
      (ctx.bestReturn ? ctx.bestReturn.range : "A1:A10");
    return `=ROWS(${ref})`;
  },
  columns(ctx) {
    const ref =
      ctx.intent.target_range ||
      (ctx.bestReturn ? ctx.bestReturn.range : "A1:D1");
    return `=COLUMNS(${ref})`;
  },
};

module.exports = referenceFunctionBuilder;
