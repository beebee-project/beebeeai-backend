const formulaUtils = require("../utils/formulaUtils");
const {
  refFromHeaderSpec,
  rangeFromSpec,
  evalSubIntentToScalar,
} = require("../utils/builderHelpers");

// --- NEW: 객체/배열 호환용 시트 배열 변환
function _sheetsArray(allSheetsData) {
  if (!allSheetsData) return [];
  return Array.isArray(allSheetsData)
    ? allSheetsData
    : Object.values(allSheetsData);
}

// =============================
// 공통 레퍼런스/유틸
// =============================

// 헤더 문자열 → 셀/범위 참조
function _ref(headerText, ctx) {
  if (!headerText || !ctx?.allSheetsData) return null;
  const sheets = _sheetsArray(ctx.allSheetsData);
  const term = formulaUtils.expandTermsFromText(headerText);
  const col = formulaUtils.findBestColumnAcrossSheets(sheets, term, "lookup");
  if (!col) return null;

  const sheetQuoted = `'${col.sheetName}'!`;
  const c = col.columnLetter;
  const start = col.startRow || 2;
  const end = col.lastDataRow || col.rowCount || 1048576;

  return {
    sheetName: col.sheetName,
    col: c,
    start,
    end,
    cell: `${sheetQuoted}${c}${start}`,
    range: `${sheetQuoted}${c}${start}:${c}${end}`,
  };
}
const _op = (o) => (o ? String(o) : "=");

function _refInSheet(headerText, sheetHint, ctx) {
  if (!ctx?.allSheetsData) return null;
  const sheets = _sheetsArray(ctx.allSheetsData);
  const term = formulaUtils.expandTermsFromText(headerText);
  const sheet = sheets.find((s) => s.sheetName === sheetHint);
  if (!sheet) return null;

  // 시트 단일 탐색
  const col =
    formulaUtils.findBestColumnInSingleSheet?.(sheet, term) ||
    formulaUtils.findBestColumnAcrossSheets([sheet], term, "lookup");
  if (!col) return null;

  const s = `'${col.sheetName}'!`,
    c = col.columnLetter;
  const st = col.startRow || 2,
    en = col.lastDataRow || col.rowCount || 1048576;
  return {
    sheetName: col.sheetName,
    col: c,
    start: st,
    end: en,
    cell: `${s}${c}${st}`,
    range: `${s}${c}${st}:${c}${en}`,
  };
}

// {"A","B"} 상수 배열
function _toArrayConst(list = []) {
  const esc = (s) => `"${String(s).replace(/"/g, '""')}"`;
  return `{${list.map(esc).join(",")}}`;
}

// date_relative → 수식 문자열 (TODAY() / DATE(yyyy,mm,dd) [+offset])
function _dateExpr(dr) {
  if (!dr) return "TODAY()";
  const baseStr = String(dr.base || "today");
  const base =
    baseStr.toLowerCase() === "today"
      ? "TODAY()"
      : (function () {
          const m = baseStr.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
          if (!m) return "TODAY()";
          return `DATE(${m[1]},${m[2]},${m[3]})`;
        })();
  const off = Number(dr.offset_days || 0);
  return off ? `${base}+${off}` : base;
}

// 스칼라(상수/셀/표현식) 문자열 생성
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

function _buildConditionString(condition, formatValue, ctx) {
  if (!condition) return "TRUE";

  if (condition.logical_operator && Array.isArray(condition.conditions)) {
    const operator = condition.logical_operator.toUpperCase();
    const sub = condition.conditions
      .map((c) => _buildConditionString(c, formatValue, ctx))
      .join(", ");
    return `${operator}(${sub})`;
  }

  if (condition.target) {
    let left;
    if (typeof condition.target === "object" && condition.target !== null) {
      if (condition.target.header) {
        const ref = refFromHeaderSpec(ctx, condition.target);
        left = ref ? ref.cell : `ERROR("타겟 열을 찾을 수 없음")`;
      } else if (condition.target.operation) {
        const sub = evalSubIntentToScalar(ctx, formatValue, condition.target);
        left =
          sub ||
          `ERROR("연계 수식 '${condition.target.operation}'을 만들 수 없음")`;
      } else {
        left = condition.target.cell || condition.target;
      }
    } else {
      left = condition.target?.cell || condition.target;
    }

    let right;
    if (typeof condition.value === "object" && condition.value !== null) {
      if (condition.value.header) {
        const ref = refFromHeaderSpec(ctx, condition.value);
        right = ref ? ref.cell : `ERROR("비교 열을 찾을 수 없음")`;
      } else if (condition.value.operation) {
        const sub = evalSubIntentToScalar(ctx, formatValue, condition.value);
        right =
          sub ||
          `ERROR("연계 수식 '${condition.value.operation}'을 만들 수 없음")`;
      } else {
        right = formatValue(condition.value);
      }
    } else {
      right = formatValue(condition.value);
    }
    return `${left}${_op(condition.operator)}${right}`;
  }
  return "TRUE";
}

// =============================
// 고급 헬퍼: 조인 MAP
// =============================
function _joinMap2(leftValRef, leftKeyRef, rightKeyRef, rightValRef, bodyFn) {
  const lV = leftValRef.range,
    lK = leftKeyRef.range;
  const rK = rightKeyRef.range,
    rV = rightValRef.range;
  const body = bodyFn("l", "rv");
  return `=MAP(${lV}, ${lK}, LAMBDA(l, k, LET(rv, XLOOKUP(k, ${rK}, ${rV}), ${body})))`;
}

function _joinMapN(leftValRef, leftKeyRef, rights, bodyFn) {
  const lV = leftValRef.range,
    lK = leftKeyRef.range;
  const lets = rights
    .map(
      (r, i) =>
        `${r.as || `r${i + 1}`}, XLOOKUP(k, ${r.key.range}, ${r.val.range})`,
    )
    .join(", ");
  const symNames = rights.map((r, i) => r.as || `r${i + 1}`);
  const body = bodyFn("l", ...symNames);
  return `=MAP(${lV}, ${lK}, LAMBDA(l, k, LET(${lets}, ${body})))`;
}

// =============================
// 패턴 빌더 (between, in, exists 등)
// =============================
function between(it) {
  return function (ctx, formatValue) {
    const inclusive = it?.inclusive !== false; // 기본 포함
    const leftRange = rangeFromSpec(ctx, it.left);
    const leftScalar =
      evalSubIntentToScalar(ctx, formatValue, it.left) ||
      (typeof it.left === "string" ? it.left : null);
    const minScalar =
      evalSubIntentToScalar(ctx, formatValue, it.min) ||
      (typeof it.min === "string" ? it.min : null);
    const maxScalar =
      evalSubIntentToScalar(ctx, formatValue, it.max) ||
      (typeof it.max === "string" ? it.max : null);

    const ge = inclusive ? ">=" : ">";
    const le = inclusive ? "<=" : "<";

    const minExpr = /[\(\)]/.test(minScalar || "")
      ? `(${minScalar})`
      : minScalar
        ? formatValue(minScalar)
        : formatValue(it.min);
    const maxExpr = /[\(\)]/.test(maxScalar || "")
      ? `(${maxScalar})`
      : maxScalar
        ? formatValue(maxScalar)
        : formatValue(it.max);

    if (leftRange) {
      const expr = `BYROW(${leftRange}, LAMBDA(x, AND(x${ge}${minExpr}, x${le}${maxExpr})))`;
      return `=${expr}`;
    }
    const lx =
      leftScalar ||
      (typeof it.left === "number" ? String(it.left) : formatValue(it.left));
    return `=AND(${lx}${ge}${minExpr}, ${lx}${le}${maxExpr})`;
  };
}

function inn(it) {
  // "in" 예약어 회피: 외부에서는 operation:"in"으로 라우팅
  return function (ctx, formatValue) {
    const leftRange = rangeFromSpec(ctx, it.left);
    const leftScalar =
      evalSubIntentToScalar(ctx, formatValue, it.left) ||
      (typeof it.left === "string" ? it.left : null);
    let setRange = rangeFromSpec(ctx, it.set);
    let setArray = null;
    if (!setRange) {
      if (Array.isArray(it.set)) {
        const esc = (s) => `"${String(s).replace(/"/g, '""')}"`;
        setArray = `{${it.set.map(esc).join(",")}}`;
      } else if (typeof it.set === "string" && /^\{.*\}$/.test(it.set.trim())) {
        setArray = it.set.trim();
      }
    }
    const container = setRange || setArray;
    if (!container) return `=ERROR("IN: set(범위/배열)을 해석할 수 없습니다.")`;

    if (leftRange) {
      const expr = `BYROW(${leftRange}, LAMBDA(x, ISNUMBER(XMATCH(x, ${container}, 0)) ))`;
      return `=${expr}`;
    }
    const lx =
      leftScalar ||
      (typeof it.left === "number" ? String(it.left) : formatValue(it.left));
    return `=ISNUMBER(XMATCH(${lx}, ${container}, 0))`;
  };
}

function exists(it) {
  return function (ctx, formatValue) {
    const r = rangeFromSpec(ctx, it.target || it.left || it.value);
    const s =
      evalSubIntentToScalar(
        ctx,
        formatValue,
        it.target || it.left || it.value,
      ) ||
      it.target_cell ||
      null;
    if (r) return `=BYROW(${r}, LAMBDA(x, NOT(ISBLANK(x))))`;
    if (s) return `=NOT(ISBLANK(${s}))`;
    return `=ERROR("EXISTS: 대상이 없습니다.")`;
  };
}

function validate(it) {
  return function (ctx, formatValue, buildConditionPairs) {
    let condExpr = null;
    if (typeof it.condition === "object" && it.condition.operation) {
      const se = evalSubIntentToScalar(ctx, formatValue, it.condition);
      if (se) condExpr = se;
    }
    if (!condExpr && typeof buildConditionPairs === "function") {
      const pairs = buildConditionPairs(ctx) || [];
      if (pairs.length === 2) condExpr = `${pairs[0]}, ${pairs[1]}`;
    }
    if (!condExpr) condExpr = "TRUE";
    const msg = it.message != null ? formatValue(it.message) : '""';
    return `=IF(${condExpr}, "", ${msg})`;
  };
}

function compute(it) {
  return function (ctx, formatValue) {
    if (it?.expr && typeof it.expr === "object" && it.expr.operation) {
      const se = evalSubIntentToScalar(ctx, formatValue, it.expr);
      if (se) return "=" + se;
    }
    if (typeof it?.expr === "string") return "=" + it.expr;
    return `=ERROR("COMPUTE: expr가 필요합니다.")`;
  };
}

function _rewriteCondToParams(condition, headerList, formatValue) {
  const idx = new Map(headerList.map((h, i) => [String(h), `k${i + 1}`]));
  function walk(n) {
    if (!n) return "TRUE";
    if (n.logical_operator && Array.isArray(n.conditions)) {
      const op = String(n.logical_operator || "AND").toUpperCase();
      return `${op}(${n.conditions.map(walk).join(", ")})`;
    }
    if (n.target) {
      const left = n.target?.header
        ? idx.get(String(n.target.header)) || `ERROR("헤더")`
        : n.target.cell || n.target;
      const right =
        n.value && n.value.header
          ? idx.get(String(n.value.header)) || `ERROR("헤더")`
          : formatValue(n.value);
      const op = n.operator || "=";
      return `${left}${op}${right}`;
    }
    return "TRUE";
  }
  return walk(condition);
}

function buildIfArithTwoColsVsCol(_ctx, _formatValue) {
  return null;
}

// =============================
// 논리 함수 본체
// =============================
const logicalFunctionBuilder = {
  // 메인 IF 라우터
  if: (ctx, formatValue) => {
    const it = ctx.intent || {};
    const cond = it.condition || {};

    // N-헤더 자동 벡터화 (3개 이상)
    const headers = new Set();
    (function walk(n) {
      if (!n) return;
      if (n.logical_operator && Array.isArray(n.conditions))
        n.conditions.forEach(walk);
      else {
        n?.target?.header && headers.add(n.target.header);
        n?.value?.header && headers.add(n.value.header);
      }
    })(cond);
    if (headers.size >= 3) {
      const refs = [...headers].map((h) => refFromHeaderSpec(ctx, h));
      if (refs.some((r) => !r)) return `=ERROR("조건 열을 찾을 수 없습니다.")`;
      const aligned = refs
        .map((r) => `(${formulaUtils.ALIGN_TO(r.range, "col")})`)
        .join(", ");
      const args = [...headers].map((_, i) => `k${i + 1}`).join(", ");
      const headerList = [...headers];
      const condExpr = _rewriteCondToParams(cond, headerList, formatValue);
      const t = formatValue(it.value_if_true ?? "");
      const f = formatValue(it.value_if_false ?? "");
      const core = `MAP(${aligned}, LAMBDA(${args}, IF(${condExpr}, ${t}, ${f})))`;
      return `=${core}`;
    }

    // 빠른 검증
    if (
      !cond &&
      !it.thresholds?.length &&
      !it.in_values?.length &&
      !it.date_relative &&
      !it.between &&
      !it.exists &&
      !it.validate &&
      !it.compute &&
      !it.row_selector
    ) {
      return `=ERROR("IF 함수에 필요한 조건 또는 패턴 정보가 없습니다.")`;
    }

    // 1) 산술식 (A+B) op C
    const arith = buildIfArithTwoColsVsCol(ctx, formatValue);
    if (arith) return arith;

    // 2) 동일 헤더 AND/OR (BETWEEN 포함)
    const sameHeadLogic = buildIfVectorSameHeaderLogic(ctx, formatValue);
    if (sameHeadLogic) return sameHeadLogic;

    // 3) 두 헤더 AND/OR
    const twoHdrLogic = buildIfVectorTwoHeadersLogic(ctx, formatValue);
    if (twoHdrLogic) return twoHdrLogic;

    // 4) 단일 행 선택자
    if (it.row_selector && ctx.allSheetsData) {
      const rowRes = buildIfRow(ctx, formatValue);
      if (rowRes) return rowRes;
    }

    // 5) 특수 패턴
    if (it.thresholds?.length) {
      const f = buildIfThresholds(ctx, formatValue);
      if (f) return f;
    }
    if (Array.isArray(it.in_values) && it.in_values.length) {
      const f = buildIfInList(ctx, formatValue);
      if (f) return f;
    }
    if (it.between && (it.between.min != null || it.between.max != null)) {
      const f = buildIfBetween(ctx, formatValue);
      if (f) return f;
    }
    if (it.date_relative) {
      const f = buildIfDateRelative(ctx, formatValue);
      if (f) return f;
    }
    if (it.exists) {
      const f = buildIfExists(ctx, formatValue);
      if (f) return f;
    }
    if (it.validate) {
      const f = buildIfValidate(ctx, formatValue);
      if (f) return f;
    }
    if (it.compute) {
      const f = buildIfCompute(ctx, formatValue);
      if (f) return f;
    }

    // 6) 일반 전행 비교
    const leftH = cond?.target?.header;
    const rightH = cond?.value?.header;
    const scopeAll =
      String(it.scope || "").toLowerCase() === "all" ||
      (!it.row_selector && (leftH || rightH));

    if (scopeAll && leftH && rightH)
      return buildIfVectorTwoCols(ctx, formatValue);
    if (scopeAll && leftH && !rightH)
      return buildIfVectorColConst(ctx, formatValue);

    // 7) compute 폴백
    if (it.compute) {
      const comp = buildIfCompute(ctx, formatValue);
      if (comp) return comp;
    }

    // 8) 단일셀 폴백
    const condStr = _buildConditionString(cond, formatValue, ctx);
    if (condStr.includes("ERROR(")) {
      return `=ERROR("조건에 필요한 열을 파일에서 찾을 수 없습니다.")`;
    }
    const t = formatValue(it.value_if_true ?? "");
    const f = formatValue(it.value_if_false ?? "");
    return `=IF(${condStr}, ${t}, ${f})`;
  },

  // 명시적 IFERROR/IFNA 연산자(사용자가 직접 요청한 경우에만 사용)
  iferror: (ctx, formatValue, formulaBuilder) => {
    const { intent } = ctx;
    const errorVal = formatValue(intent.value_if_error ?? "");
    if (
      !intent.value ||
      typeof intent.value !== "object" ||
      !intent.value.operation
    ) {
      const targetCell = intent.target_cell || "A1";
      return `=IFERROR(${targetCell}, ${errorVal})`;
    }
    const valueOperation = intent.value.operation;
    if (valueOperation === "iferror") {
      return `=ERROR("IFERROR 함수는 자기 자신을 포함할 수 없습니다.")`;
    }
    const innerBuilder = formulaBuilder[valueOperation];
    if (!innerBuilder) {
      return `=ERROR("알 수 없는 함수 '${valueOperation}'를 IFERROR로 감쌀 수 없습니다.")`;
    }
    const innerCtx = { ...ctx, intent: intent.value };
    const valueFormula = innerBuilder
      .call(
        formulaBuilder,
        innerCtx,
        formatValue,
        formulaBuilder._buildConditionPairs,
      )
      .substring(1);
    return `=IFERROR(${valueFormula}, ${errorVal})`;
  },

  ifna: (ctx, formatValue, formulaBuilder) => {
    const { intent } = ctx;
    const valueOperation = intent.value?.operation;
    const errorVal = `"${intent.value_if_na || ""}"`;
    if (valueOperation && formulaBuilder[valueOperation]) {
      const valueFormula = formulaBuilder[valueOperation](
        ctx,
        formatValue,
      ).substring(1);
      return `=IFNA(${valueFormula}, ${errorVal})`;
    }
    return `=IFNA(A1, ${errorVal})`;
  },

  true: () => `=TRUE()`,
  false: () => `=FALSE()`,
};

// ===== AND/OR/NOT 및 IS* 계열 =====
function _buildLogicVector(ctx, formatValue, logicalOp) {
  const it = ctx.intent;
  const c = { logical_operator: logicalOp, conditions: it.conditions || [] };

  // 1) 동일 헤더 트리 → BYROW
  const same = buildIfVectorSameHeaderLogic(
    {
      ...ctx,
      intent: { condition: c, value_if_true: "TRUE", value_if_false: "FALSE" },
    },
    (x) => x,
  );
  if (same) return `=${same.slice(1)}`.replace(/IF\(/g, "");

  // 2) 두 헤더 트리 → MAP
  if ((c.conditions || []).length === 2) {
    const asIf = buildIfVectorTwoHeadersLogic(
      {
        ...ctx,
        intent: {
          condition: c,
          value_if_true: "TRUE",
          value_if_false: "FALSE",
        },
      },
      (x) => x,
    );
    if (asIf) return `=${asIf.slice(1)}`.replace(/IF\(/g, "");
  }
  return null;
}

logicalFunctionBuilder.and = (ctx, formatValue) => {
  const v = _buildLogicVector(ctx, formatValue, "AND");
  if (v) return v;
  const conditionStr = _buildConditionString(
    { logical_operator: "AND", conditions: ctx.intent.conditions || [] },
    formatValue,
    ctx,
  );
  return `=${conditionStr}`;
};

logicalFunctionBuilder.or = (ctx, formatValue) => {
  const v = _buildLogicVector(ctx, formatValue, "OR");
  if (v) return v;
  const conditionStr = _buildConditionString(
    { logical_operator: "OR", conditions: ctx.intent.conditions || [] },
    formatValue,
    ctx,
  );
  return `=${conditionStr}`;
};

logicalFunctionBuilder.not = (ctx, formatValue) => {
  const it = ctx.intent;
  if (it.conditions?.length) {
    const inner = logicalFunctionBuilder.and(
      { ...ctx, intent: { conditions: it.conditions } },
      formatValue,
    );
    return `=NOT(${inner.slice(1)})`;
  }
  const c = it.condition || {};
  if (c?.target?.header && (c.value !== undefined || c.value === 0)) {
    const L = _ref(c.target.header, ctx);
    if (L) {
      return `=BYROW(${L.range}, LAMBDA(l, NOT(l${_op(c.operator)}${formatValue(
        c.value,
      )})))`;
    }
  }
  const condStr = _buildConditionString(c, formatValue, ctx);
  return `=NOT(${condStr})`;
};

function _isUnaryVector(op) {
  return (ctx, formatValue) => {
    const it = ctx.intent;
    const h = it.expression?.header || it.header_hint;
    if (h) {
      const R = refFromHeaderSpec(ctx, h);
      if (R) return `=BYROW(${R.range}, LAMBDA(l, ${op}(l)))`;
    }
    const cell = it.expression?.cell || it.target_cell || "A1";
    return `=${op}(${cell})`;
  };
}

logicalFunctionBuilder.isblank = _isUnaryVector("ISBLANK");
logicalFunctionBuilder.isnumber = _isUnaryVector("ISNUMBER");
logicalFunctionBuilder.istext = _isUnaryVector("ISTEXT");
logicalFunctionBuilder.iserror = _isUnaryVector("ISERROR");
logicalFunctionBuilder.iserr = _isUnaryVector("ISERR");
logicalFunctionBuilder.isna = _isUnaryVector("ISNA");

logicalFunctionBuilder.switch = (ctx, formatValue) => {
  const it = ctx.intent;
  const def = it.default != null ? formatValue(it.default) : `""`;

  // 3) 테이블 매핑
  if (it.map_table?.key_header && it.map_table?.value_header) {
    const K = refFromHeaderSpec(ctx, it.map_table.key_header);
    const V = refFromHeaderSpec(ctx, it.map_table.value_header);
    if (!K || !V) return `=ERROR("SWITCH: 매핑 표 열을 찾을 수 없습니다.")`;

    // 행 선택자면 단일 XLOOKUP (IFERROR 제거 → XLOOKUP 4번째 인수 사용)
    if (it.row_selector) {
      const keyCol = refFromHeaderSpec(ctx, it.row_selector.hint);
      if (!keyCol) return `=ERROR("행 선택자 열을 찾을 수 없습니다.")`;
      const keyVal = formatValue(it.row_selector.value);
      const expr = `XLOOKUP(${keyVal}, ${keyCol.range}, ${K.range})`;
      return `=XLOOKUP(${expr}, ${K.range}, ${V.range}, ${def})`;
    }

    // 전행: BYROW + XLOOKUP (IFERROR 제거)
    const exprHeader = it.expression?.header;
    if (exprHeader) {
      const E = refFromHeaderSpec(ctx, exprHeader);
      if (!E) return `=ERROR("SWITCH: 표현식 열을 찾을 수 없습니다.")`;
      return `=BYROW(${E.range}, LAMBDA(l, XLOOKUP(l, ${K.range}, ${V.range}, ${def})))`;
    }
  }

  // 1) BYROW + SWITCH
  const exprH = it.expression?.header;
  if (exprH) {
    const R = _ref(exprH, ctx);
    if (!R) return `=ERROR("SWITCH: 표현식 열을 찾을 수 없습니다.")`;
    const casesStr = (it.cases || [])
      .map((c) => `${formatValue(c.value)}, ${formatValue(c.result)}`)
      .join(", ");
    return `=BYROW(${R.range}, LAMBDA(l, SWITCH(l, ${casesStr}${
      it.default != null ? `, ${def}` : ""
    })))`;
  }

  // 2) 단일
  const exprCell = it.expression?.cell || it.target_cell || "A1";
  const casesStr = (it.cases || [])
    .map((c) => `${formatValue(c.value)}, ${formatValue(c.result)}`)
    .join(", ");
  return `=SWITCH(${exprCell}, ${casesStr}${
    it.default != null ? `, ${def}` : ""
  })`;
};

logicalFunctionBuilder.choose = (ctx, formatValue) => {
  const it = ctx.intent;
  const choicesStr = (it.choices || []).map((c) => formatValue(c)).join(", ");

  // 행 전체: 인덱스가 열일 때
  const idxH = it.index_header || it.expression?.header;
  if (idxH) {
    const R = refFromHeaderSpec(ctx, idxH);
    if (!R) return `=ERROR("CHOOSE: 인덱스 열을 찾을 수 없습니다.")`;
    return `=BYROW(${R.range}, LAMBDA(i, CHOOSE(i, ${choicesStr})))`;
  }

  // 단일: 행 선택자 → XLOOKUP으로 index 값 얻기
  if (it.row_selector) {
    const keyCol = refFromHeaderSpec(ctx, it.row_selector.hint);
    const idxCol = refFromHeaderSpec(
      ctx,
      idxH || it.index_num_header || "인덱스",
    );
    if (keyCol && idxCol) {
      const keyVal = formatValue(it.row_selector.value);
      const idx = `XLOOKUP(${keyVal}, ${keyCol.range}, ${idxCol.range})`;
      return `=CHOOSE(${idx}, ${choicesStr})`;
    }
  }

  // 기본 단일
  const indexNum = it.index_num || "A1";
  return `=CHOOSE(${indexNum}, ${choicesStr})`;
};

// === Row/Vector 전용 빌더들 ===
function buildIfRow(ctx, formatValue) {
  const it = ctx.intent;
  const cond = it.condition || {};
  const key = it.row_selector;
  if (!key) return null;

  const keyRefGlobal = refFromHeaderSpec(ctx, {
    header: key.hint,
    sheet: key.sheet,
  });
  const keyVal = formatValue(key.value);

  const left = cond?.target;
  const right = cond?.value;
  const op = _op(cond.operator);
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");

  const L = left ? refFromHeaderSpec(ctx, left) : null;
  const R = right ? refFromHeaderSpec(ctx, right) : null;

  const keyForL =
    (L && refFromHeaderSpec(ctx, { header: key.hint, sheet: L.sheetName })) ||
    keyRefGlobal;
  const keyForR =
    (R && refFromHeaderSpec(ctx, { header: key.hint, sheet: R.sheetName })) ||
    keyRefGlobal;

  if (L && R) {
    if (!keyForL || !keyForR)
      return `=ERROR("행 선택용 키 열을 찾지 못했습니다.")`;
    return _joinMap2(
      L,
      keyForL,
      keyForR,
      R,
      (l, rv) => `IF(${l}${op}${rv}, ${t}, ${f})`,
    );
  }
  if (L && !R) {
    if (!keyForL) return `=ERROR("행 선택용 키 열을 찾지 못했습니다.")`;
    const le = `XLOOKUP(${keyVal}, ${keyForL.range}, ${L.range})`;
    return `=IF(${le}${op}${formatValue(cond.value)}, ${t}, ${f})`;
  }
  if (!L && R) {
    if (!keyForR) return `=ERROR("행 선택용 키 열을 찾지 못했습니다.")`;
    const re = `XLOOKUP(${keyVal}, ${keyForR.range}, ${R.range})`;
    const le = cond.target?.cell || cond.target || "0";
    return `=IF(${le}${op}${re}, ${t}, ${f})`;
  }

  const condStr = _buildConditionString(cond, formatValue, ctx);
  return `=IF(${condStr}, ${t}, ${f})`;
}

function buildIfVectorTwoCols(ctx, formatValue) {
  const it = ctx.intent;
  const cond = it.condition || {};
  const L = refFromHeaderSpec(ctx, cond?.target);
  const R = refFromHeaderSpec(ctx, cond?.value);
  if (!L || !R) return null;

  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");
  const op = _op(cond.operator);
  const join = it.join_by || it.join_on;
  if (join) {
    const leftKey =
      refFromHeaderSpec(
        ctx,
        join?.left_key || { header: join, sheet: L?.sheetName },
      ) || refFromHeaderSpec(ctx, { header: join, sheet: L?.sheetName });
    const rightKey =
      refFromHeaderSpec(
        ctx,
        join?.right_key || { header: join, sheet: R?.sheetName },
      ) || refFromHeaderSpec(ctx, { header: join, sheet: R?.sheetName });

    if (leftKey && rightKey) {
      return _joinMap2(
        L,
        leftKey,
        rightKey,
        R,
        (l, rv) => `IF(${l}${op}${rv}, ${t}, ${f})`,
      );
    }
  }

  const start = Math.max(L.start, R.start);
  const end = Math.min(L.end, R.end);
  const Lrng = `'${L.sheetName}'!${L.col}${start}:${L.col}${end}`;
  const Rrng = `'${R.sheetName}'!${R.col}${start}:${R.col}${end}`;
  return `=MAP(${Lrng}, ${Rrng}, LAMBDA(l, r, IF(l${op}r, ${t}, ${f})))`;
}

function buildIfVectorColConst(ctx, formatValue) {
  const it = ctx.intent;
  const cond = it.condition || {};
  let L = refFromHeaderSpec(ctx, cond?.target);
  if (!L && cond?.target?.header)
    L = refFromHeaderSpec(ctx, cond.target.header);
  if (!L) return null;
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");
  const op = cond.operator || "=";
  const right = formatValue(cond.value);
  return `=BYROW(${L.range}, LAMBDA(l, IF(l${op}${right}, ${t}, ${f})))`;
}

function buildIfThresholds(ctx, formatValue) {
  const it = ctx.intent;
  if (!Array.isArray(it.thresholds) || !it.thresholds.length) return null;
  let ths = [...it.thresholds];
  const opStr = (t) => String(t.operator || ">=");
  const allGE = ths.every((t) => opStr(t).startsWith(">"));
  const allLE = ths.every((t) => opStr(t).startsWith("<"));
  if (allGE) ths.sort((a, b) => Number(b.value) - Number(a.value));
  else if (allLE) ths.sort((a, b) => Number(a.value) - Number(b.value));

  const L =
    refFromHeaderSpec(ctx, it?.condition?.target?.header) ||
    refFromHeaderSpec(ctx, it?.target) ||
    refFromHeaderSpec(ctx, it?.target?.header);
  if (!L) return `=ERROR("기준 열을 찾을 수 없습니다.")`;

  const cases = ths
    .map(
      (th) =>
        `l${_op(th.operator)}${formatValue(th.value)}, ${formatValue(th.label)}`,
    )
    .join(", ");
  const def = it.value_if_false != null ? formatValue(it.value_if_false) : '""';
  return `=BYROW(${L.range}, LAMBDA(l, IFS(${cases}, TRUE, ${def})))`;
}

function buildIfInList(ctx, formatValue) {
  const it = ctx.intent;
  if (!Array.isArray(it.in_values) || !it.in_values.length) return null;
  const L = refFromHeaderSpec(ctx, it?.condition?.target?.header);
  if (!L) return `=ERROR("대상 열을 찾을 수 없습니다.")`;
  const arr = _toArrayConst(it.in_values);
  const test = `ISNUMBER(XMATCH(l, ${arr}, 0))`;
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");
  const negate = String(it.operator || "").toLowerCase() === "not_in";
  return `=BYROW(${L.range}, LAMBDA(l, IF(${
    negate ? `NOT(${test})` : test
  }, ${t}, ${f})))`;
}

function buildIfDateRelative(ctx, formatValue) {
  const it = ctx.intent;
  const cond = it.condition || {};
  const dr = it.date_relative;
  if (!dr) return null;

  const side = String(dr.side || "target").toLowerCase();
  const op = _op(dr.op);
  const rhs = _dateExpr(dr);
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");

  if (side === "target") {
    const L = refFromHeaderSpec(ctx, cond?.target?.header);
    if (!L) return `=ERROR("날짜 열을 찾을 수 없습니다.")`;
    return `=BYROW(${L.range}, LAMBDA(l, IF(l${op}${rhs}, ${t}, ${f})))`;
  } else {
    const R = refFromHeaderSpec(ctx, cond?.value?.header);
    if (!R) return `=ERROR("비교 날짜 열을 찾을 수 없습니다.")`;
    return `=BYROW(${R.range}, LAMBDA(r, IF(${rhs}${op}r, ${t}, ${f})))`;
  }
}

function buildIfBetween(ctx, formatValue) {
  const it = ctx.intent;
  const cond = it.condition || {};
  const bw = it.between || cond.between;
  if (!bw) return null;
  const L = refFromHeaderSpec(ctx, cond?.target?.header);
  if (!L) return `=ERROR("대상 열을 찾을 수 없습니다.")`;
  const minV = formatValue(bw.min);
  const maxV = formatValue(bw.max);
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");
  return `=BYROW(${L.range}, LAMBDA(l, IF(AND(l>=${minV}, l<=${maxV}), ${t}, ${f})))`;
}

function buildIfExists(ctx, formatValue) {
  const it = ctx.intent;
  const ex = it.exists;
  if (!ex) return null;
  const keyRef = _ref(ex.lookup_hint, ctx);
  const inRef = _ref(ex.in_hint, ctx);
  if (!keyRef || !inRef) return `=ERROR("exists용 열을 찾을 수 없습니다.")`;
  const keyVal = formatValue(ex.value);
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");
  const test = `ISNUMBER(XMATCH(${keyVal}, ${inRef.range}))`;
  return `=IF(${test}, ${t}, ${f})`;
}

function buildIfValidate(ctx, formatValue) {
  const it = ctx.intent;
  const v = it.validate;
  if (!v?.headers || !v.expr) return null;
  const keys = Object.keys(v.headers);
  if (!keys.length) return `=ERROR("validate.headers 가 비었습니다.")`;
  const refs = keys.map((k) => refFromHeaderSpec(ctx, v.headers[k]));
  if (refs.some((r) => !r)) return `=ERROR("validate: 열을 찾을 수 없습니다.")`;
  const ranges = refs.map((r) => r.range).join(", ");
  const params = keys.map((k) => k.toLowerCase()).join(", ");
  const expr = v.expr.replace(/\b([A-Za-z]+)\b/g, (m) => m.toLowerCase());
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");
  return `=MAP(${ranges}, LAMBDA(${params}, IF(${expr}, ${t}, ${f})))`;
}

function buildIfCompute(ctx, formatValue) {
  const it = ctx.intent;
  const c = it.compute;
  if (!c?.headers || !c.expr) return null;
  const scopeAll = String(c.scope || it.scope || "all").toLowerCase() === "all";
  const headers = c.headers;
  const keys = Object.keys(headers);
  if (!keys.length) return `=ERROR("compute.headers 가 비었습니다.")`;
  const refs = keys.map((k) => refFromHeaderSpec(ctx, headers[k]));
  if (refs.some((r) => !r)) return `=ERROR("compute: 열을 찾을 수 없습니다.")`;
  let expr = String(c.expr || "");
  if (c.subexpr && typeof c.subexpr === "object") {
    for (const [name, subNode] of Object.entries(c.subexpr)) {
      const sub =
        evalSubIntentToScalar(ctx, formatValue, subNode) || 'ERROR("subexpr")';
      const re1 = new RegExp(String.raw`\$\{${name}\}`, "g");
      const re2 = new RegExp(`\\b${name}\\b`, "g");
      expr = expr.replace(re1, sub).replace(re2, sub);
    }
  }
  expr = expr.replace(/\b([A-Za-z_][A-Za-z0-9_]*)\b/g, (m) => m.toLowerCase());
  expr = expr.replace(/\bAND\b/gi, "*").replace(/\bOR\b/gi, "+");
  expr = expr.replace(
    /\b([a-z_][a-z0-9_]*)\s+IN\s+\{([^}]+)\}/gi,
    (_, v, list) =>
      `ISNUMBER(XMATCH(${v}, ${_toArrayConst(list.split(/\s*,\s*/))}))`,
  );
  const firstParam = keys[0].toLowerCase();
  expr = expr.replace(
    /\{([^}]+)\}/g,
    (_, list) =>
      `ISNUMBER(XMATCH(${firstParam}, ${_toArrayConst(list.split(/\s*,\s*/))}))`,
  );
  const t = formatValue(it.value_if_true ?? "");
  const f = formatValue(it.value_if_false ?? "");
  if (scopeAll) {
    const ranges = refs.map((r) => r.range).join(", ");
    const params = keys.map((k) => k.toLowerCase()).join(", ");
    return `=MAP(${ranges}, LAMBDA(${params}, IF((${expr}), ${t}, ${f})))`;
  }
  const firstCell = refs[0].cell;
  const singleExpr = expr.replace(
    new RegExp(`\\b${firstParam}\\b`, "g"),
    firstCell,
  );
  return `=IF((${singleExpr}), ${t}, ${f})`;
}

// 공개 API
module.exports = logicalFunctionBuilder;
logicalFunctionBuilder.between = between;
logicalFunctionBuilder.inn = inn;
logicalFunctionBuilder.exists = exists;
logicalFunctionBuilder.validate = validate;
logicalFunctionBuilder.compute = compute;
logicalFunctionBuilder._helpers = {
  _ref,
  _refInSheet,
  _scalarFrom,
  _joinMap2,
  _joinMapN,
  _rewriteCondToParams,
  _dateExpr,
};
