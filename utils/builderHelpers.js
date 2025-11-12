/* =========================
 * 시트 / 헤더 메타 공통 헬퍼
 * ========================= */

/** 시트 메타 조회 */
function sheetInfoOf(ctx, sheetName) {
  return (ctx.allSheetsData && ctx.allSheetsData[sheetName]) || null;
}

/** 특정 시트 안에서 헤더 이름으로 열 메타 찾기 */
function resolveHeaderInSheet(ctx, header, sheetName) {
  const info = sheetInfoOf(ctx, sheetName);
  if (!info || !info.metaData) return null;
  const m = info.metaData[header];
  if (!m) return null;
  const col = m.columnLetter;
  const sr = info.startRow;
  const lr = info.lastDataRow;
  return {
    sheetName,
    header,
    columnLetter: col,
    startRow: sr,
    lastDataRow: lr,
    cell: `'${sheetName}'!${col}${sr}`,
    range: `'${sheetName}'!${col}${sr}:${col}${lr}`,
  };
}

/** 모든 시트를 돌면서 헤더 이름에 맞는 열 찾기 */
function resolveHeaderAnySheet(ctx, header) {
  if (!ctx.allSheetsData) return null;
  for (const [sheetName, info] of Object.entries(ctx.allSheetsData)) {
    if (!info.metaData) continue;
    const m = info.metaData[header];
    if (!m) continue;
    const col = m.columnLetter;
    const sr = info.startRow;
    const lr = info.lastDataRow;
    return {
      sheetName,
      header,
      columnLetter: col,
      startRow: sr,
      lastDataRow: lr,
      range: `'${sheetName}'!${col}${sr}:${col}${lr}`,
    };
  }
  return null;
}

/** "시트!헤더" | "헤더" | {header,sheet} → ColumnRange */
function refFromHeaderSpec(ctx, spec) {
  if (!spec) return null;

  // { header, sheet } 형태
  if (typeof spec === "object" && spec.header) {
    if (spec.sheet) return resolveHeaderInSheet(ctx, spec.header, spec.sheet);

    if (ctx.bestReturn?.sheetName) {
      const r = resolveHeaderInSheet(
        ctx,
        spec.header,
        ctx.bestReturn.sheetName
      );
      if (r) return r;
    }
    return resolveHeaderAnySheet(ctx, spec.header);
  }

  // "시트!헤더" 문자열 형태
  const s = String(spec).trim();
  const m = s.match(/^\s*'?([^'!]+)'?\s*!\s*(.+)\s*$/);
  if (m) {
    const sheet = m[1];
    const header = m[2];
    return resolveHeaderInSheet(ctx, header, sheet);
  }

  // 그냥 "헤더"로만 온 경우
  if (ctx.bestReturn?.sheetName) {
    const r = resolveHeaderInSheet(ctx, s, ctx.bestReturn.sheetName);
    if (r) return r;
  }
  return resolveHeaderAnySheet(ctx, s);
}

/** spec → 실제 범위 문자열 (예: "Sheet1!B2:B100") */
function rangeFromSpec(ctx, spec) {
  if (!spec) return null;

  if (typeof spec === "string") {
    // 이미 A1:Z100 같은 범위이면 그대로 반환
    if (/!/.test(spec) && /:/.test(spec)) return spec;
    const r = refFromHeaderSpec(ctx, spec);
    return r ? r.range : null;
  }

  if (typeof spec === "object") {
    // 서브-의도인 경우는 여기서 처리하지 않음
    if (spec.operation) return null;
    const r = refFromHeaderSpec(ctx, spec);
    return r ? r.range : null;
  }

  return null;
}

/* =========================
 * 서브-의도 평가 공통 헬퍼
 * ========================= */

/** 서브-의도 → 스칼라 수식 (예: 내부 XLOOKUP 결과 등) */
function evalSubIntentToScalar(ctx, formatValue, node) {
  if (!node || typeof node !== "object" || !node.operation) return null;
  const fb = ctx.formulaBuilder;
  if (fb && typeof fb[node.operation] === "function") {
    const res = fb[node.operation].call(
      fb,
      { ...ctx, intent: node },
      formatValue,
      fb._buildConditionPairs
    );
    if (typeof res === "string" && res.startsWith("=")) {
      return res.slice(1);
    }
    return res;
  }
  return null;
}

/** 서브-의도 → 범위/배열 수식 (지금은 스칼라와 동일하게 처리) */
function evalSubIntentToRange(ctx, formatValue, node) {
  return evalSubIntentToScalar(ctx, formatValue, node);
}

/* =========================
 * 기타 유틸
 * ========================= */

const asArray = (v) => (Array.isArray(v) ? v : [v]);

module.exports = {
  sheetInfoOf,
  resolveHeaderInSheet,
  resolveHeaderAnySheet,
  refFromHeaderSpec,
  rangeFromSpec,
  evalSubIntentToScalar,
  evalSubIntentToRange,
  asArray,
};
