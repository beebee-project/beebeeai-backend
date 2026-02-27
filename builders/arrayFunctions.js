const formulaUtils = require("../utils/formulaUtils");
const { rangeFromSpec } = require("../utils/builderHelpers");

// 모든 인자를 "1열 벡터"로 정규화 (공백 무시)
function _broadcastToColumn(exprOrRange, ctx) {
  const e = rangeFromSpec(ctx, exprOrRange) || exprOrRange;
  return `TOCOL(${e}, 1)`; // 1 => 공백 무시
}

// N개의 인자를 같은 길이의 1열 벡터로 정렬
// - 전략: 먼저 모두 TOCOL(…,1)로 세로정렬 → 각 ROWS를 구해 최솟값 L
// - 각 벡터를 TAKE(vec, L)로 잘라 길이 일치
function _alignManyToColumn(exprList, ctx) {
  const cols = exprList.map((e) => _broadcastToColumn(e, ctx));
  const lens = cols.map((c) => `ROWS(${c})`);
  const minLen = lens.length === 1 ? lens[0] : `MIN(${lens.join(", ")})`;
  const taken = cols.map((c) => `TAKE(${c}, ${minLen})`);
  return { vectors: taken, rowsExpr: minLen };
}

function _alignTo(targetRange, valueExpr) {
  const tr = String(targetRange);
  const vx = String(valueExpr);
  const asCol = (expr) => `TOCOL(${expr}, 1)`;
  return {
    asColumn: asCol(vx),
    indexAt: (rowBaseA1) => `INDEX(${asCol(vx)}, ROW(r)-ROW(${rowBaseA1})+1)`,
  };
}

function _normalizeOp(op) {
  const m = { "==": "=", "!=": "<>", "≠": "<>", "≤": "<=", "≥": ">=" };
  return m[op] || op || "=";
}

function _isISODate(v) {
  return /^\d{4}-\d{1,2}-\d{1,2}$/.test(String(v || "").trim());
}

function _isNumericLiteral(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).replace(/,/g, "").trim();
  return /^-?\d+(\.\d+)?$/.test(s);
}

function _q(s) {
  return `"${String(s ?? "").replace(/"/g, '""')}"`;
}

function _dateVal(iso) {
  return `DATEVALUE(${_q(iso)})`;
}

// ---------- Step2: normalize/coerce helpers ----------
function _trimText(expr) {
  // 숫자/빈값/날짜 모두 문자열로 안전 변환 후 TRIM
  return `TRIM(${expr}&"")`;
}

function _normText(expr, cs) {
  // case-insensitive면 LOWER 적용
  const t = _trimText(expr);
  return cs ? t : `LOWER(${t})`;
}

function _coerceNumber(expr) {
  // "00123" 같은 문자열 숫자를 숫자로
  // 실패하면 원본 유지(에러 방지)
  return `IFERROR(VALUE(${_trimText(expr)}), ${expr})`;
}

function _coerceDate(expr) {
  // 텍스트 날짜면 DATEVALUE로, 이미 날짜/시리얼이면 원본 유지
  return `IFERROR(DATEVALUE(${_trimText(expr)}), ${expr})`;
}

function _isSheetsContext(ctx) {
  // ctx / intent에 "Sheets" 힌트가 있으면 Sheets로 간주
  const it = ctx?.intent || {};
  const hint = String(
    it.engine ||
      it.platform ||
      it.target_app ||
      it.conversionType ||
      ctx?.conversionType ||
      "",
  ).toLowerCase();
  return hint.includes("sheet");
}

// ---------- 텍스트 연산 보조(케이스 민감도 토글) ----------
function _hasCaseInsensInlineFlag(pattern) {
  const p = String(pattern || "");
  return (
    /^\(\?[a-z]*i[a-z]*\)/i.test(p) ||
    /^\(\?[a-z]*i[a-z]*:/.test(p) ||
    /\(\?[a-z]*i[a-z]*:/.test(p)
  );
}

function _stripCaseInsensInlineFlag(pattern) {
  let p = String(pattern ?? "");
  p = p.replace(/\(\?([a-z]*?)i([a-z]*?):/gi, (_m, a, b) => {
    const flags = (String(a) + String(b)).replace(/i/gi, "");
    return flags ? `(?${flags}:` : `(?:`;
  });
  p = p.replace(/\(\?([a-z]*?)i([a-z]*?)\)/gi, (_m, a, b) => {
    const flags = (String(a) + String(b)).replace(/i/gi, "");
    return flags ? `(?${flags})` : ``;
  });
  return p;
}

function _containsExpr(colA1, needle, cs) {
  // Step2: 공백/타입 혼합 안정화 (TRIM + &"")
  const colN = _normText(colA1, cs);
  const ndl = cs ? String(needle ?? "") : String(needle ?? "").toLowerCase();
  return cs
    ? `ISNUMBER(FIND(${_q(ndl)}, ${colN}))`
    : `ISNUMBER(SEARCH(${_q(ndl)}, ${colN}))`;
}

function _startsWithExpr(colA1, needle, cs) {
  const colN = _normText(colA1, cs);
  const ndl = cs ? String(needle ?? "") : String(needle ?? "").toLowerCase();
  const q = _q(ndl);
  return cs
    ? `EXACT(LEFT(${colN}, LEN(${q})), ${q})`
    : `LEFT(${colN}, LEN(${q}))=${q}`;
}

function _endsWithExpr(colA1, needle, cs) {
  const colN = _normText(colA1, cs);
  const ndl = cs ? String(needle ?? "") : String(needle ?? "").toLowerCase();
  const q = _q(ndl);
  return cs
    ? `EXACT(RIGHT(${colN}, LEN(${q})), ${q})`
    : `RIGHT(${colN}, LEN(${q}))=${q}`;
}

function _regexMatchExpr(colA1, pattern, cs, strict) {
  let pat = String(pattern ?? "");
  if (cs && strict) pat = _stripCaseInsensInlineFlag(pat);
  if (!cs && !_hasCaseInsensInlineFlag(pat)) pat = `(?i)${pat}`;
  return `REGEXMATCH(${colA1}, ${_q(pat)})`;
}

function _normRange(rg) {
  return `UPPER(TRIM(${rg}&""))`;
}

// --- Build concatenated normalized key vector from multiple ranges ---
function _concatKeyVector(ranges, sep = "|") {
  if (!Array.isArray(ranges) || !ranges.length) return null;
  // Step3-2: 복합키 안정화
  // - BYROW의 r는 "현재 행" 배열이므로 ROW(r)을 그대로 INDEX에 넣으면(절대행번호) 범위 시작행/형태에 따라 흔들릴 수 있음
  // - i = 현재행의 상대 인덱스(1부터)로 계산해서 모든 range에 동일하게 적용
  // - 각 키 파트는 TRIM + UPPER + 문자열 강제(&"")로 타입/공백 혼합을 안정화
  const base = ranges[0];
  const parts = ranges
    .map((rg) => `UPPER(TRIM(INDEX(${rg}, i)&""))`)
    .join(`&${_q(sep)}&`);

  // i: base 범위의 첫 셀을 기준으로 현재 행의 상대 인덱스
  // ROW(r)은 현재 행의 절대 행번호 → base의 시작행을 빼서 1부터 만드는 방식
  return `BYROW(${base}, LAMBDA(r, LET(i, ROW(r)-ROW(INDEX(${base}, 1, 1))+1, ${parts})))`;
}

// 공통: 기본 대상 범위 해석
function _resolveRangeOrError(it, ctx) {
  const r =
    rangeFromSpec(ctx, it.range || it.target_header || it.header_hint) ||
    (ctx.bestReturn
      ? `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}:${ctx.bestReturn.columnLetter}${ctx.bestReturn.lastDataRow}`
      : null);
  return r;
}

// ================================
// FILTER / TOPN 공용: 조건 마스크 생성
// ================================
function _buildMaskExprFromIntent(ctx, sheetInfo, sheetName) {
  const intent = ctx.intent || {};
  const isSheets = _isSheetsContext(ctx);

  // Step2: regex가 Excel에서 불가하므로, 필요 시 조기에 ERROR 반환
  let earlyError = null;

  // 1) 기본 conditions
  const condNodes = Array.isArray(intent.conditions) ? intent.conditions : [];
  const masks = condNodes
    .map((cond) => {
      const termSet = formulaUtils.expandTermsFromText(cond.hint);
      const bestCol = formulaUtils.bestHeaderInSheet(
        sheetInfo,
        sheetName,
        termSet,
        "lookup",
      );
      if (!bestCol?.col) return null;
      // Step1 연계: 열 후보가 모호하면 오답 대신 중단
      if (bestCol.isAmbiguous) {
        earlyError = `=ERROR("조건 열이 모호합니다: '${bestCol.header}' 또는 '${
          bestCol.runnerUp?.header || "다른 후보"
        }' 중 선택이 필요합니다.")`;
        return null;
      }
      const colA1 = `'${sheetName}'!${bestCol.col.columnLetter}${sheetInfo.startRow}:${bestCol.col.columnLetter}${sheetInfo.lastDataRow}`;

      const rawOp = String(cond.operator || "=").toLowerCase();
      const op = _normalizeOp(rawOp);
      const rawVal = cond.value;

      // 날짜/숫자 비교는 열 값을 안전 coercion
      if (_isISODate(rawVal))
        return `${_coerceDate(colA1)}${op}${_dateVal(rawVal)}`;
      if (_isNumericLiteral(rawVal))
        return `${_coerceNumber(colA1)}${op}${String(rawVal).replace(/,/g, "")}`;

      const cs = (cond.case_sensitive ?? intent.case_sensitive) === true;
      if (["contains", "포함"].includes(rawOp))
        return _containsExpr(colA1, rawVal, cs);
      if (["startswith", "startsWith"].includes(rawOp))
        return _startsWithExpr(colA1, rawVal, cs);
      if (["endswith", "endsWith"].includes(rawOp))
        return _endsWithExpr(colA1, rawVal, cs);
      if (
        ["in", "any_of"].includes(rawOp) &&
        Array.isArray(cond.values) &&
        cond.values.length
      ) {
        // IN은 MATCH 기반(텍스트는 TRIM/LOWER 정규화)
        const colN = _normText(colA1, cs);
        const values = cond.values.map((v) => {
          if (_isNumericLiteral(v)) return String(v).replace(/,/g, "");
          const s = cs
            ? String(v ?? "").trim()
            : String(v ?? "")
                .trim()
                .toLowerCase();
          return _q(s);
        });
        return `ISNUMBER(MATCH(${colN}, {${values.join(",")}}, 0))`;
      }
      if (rawOp === "between" && cond.min != null && cond.max != null) {
        const L = _q(cond.min);
        const R = _q(cond.max);
        const left = _trimText(colA1);
        return `(${left}>=${L})*(${left}<=${R})`;
      }
      if (["matches", "regex"].includes(rawOp)) {
        if (!isSheets) {
          earlyError = `=ERROR("정규식 조건은 Google Sheets에서만 지원됩니다.")`;
          return null;
        }
        const strict =
          (cond.strip_inline_flags ?? intent.strip_inline_flags) === true;
        return _regexMatchExpr(colA1, rawVal, cs, strict);
      }
      return `${_trimText(colA1)}${op}${_q(rawVal)}`;
    })
    .filter(Boolean);

  // 2) groups 조건 (조건 그룹)
  const groupMasks = (
    Array.isArray(intent.condition_groups)
      ? intent.condition_groups
      : Array.isArray(intent.groups)
        ? intent.groups
        : []
  )
    .map((g) => {
      const items = Array.isArray(g.conditions) ? g.conditions : [];
      const masksInGroup = items
        .map((cond) => {
          const termSet = formulaUtils.expandTermsFromText(cond.hint);
          const bestCol = formulaUtils.bestHeaderInSheet(
            sheetInfo,
            sheetName,
            termSet,
            "lookup",
          );
          if (!bestCol?.col) return null;
          if (bestCol.isAmbiguous) {
            earlyError = `=ERROR("조건 열이 모호합니다: '${bestCol.header}' 또는 '${
              bestCol.runnerUp?.header || "다른 후보"
            }' 중 선택이 필요합니다.")`;
            return null;
          }
          const colA1 = `'${sheetName}'!${bestCol.col.columnLetter}${sheetInfo.startRow}:${bestCol.col.columnLetter}${sheetInfo.lastDataRow}`;
          const rawOp = String(cond.operator || "=").toLowerCase();
          const op = _normalizeOp(rawOp);
          const rawVal = cond.value;
          const cs = (cond.case_sensitive ?? intent.case_sensitive) === true;

          if (_isISODate(rawVal))
            return `${_coerceDate(colA1)}${op}${_dateVal(rawVal)}`;
          if (_isNumericLiteral(rawVal))
            return `${_coerceNumber(colA1)}${op}${String(rawVal).replace(/,/g, "")}`;

          if (["contains", "포함"].includes(rawOp))
            return _containsExpr(colA1, rawVal, cs);
          if (["startswith", "startsWith"].includes(rawOp))
            return _startsWithExpr(colA1, rawVal, cs);
          if (["endswith", "endsWith"].includes(rawOp))
            return _endsWithExpr(colA1, rawVal, cs);
          if (rawOp === "between" && cond.min != null && cond.max != null) {
            const L = _q(cond.min);
            const R = _q(cond.max);
            const left = _trimText(colA1);
            return `(${left}>=${L})*(${left}<=${R})`;
          }
          if (["matches", "regex"].includes(rawOp)) {
            if (!isSheets) {
              earlyError = `=ERROR("정규식 조건은 Google Sheets에서만 지원됩니다.")`;
              return null;
            }
            const strict =
              (cond.strip_inline_flags ?? intent.strip_inline_flags) === true;
            return _regexMatchExpr(colA1, rawVal, cs, strict);
          }
          return `${_trimText(colA1)}${op}${_q(rawVal)}`;
        })
        .filter(Boolean);
      const useOR = String(g.logical || "AND").toUpperCase() === "OR";
      return masksInGroup.length
        ? `(${masksInGroup.join(useOR ? " + " : " * ")})`
        : null;
    })
    .filter(Boolean);

  // 3) 기존 conditions + 그룹 마스크 결합
  const isOR =
    String(
      intent.logical || intent.conditions_logical || "AND",
    ).toUpperCase() === "OR";
  const baseMask = masks.length ? `(${masks.join(isOR ? " + " : " * ")})` : "";
  const groupsLogicalOR =
    String(intent.groups_logical || "AND").toUpperCase() === "OR";
  const combinedMask = [baseMask, ...groupMasks]
    .filter(Boolean)
    .join(groupsLogicalOR ? " + " : " * ");

  // 4) 빈값(공백) 제외 옵션
  const blanks = Array.isArray(intent.exclude_blank_in)
    ? intent.exclude_blank_in
    : [];
  const blankMasks = blanks
    .map((h) => {
      const termSet = formulaUtils.expandTermsFromText(h);
      const colInfo = formulaUtils.bestHeaderInSheet(
        sheetInfo,
        sheetName,
        termSet,
        "lookup",
      );
      if (!colInfo?.col) return null;
      const a1 = `'${sheetName}'!${colInfo.col.columnLetter}${sheetInfo.startRow}:${colInfo.col.columnLetter}${sheetInfo.lastDataRow}`;
      return `LEN(TRIM(${a1}&""))>0`;
    })
    .filter(Boolean);
  const blankMaskExpr = blankMasks.length
    ? ` * (${blankMasks.join(" * ")})`
    : "";

  const finalMask = (combinedMask || "TRUE") + blankMaskExpr; // 조건 없을 때도 TRUE에서 시작
  return { maskExpr: finalMask, earlyError };
}

const arrayFunctionBuilder = {
  // ---------------------- FILTER ----------------------
  filter: function (ctx) {
    const { bestReturn, intent, allSheetsData } = ctx;
    if (!bestReturn) return `=ERROR("반환할 열을 찾을 수 없습니다.")`;

    const sheetName = bestReturn.sheetName;
    const sheetInfo = allSheetsData[sheetName];
    if (!sheetInfo) return `=ERROR("시트 정보를 찾을 수 없습니다.")`;

    // 0) 시트 전체 폭 (FILTER→CHOOSECOLS를 위한 기본 fullRange)
    const metaEntries = Object.entries(sheetInfo.metaData || {});
    if (!metaEntries.length)
      return `=ERROR("시트의 열 정보를 찾을 수 없습니다.")`;
    metaEntries.sort((a, b) => {
      const ai = formulaUtils.columnLetterToIndex(a[1].columnLetter);
      const bi = formulaUtils.columnLetterToIndex(b[1].columnLetter);
      return ai - bi;
    });
    const firstCol = metaEntries[0][1].columnLetter;
    const lastCol = metaEntries[metaEntries.length - 1][1].columnLetter;
    const fullRange = `'${sheetName}'!${firstCol}${sheetInfo.startRow}:${lastCol}${sheetInfo.lastDataRow}`;
    const returnRangeSingle = `'${sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}:${bestReturn.columnLetter}${bestReturn.lastDataRow}`;

    // 1) 조건 마스크 (AND/*, OR/+)
    // Step2: regex가 Excel에서 불가하므로, 필요 시 조기에 ERROR 반환
    const { maskExpr, earlyError } = _buildMaskExprFromIntent(
      ctx,
      sheetInfo,
      sheetName,
    );
    if (earlyError) return earlyError;

    // 2) 조인(inner/left) 및 오른쪽 열 픽업
    const joins = Array.isArray(intent.joins) ? intent.joins : [];
    const rightPickExprs = [];
    for (const j of joins) {
      if (!j?.sheet || !Array.isArray(j.on) || !j.on.length) continue;
      const leftRanges = [];
      const rightRanges = [];
      for (const pair of j.on) {
        const lTerm = formulaUtils.expandTermsFromText(pair.left);
        const lCol = formulaUtils.bestHeaderInSheet(
          sheetInfo,
          sheetName,
          lTerm,
          "lookup",
        );
        if (!lCol?.col) continue;
        leftRanges.push(
          `'${sheetName}'!${lCol.col.columnLetter}${sheetInfo.startRow}:${lCol.col.columnLetter}${sheetInfo.lastDataRow}`,
        );
        const rightSheet = allSheetsData[j.sheet];
        if (!rightSheet) continue;
        const rTerm = formulaUtils.expandTermsFromText(pair.right);
        const rCol = formulaUtils.bestHeaderInSheet(
          rightSheet,
          j.sheet,
          rTerm,
          "lookup",
        );
        if (!rCol?.col) continue;
        rightRanges.push(
          `'${j.sheet}'!${rCol.col.columnLetter}${rightSheet.startRow}:${rCol.col.columnLetter}${rightSheet.lastDataRow}`,
        );
      }
      if (!leftRanges.length || !rightRanges.length) continue;

      // Step3: JOIN 존재 마스크를 "행 단위(MAP)"로 고정 (배열 MATCH 흔들림 방지)
      const joinMasks = leftRanges.map((lr, i) => {
        const L = _normRange(lr);
        const R = _normRange(rightRanges[i]);
        return `MAP(${L}, LAMBDA(k, ISNUMBER(MATCH(k, ${R}, 0))))`;
      });
      const joinMaskExpr = joinMasks.join(" * ");
      const joinType = String(j.type || "inner").toLowerCase();
      if (joinType === "inner") maskExpr = `${maskExpr} * (${joinMaskExpr})`;

      const picks = Array.isArray(j.pick_from_right) ? j.pick_from_right : [];
      const notFoundFill = j.if_not_found != null ? _q(j.if_not_found) : '""';
      for (const hdr of picks) {
        const rightSheet = allSheetsData[j.sheet];
        if (!rightSheet) continue;
        const term = formulaUtils.expandTermsFromText(hdr);
        const col = formulaUtils.bestHeaderInSheet(
          rightSheet,
          j.sheet,
          term,
          "lookup",
        );
        if (!col?.col) continue;
        const retRange = `'${j.sheet}'!${col.col.columnLetter}${rightSheet.startRow}:${col.col.columnLetter}${rightSheet.lastDataRow}`;
        if (leftRanges.length === 1 && rightRanges.length === 1) {
          const L = _normRange(leftRanges[0]);
          const R = _normRange(rightRanges[0]);
          rightPickExprs.push(
            // Step3: 픽업도 "행 단위(MAP)"로 고정
            `MAP(${L}, LAMBDA(k, XLOOKUP(k, ${R}, ${retRange}, ${notFoundFill}, 0)))`,
          );
        } else {
          const leftKeyVec = _concatKeyVector(leftRanges);
          const rightKeyVec = _concatKeyVector(rightRanges);
          if (leftKeyVec && rightKeyVec) {
            rightPickExprs.push(
              `MAP(${leftKeyVec}, LAMBDA(k, XLOOKUP(k, ${rightKeyVec}, ${retRange}, ${notFoundFill}, 0)))`,
            );
          }
        }
      }
    }

    // --- 반환열 제어(선택)
    const headerOpts =
      intent.return_headers || intent.select_headers || intent.return_cols;
    if (!headerOpts || !Array.isArray(headerOpts) || headerOpts.length === 0) {
      return `=FILTER(${returnRangeSingle}, ${maskExpr})`;
    }

    const filteredAll = `FILTER(${fullRange}, ${maskExpr})`;

    const nameToIndex = new Map(metaEntries.map(([h, m], i) => [h, i + 1]));
    const wantedIdx = [];
    for (const hSpec of headerOpts) {
      let hName = null;
      if (typeof hSpec === "string") {
        const mm = hSpec.match(/^\s*'?([^'!]+)'?\s*!\s*(.+)\s*$/);
        hName = mm ? mm[2].trim() : hSpec.trim();
      } else if (hSpec && typeof hSpec === "object" && hSpec.header) {
        hName = String(hSpec.header).trim();
      }
      const idx = nameToIndex.get(hName);
      if (idx) wantedIdx.push(idx);
    }

    const pickedLeft = wantedIdx.length
      ? `CHOOSECOLS(${filteredAll}, ${wantedIdx.join(", ")})`
      : filteredAll;

    const selectedIndexMap = new Map();
    wantedIdx.forEach((idx, i) => {
      const name = String(headerOpts[i]?.header || headerOpts[i] || "").trim();
      if (name) selectedIndexMap.set(name, i + 1);
    });

    if (rightPickExprs.length) {
      const joined =
        rightPickExprs.length === 1
          ? rightPickExprs[0]
          : `HSTACK(${rightPickExprs.join(", ")})`;
      return pipeSortIfRequested(
        ctx,
        intent,
        `HSTACK(${pickedLeft}, ${joined})`,
        selectedIndexMap,
      );
    }
    return pipeSortIfRequested(ctx, intent, pickedLeft, selectedIndexMap);
  },

  // ---------------------- TOP N (Sheets) ----------------------
  // 목적:
  //  - "마케팅 부서 연봉 Top3의 직원ID, 이름, 연봉" 같은 요청을 안정적으로 처리
  //
  // intent 기대(없어도 최대한 유추):
  //  - top_n: number (default 3)
  //  - order_by_header or header_hint: 정렬 기준 열 (예: "연봉")
  //  - return_headers: ["직원 ID","이름","연봉"] (없으면 return_hint를 쉼표로 split 시도)
  //  - conditions: [{hint:"부서", operator:"=", value:"마케팅"} ...] (filter()와 동일 포맷)
  topn: function (ctx) {
    const { intent, allSheetsData } = ctx;
    const isSheets = _isSheetsContext(ctx);
    if (!isSheets)
      return `=ERROR("TopN은 현재 Google Sheets에서만 지원됩니다.")`;

    // 기본 시트 선택: bestReturn이 있으면 그 시트, 없으면 첫 시트
    const sheetName =
      ctx.bestReturn?.sheetName ||
      (allSheetsData ? Object.keys(allSheetsData)[0] : null);
    if (!sheetName) return `=ERROR("시트 정보를 찾을 수 없습니다.")`;
    const sheetInfo = allSheetsData[sheetName];
    if (!sheetInfo) return `=ERROR("시트 정보를 찾을 수 없습니다.")`;

    // fullRange(폭 계산) - filter()와 동일
    const metaEntries = Object.entries(sheetInfo.metaData || {});
    if (!metaEntries.length)
      return `=ERROR("시트의 열 정보를 찾을 수 없습니다.")`;
    metaEntries.sort((a, b) => {
      const ai = formulaUtils.columnLetterToIndex(a[1].columnLetter);
      const bi = formulaUtils.columnLetterToIndex(b[1].columnLetter);
      return ai - bi;
    });
    const firstCol = metaEntries[0][1].columnLetter;
    const lastCol = metaEntries[metaEntries.length - 1][1].columnLetter;
    const fullRange = `'${sheetName}'!${firstCol}${sheetInfo.startRow}:${lastCol}${sheetInfo.lastDataRow}`;

    const N = Math.max(
      1,
      parseInt(intent.top_n || intent.n || intent.limit || 3, 10) || 3,
    );

    // 반환 열들 결정
    let returnHeaders = Array.isArray(intent.return_headers)
      ? intent.return_headers.slice()
      : [];
    if (!returnHeaders.length && typeof intent.return_hint === "string") {
      // "직원 ID, 이름, 연봉" 같은 힌트 처리
      returnHeaders = intent.return_hint
        .split(/[,/|]/)
        .map((s) => s.trim())
        .filter(Boolean);
    }
    if (!returnHeaders.length && ctx.bestReturn?.header) {
      returnHeaders = [ctx.bestReturn.header];
    }

    // 정렬 기준 열
    const orderHeader =
      intent.order_by_header ||
      intent.order_header ||
      (typeof intent.order_by === "string" ? intent.order_by : null) ||
      intent.header_hint;
    if (!orderHeader) return `=ERROR("TopN: 정렬 기준 열을 찾을 수 없습니다.")`;

    // 헤더 → 열 범위(A1) 변환(모호성 차단)
    function _colRangeByHeader(h, role = "return") {
      const termSet = formulaUtils.expandTermsFromText(h);
      const bestCol = formulaUtils.bestHeaderInSheet(
        sheetInfo,
        sheetName,
        termSet,
        role,
      );
      if (!bestCol?.col) return null;
      if (bestCol.isAmbiguous) {
        return {
          error: `=ERROR("열이 모호합니다: '${bestCol.header}' 또는 '${bestCol.runnerUp?.header || "다른 후보"}' 중 선택이 필요합니다.")`,
        };
      }
      return {
        range: `'${sheetName}'!${bestCol.col.columnLetter}${sheetInfo.startRow}:${bestCol.col.columnLetter}${sheetInfo.lastDataRow}`,
      };
    }

    const orderRef = _colRangeByHeader(orderHeader, "lookup");
    if (!orderRef) return `=ERROR("TopN: 정렬 기준 열을 찾을 수 없습니다.")`;
    if (orderRef.error) return orderRef.error;

    const retRefs = [];
    for (const h of returnHeaders) {
      const rr = _colRangeByHeader(h, "return");
      if (!rr) return `=ERROR("TopN: 반환 열을 찾을 수 없습니다: ${_q(h)}")`;
      if (rr.error) return rr.error;
      retRefs.push(rr.range);
    }

    // 조건 마스크 생성: filter()의 핵심 로직 일부만 재사용(AND/OR 포함)
    // (단, join 등 고급 기능은 topn에서는 제외)
    let earlyError = null;
    const condNodes = Array.isArray(intent.conditions) ? intent.conditions : [];
    const masks = condNodes
      .map((cond) => {
        const termSet = formulaUtils.expandTermsFromText(cond.hint);
        const bestCol = formulaUtils.bestHeaderInSheet(
          sheetInfo,
          sheetName,
          termSet,
          "lookup",
        );
        if (!bestCol?.col) return null;
        if (bestCol.isAmbiguous) {
          earlyError = `=ERROR("조건 열이 모호합니다: '${bestCol.header}' 또는 '${bestCol.runnerUp?.header || "다른 후보"}' 중 선택이 필요합니다.")`;
          return null;
        }
        const colA1 = `'${sheetName}'!${bestCol.col.columnLetter}${sheetInfo.startRow}:${bestCol.col.columnLetter}${sheetInfo.lastDataRow}`;
        const rawOp = String(cond.operator || "=").toLowerCase();
        const op = _normalizeOp(rawOp);
        const rawVal = cond.value;
        if (_isISODate(rawVal))
          return `${_coerceDate(colA1)}${op}${_dateVal(rawVal)}`;
        if (_isNumericLiteral(rawVal))
          return `${_coerceNumber(colA1)}${op}${String(rawVal).replace(/,/g, "")}`;
        const cs = (cond.case_sensitive ?? intent.case_sensitive) === true;
        if (["contains", "포함"].includes(rawOp))
          return _containsExpr(colA1, rawVal, cs);
        if (["startswith", "startswith"].includes(rawOp))
          return _startsWithExpr(colA1, rawVal, cs);
        if (["endswith", "endswith"].includes(rawOp))
          return _endsWithExpr(colA1, rawVal, cs);
        return `${_trimText(colA1)}${op}${_q(rawVal)}`;
      })
      .filter(Boolean);

    if (earlyError) return earlyError;
    const isOR =
      String(
        intent.logical || intent.conditions_logical || "AND",
      ).toUpperCase() === "OR";
    const maskExpr = masks.length
      ? `(${masks.join(isOR ? " + " : " * ")})`
      : "TRUE";

    // 데이터: {returnCols..., orderCol} 를 필터 → orderCol desc 정렬 → 상위 N 추출 → returnCols만
    const dataBlock = `{${retRefs.join(",")},${orderRef.range}}`;
    const sortIndex = retRefs.length + 1;
    const chooseCols = retRefs.map((_r, i) => i + 1).join(",");

    return `=LET(_d, FILTER(${dataBlock}, ${maskExpr}), _s, SORT(_d, ${sortIndex}, FALSE), TAKE(CHOOSECOLS(_s, ${chooseCols}), ${N}))`;
  },

  // ---------------------- UNIQUE ----------------------
  unique: (ctx) => {
    const { bestReturn } = ctx;
    if (!bestReturn) return `=ERROR("범위를 찾을 수 없습니다.")`;
    const sheetName = bestReturn.sheetName;
    const targetRange = `'${sheetName}'!${bestReturn.columnLetter}${bestReturn.startRow}:${bestReturn.columnLetter}${bestReturn.lastDataRow}`;
    const it = ctx.intent || {};
    if (Array.isArray(it.unique_by) && it.unique_by.length) {
      const info = ctx.allSheetsData[sheetName];
      const meta = Object.entries(info.metaData || {}).sort(
        (a, b) =>
          formulaUtils.columnLetterToIndex(a[1].columnLetter) -
          formulaUtils.columnLetterToIndex(b[1].columnLetter),
      );
      const nameToIndex = new Map(meta.map(([h, m], i) => [h, i + 1]));
      const idxs = it.unique_by.map((h) => nameToIndex.get(h)).filter(Boolean);
      const full = `'${sheetName}'!${meta[0][1].columnLetter}${info.startRow}:${
        meta[meta.length - 1][1].columnLetter
      }${info.lastDataRow}`;
      return `=UNIQUE(CHOOSECOLS(${full}, ${idxs.join(", ")}))`;
    }
    return `=UNIQUE(${targetRange})`;
  },

  // ---------------------- SORT ----------------------
  sort: (ctx) => {
    const { bestReturn, allSheetsData } = ctx;
    if (!bestReturn) return `=ERROR("정렬 기준 열을 찾을 수 없습니다.")`;

    const sheetName = bestReturn.sheetName;
    const sheetInfo = allSheetsData[sheetName];
    if (!sheetInfo) return `=ERROR("시트 정보를 찾을 수 없습니다.")`;

    const metaEntries = Object.entries(sheetInfo.metaData || {}).sort(
      (a, b) =>
        formulaUtils.columnLetterToIndex(a[1].columnLetter) -
        formulaUtils.columnLetterToIndex(b[1].columnLetter),
    );
    if (!metaEntries.length)
      return `=ERROR("시트의 열 정보를 찾을 수 없습니다.")`;

    const firstCol = metaEntries[0][1].columnLetter;
    const lastCol = metaEntries[metaEntries.length - 1][1].columnLetter;
    const fullRange = `${firstCol}${sheetInfo.startRow}:${lastCol}${sheetInfo.lastDataRow}`;

    const sortIndex =
      metaEntries.findIndex(
        ([_h, m]) => m.columnLetter === bestReturn.columnLetter,
      ) + 1;
    if (sortIndex === 0)
      return `=ERROR("정렬 기준 열의 위치를 찾을 수 없습니다.")`;

    const it = ctx.intent || {};
    const order =
      String(it.sort_order || "desc").toLowerCase() === "asc" ? 1 : -1;
    return `=SORT('${sheetName}'!${fullRange}, ${sortIndex}, ${order})`;
  },

  // ---------------------- SORTBY ----------------------
  sortby: function (ctx) {
    const { bestReturn, bestLookup } = ctx;
    if (!bestReturn || !bestLookup)
      return `=ERROR("필요한 열을 모두 찾을 수 없습니다.")`;
    if (bestReturn.sheetName !== bestLookup.sheetName)
      return `=ERROR("정렬할 열과 기준 열은 같은 시트에 있어야 합니다.")`;

    const sheetName = bestReturn.sheetName;
    const returnRange = `${bestReturn.columnLetter}${bestReturn.startRow}:${bestReturn.columnLetter}${bestReturn.lastDataRow}`;
    const criteriaRange = `${bestLookup.columnLetter}${bestLookup.startRow}:${
      bestLookup.lastDataRow
        ? bestLookup.columnLetter + bestLookup.lastDataRow
        : bestLookup.columnLetter + bestLookup.lastDataRow
    }`; // safety

    const it = ctx.intent || {};
    const multi = Array.isArray(it.sort_by) ? it.sort_by : null;
    if (multi && multi.length) {
      const parts = [];
      for (const k of multi) {
        const ord =
          String(k.order || it.sort_order || "desc").toLowerCase() === "asc"
            ? 1
            : -1;
        if (k.range) {
          parts.push(k.range, ord);
        } else {
          const term = formulaUtils.expandTermsFromText(k.header || k);
          const col = formulaUtils.bestHeaderInSheet(
            ctx.allSheetsData[sheetName],
            sheetName,
            term,
            "lookup",
          );
          if (!col?.col) continue;
          const rng = `'${sheetName}'!${col.col.columnLetter}${bestLookup.startRow}:${col.col.columnLetter}${bestLookup.lastDataRow}`;
          parts.push(rng, ord);
        }
      }
      if (parts.length)
        return `=SORTBY('${sheetName}'!${returnRange}, ${parts.join(", ")})`;
    }
    const order =
      String(it.sort_order || "desc").toLowerCase() === "asc" ? 1 : -1;
    return `=SORTBY('${sheetName}'!${returnRange}, '${sheetName}'!${criteriaRange}, ${order})`;
  },

  // ---------------------- 고급 동적배열 ----------------------
  byrow: function (ctx) {
    const it = ctx.intent || {};
    const range = _resolveRangeOrError(it, ctx) || "A1:C10";
    const param = (it.lambda_params && it.lambda_params[0]) || "row";
    const agg = String(it.aggregate || "").toLowerCase();
    const body =
      it.lambda_body ||
      (agg === "sum"
        ? `SUM(${param})`
        : agg === "max"
          ? `MAX(${param})`
          : agg === "min"
            ? `MIN(${param})`
            : agg === "avg"
              ? `AVERAGE(${param})`
              : `SUM(${param})`);
    return `=BYROW(${range}, LAMBDA(${param}, ${body}))`;
  },
  bycol: function (ctx) {
    const it = ctx.intent || {};
    const range = _resolveRangeOrError(it, ctx) || "A1:C10";
    const param = (it.lambda_params && it.lambda_params[0]) || "col";
    const agg = String(it.aggregate || "").toLowerCase();
    const body =
      it.lambda_body ||
      (agg === "sum"
        ? `SUM(${param})`
        : agg === "max"
          ? `MAX(${param})`
          : agg === "min"
            ? `MIN(${param})`
            : agg === "avg"
              ? `AVERAGE(${param})`
              : `MAX(${param})`);
    return `=BYCOL(${range}, LAMBDA(${param}, ${body}))`;
  },
  map: function (ctx) {
    const it = ctx.intent || {};
    const arrSpecs = Array.isArray(it.arrays)
      ? it.arrays
      : it.array
        ? [it.array]
        : [];
    if (!arrSpecs.length) return `=ERROR("MAP: arrays 파라미터가 필요합니다.")`;
    const normalized = arrSpecs.map((s) => _broadcastToColumn(s, ctx));
    const { vectors } = _alignManyToColumn(normalized, ctx);
    const params = (
      Array.isArray(it.lambda_params) && it.lambda_params.length
        ? it.lambda_params
        : ["x", "y", "z"]
    ).slice(0, vectors.length);
    const body = it.lambda_body || params[0] || "x";
    return `=MAP(${vectors.join(", ")}, LAMBDA(${params.join(", ")}, ${body}))`;
  },
  makearray: function (ctx) {
    const it = ctx.intent || {};
    const rows = Number(it.rows || it.m || 10);
    const cols = Number(it.cols || it.n || 1);
    const params = (
      Array.isArray(it.lambda_params) && it.lambda_params.length
        ? it.lambda_params
        : ["r", "c"]
    ).slice(0, 2);
    const [pr, pc] = params.length === 2 ? params : ["r", "c"];
    const body = it.lambda_body || pr;
    return `=MAKEARRAY(${rows}, ${cols}, LAMBDA(${pr}, ${pc}, ${body}))`;
  },

  vstack: (ctx) =>
    (() => {
      const it = ctx.intent || {};
      const src =
        Array.isArray(it.ranges) && it.ranges.length
          ? it.ranges
          : ["A1:B5", "A6:B10"];
      const parts = src.map((s) => rangeFromSpec(ctx, s) || s);
      if (it.ignore_blank_rows) {
        const filtered = parts.map(
          (p) => `FILTER(${p}, BYROW(${p}, LAMBDA(r, COUNTIF(r, "<>")>0)))`,
        );
        return `=VSTACK(${filtered.join(", ")})`;
      }
      return `=VSTACK(${parts.join(", ")})`;
    })(),
  hstack: (ctx) =>
    (() => {
      const it = ctx.intent || {};
      const src =
        Array.isArray(it.ranges) && it.ranges.length
          ? it.ranges
          : ["A1:B5", "C1:D5"];
      const parts = src.map((s) => rangeFromSpec(ctx, s) || s);
      return `=HSTACK(${parts.join(", ")})`;
    })(),
  tocol: (ctx) => {
    const it = ctx.intent || {};
    const rg = rangeFromSpec(ctx, it.range) || it.range || "A1:C5";
    const ignore = it.ignore ?? 0;
    const scan = it.scan ?? 0;
    return `=TOCOL(${rg}, ${ignore}, ${scan})`;
  },
  torow: (ctx) => {
    const it = ctx.intent || {};
    const rg = rangeFromSpec(ctx, it.range) || it.range || "A1:C5";
    const ignore = it.ignore ?? 0;
    const scan = it.scan ?? 0;
    return `=TOROW(${rg}, ${ignore}, ${scan})`;
  },
  transpose: (ctx) => {
    const it = ctx.intent || {};
    const rg = rangeFromSpec(ctx, it.range) || it.range || "A1:C5";
    return `=TRANSPOSE(${rg})`;
  },
  wraprows: (ctx) =>
    (() => {
      const it = ctx.intent || {};
      const vec = rangeFromSpec(ctx, it.vector) || it.vector || "A1:L1";
      const cnt = it.wrap_count || 3;
      const pad = it.pad_with != null ? _q(it.pad_with) : "";
      return pad
        ? `=WRAPROWS(${vec}, ${cnt}, ${pad})`
        : `=WRAPROWS(${vec}, ${cnt})`;
    })(),
  wrapcols: (ctx) =>
    (() => {
      const it = ctx.intent || {};
      const vec = rangeFromSpec(ctx, it.vector) || it.vector || "A1:A12";
      const cnt = it.wrap_count || 3;
      const pad = it.pad_with != null ? _q(it.pad_with) : "";
      return pad
        ? `=WRAPCOLS(${vec}, ${cnt}, ${pad})`
        : `=WRAPCOLS(${vec}, ${cnt})`;
    })(),

  _align_to: function (ctx, formatValue) {
    const it = ctx.intent || {};
    const tr = rangeFromSpec(
      ctx,
      it.target_range ||
        it.target ||
        it.best ||
        (ctx.bestReturn &&
          `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}:${ctx.bestReturn.columnLetter}${ctx.bestReturn.lastDataRow}`),
    );
    if (!tr) return `=ERROR("ALIGN_TO: target_range 없음")`;
    const ve = it.value_expr || formatValue(it.value);
    const helper = _alignTo(tr, ve);
    return `=${helper.asColumn}`; // 필요 시 indexAt는 호출부의 BYROW 안에서 사용
  },
  expand: function (ctx) {
    const it = ctx.intent || {};
    const arr = rangeFromSpec(ctx, it.array) || it.array || "A1:B2";
    const args = [arr];
    if (it.rows != null) args.push(String(it.rows));
    if (it.cols != null) args.push(String(it.cols));
    if (it.pad_with != null) args.push(_q(it.pad_with));
    return `=EXPAND(${args.join(", ")})`;
  },

  // ---------------------- TOP N (GROUPED) ----------------------
  // 예) "부서별 연봉 Top3"  → group_by + sort_by + top_n/limit 기반
  topn_grouped: function (ctx) {
    const { bestReturn, intent, allSheetsData } = ctx;
    if (!bestReturn) return `=ERROR("반환할 열을 찾을 수 없습니다.")`;

    const sheetName = bestReturn.sheetName;
    const sheetInfo = allSheetsData?.[sheetName];
    if (!sheetInfo) return `=ERROR("시트 정보를 찾을 수 없습니다.")`;

    const groupHint = intent.group_by;
    if (!groupHint) return `=ERROR("topn_grouped: group_by가 필요합니다.")`;

    const sortHint = intent.sort_by || intent.order_by;
    if (!sortHint)
      return `=ERROR("topn_grouped: sort_by(정렬 기준)가 필요합니다.")`;

    const n =
      Number(intent.top_n ?? intent.limit ?? intent.n ?? intent.count ?? 3) ||
      3;
    const sortOrder =
      String(intent.sort_order || intent.order || "desc").toLowerCase() ===
      "asc"
        ? 1
        : -1;

    // 그룹 열
    const gTerm = formulaUtils.expandTermsFromText(
      typeof groupHint === "string" ? groupHint : groupHint.header || "",
    );
    const gCol = formulaUtils.bestHeaderInSheet(
      sheetInfo,
      sheetName,
      gTerm,
      "lookup",
    );
    if (!gCol?.col) return `=ERROR("group_by: 키 열을 찾을 수 없습니다.")`;
    if (gCol.isAmbiguous)
      return `=ERROR("group_by 열이 모호합니다: '${gCol.header}' 또는 '${
        gCol.runnerUp?.header || "다른 후보"
      }' 중 선택이 필요합니다.")`;

    const groupRange = `'${sheetName}'!${gCol.col.columnLetter}${sheetInfo.startRow}:${gCol.col.columnLetter}${sheetInfo.lastDataRow}`;

    // 정렬 열
    const sTerm = formulaUtils.expandTermsFromText(
      typeof sortHint === "string" ? sortHint : sortHint.header || "",
    );
    const sCol = formulaUtils.bestHeaderInSheet(
      sheetInfo,
      sheetName,
      sTerm,
      "lookup",
    );
    if (!sCol?.col) return `=ERROR("sort_by: 기준 열을 찾을 수 없습니다.")`;
    if (sCol.isAmbiguous)
      return `=ERROR("sort_by 열이 모호합니다: '${sCol.header}' 또는 '${
        sCol.runnerUp?.header || "다른 후보"
      }' 중 선택이 필요합니다.")`;
    const sortRange = `'${sheetName}'!${sCol.col.columnLetter}${sheetInfo.startRow}:${sCol.col.columnLetter}${sheetInfo.lastDataRow}`;

    // 반환열 (명시되면 그걸 우선, 없으면 bestReturn 단일)
    const headerOpts =
      intent.return_headers || intent.select_headers || intent.return_cols;
    const metaEntries = Object.entries(sheetInfo.metaData || {});
    metaEntries.sort((a, b) => {
      const ai = formulaUtils.columnLetterToIndex(a[1].columnLetter);
      const bi = formulaUtils.columnLetterToIndex(b[1].columnLetter);
      return ai - bi;
    });
    const nameToIndex = new Map(metaEntries.map(([h, m], i) => [h, i + 1]));

    const wantedHeaders =
      Array.isArray(headerOpts) && headerOpts.length
        ? headerOpts
        : [bestReturn.header || bestReturn.columnLetter];

    const wantedRanges = [];
    const wantedLabels = [];
    for (const hSpec of wantedHeaders) {
      const hName =
        typeof hSpec === "string"
          ? (
              hSpec.match(/^\s*'?([^'!]+)'?\s*!\s*(.+)\s*$/)?.[2] || hSpec
            ).trim()
          : String(hSpec?.header || "").trim();
      if (!hName) continue;
      const idx = nameToIndex.get(hName);
      if (!idx) continue;
      const meta = metaEntries[idx - 1]?.[1];
      if (!meta?.columnLetter) continue;
      wantedRanges.push(
        `'${sheetName}'!${meta.columnLetter}${sheetInfo.startRow}:${meta.columnLetter}${sheetInfo.lastDataRow}`,
      );
      wantedLabels.push(hName);
    }
    if (!wantedRanges.length)
      return `=ERROR("topn_grouped: 반환할 열을 찾을 수 없습니다.")`;

    // 조건 마스크(공용)
    const { maskExpr, earlyError } = _buildMaskExprFromIntent(
      ctx,
      sheetInfo,
      sheetName,
    );
    if (earlyError) return earlyError;

    const groupLabel =
      String(gCol.header || groupHint.header || groupHint).trim() || "group";
    const headerRow = `{${[_q(groupLabel), ...wantedLabels.map(_q)].join(",")}}`;

    // data: [wanted..., sortKey] 로 만들고 정렬/TAKE 후 sortKey 컬럼 제거
    const dataWithSort = `HSTACK(${wantedRanges.join(", ")}, ${sortRange})`;
    const sortKeyIndex = wantedRanges.length + 1;
    const pickCols = Array.from(
      { length: wantedRanges.length },
      (_, i) => i + 1,
    ).join(", ");

    const gList = `SORT(UNIQUE(IFERROR(FILTER(${groupRange}, ${maskExpr}), )))`;

    return `=LET(_g, ${gList}, _hdr, ${headerRow}, _body, REDUCE("", _g, LAMBDA(acc, k, LET(_m, (${groupRange}=k)*(${maskExpr}), _d, IFERROR(FILTER(${dataWithSort}, _m), ), _s, IFERROR(SORT(_d, ${sortKeyIndex}, ${sortOrder}), ), _t, IFERROR(TAKE(CHOOSECOLS(_s, ${pickCols}), ${n}), ), VSTACK(acc, HSTACK(EXPAND(k, ROWS(_t), 1), _t)) )))), IF(LEN(_body)=0, _hdr, VSTACK(_hdr, FILTER(_body, INDEX(_body,,2)<>\"\"))))`;
  },
};

// ---- 정렬 파이프 헬퍼: FILTER/CHOOSECOLS/HSTACK 결과에 SORT or SORTBY 적용 ----
function pipeSortIfRequested(ctx, intent, expr, selectedIndexMap) {
  const fmt = (x) => String(x || "").trim();
  const sortKey = intent.sort_by || intent.order_by;
  if (!sortKey) return `=${expr}`;

  if (Array.isArray(sortKey) && sortKey.length) {
    const pairs = [];
    for (const k of sortKey) {
      const name = String(k.header || k).trim();
      const ord =
        String(k.order || intent.sort_order || "desc").toLowerCase() === "asc"
          ? 1
          : -1;
      const idxFromSelected = selectedIndexMap?.get?.(name);
      if (idxFromSelected) {
        pairs.push(`CHOOSECOLS(${expr}, ${idxFromSelected})`, ord);
      } else {
        const sheetInfo = ctx.allSheetsData[ctx.bestReturn.sheetName];
        const meta = Object.entries(sheetInfo.metaData || {}).sort(
          (a, b) =>
            formulaUtils.columnLetterToIndex(a[1].columnLetter) -
            formulaUtils.columnLetterToIndex(b[1].columnLetter),
        );
        const map = new Map(meta.map(([h, m], i) => [h, i + 1]));
        const idx = map.get(name);
        if (idx) pairs.push(`CHOOSECOLS(${expr}, ${idx})`, ord);
      }
    }
    return pairs.length ? `=SORTBY(${expr}, ${pairs.join(", ")})` : `=${expr}`;
  }

  const sheetInfo = ctx.allSheetsData[ctx.bestReturn.sheetName];
  const meta = Object.entries(sheetInfo.metaData || {}).sort(
    (a, b) =>
      formulaUtils.columnLetterToIndex(a[1].columnLetter) -
      formulaUtils.columnLetterToIndex(b[1].columnLetter),
  );
  const nameToIndex = new Map(meta.map(([h, m], i) => [h, i + 1]));
  const idx =
    selectedIndexMap?.get?.(fmt(sortKey)) || nameToIndex.get(fmt(sortKey));
  const order =
    String(intent.sort_order || "desc").toLowerCase() === "asc" ? 1 : -1;

  if (idx) return `=SORTBY(${expr}, CHOOSECOLS(${expr}, ${idx}), ${order})`;

  const joinSpec = (intent.joins || [])[0];
  if (
    joinSpec &&
    joinSpec.sheet &&
    Array.isArray(joinSpec.on) &&
    joinSpec.on.length
  ) {
    const rightInfo = ctx.allSheetsData[joinSpec.sheet];
    const leftRanges = [];
    const rightRanges = [];
    for (const pair of joinSpec.on) {
      const lTerm = formulaUtils.expandTermsFromText(pair.left);
      const rTerm = formulaUtils.expandTermsFromText(pair.right);
      const lCol = formulaUtils.bestHeaderInSheet(
        sheetInfo,
        ctx.bestReturn.sheetName,
        lTerm,
        "lookup",
      );
      const rCol = formulaUtils.bestHeaderInSheet(
        rightInfo,
        joinSpec.sheet,
        rTerm,
        "lookup",
      );
      if (!lCol?.col || !rCol?.col) continue;
      // ✅ join 키가 모호하면 조인 자체가 "그럴듯하게 틀림"을 만들기 쉬움 → 스킵
      if (lCol.isAmbiguous || rCol.isAmbiguous) continue;
      leftRanges.push(
        `'${ctx.bestReturn.sheetName}'!${lCol.col.columnLetter}${sheetInfo.startRow}:${lCol.col.columnLetter}${sheetInfo.lastDataRow}`,
      );
      rightRanges.push(
        `'${joinSpec.sheet}'!${rCol.col.columnLetter}${rightInfo.startRow}:${rCol.col.columnLetter}${rightInfo.lastDataRow}`,
      );
    }
    const sortHdr = formulaUtils.bestHeaderInSheet(
      rightInfo,
      joinSpec.sheet,
      formulaUtils.expandTermsFromText(sortKey),
      "lookup",
    );
    if (rightRanges.length && sortHdr?.col) {
      const rightSortRange = `'${joinSpec.sheet}'!${sortHdr.col.columnLetter}${rightInfo.startRow}:${sortHdr.col.columnLetter}${rightInfo.lastDataRow}`;
      const Lvec =
        leftRanges.length === 1
          ? _normRange(leftRanges[0])
          : _concatKeyVector(leftRanges);
      const Rvec =
        rightRanges.length === 1
          ? _normRange(rightRanges[0])
          : _concatKeyVector(rightRanges);
      // Step3: 조인 기반 sortKey도 행 단위로 안정화
      return `=LET(LK, ${Lvec}, RK, ${Rvec}, SV, MAP(LK, LAMBDA(k, XLOOKUP(k, RK, ${rightSortRange}, , 0))), SORTBY(${expr}, SV, ${order}))`;
    }
  }
  return `=${expr}`;
}

module.exports = arrayFunctionBuilder;
