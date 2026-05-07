const XLSX = require("xlsx");
const { buildAllSheetsData } = require("./sheetMetaBuilder");
const { CLUSTER_DEFS, inferClusterFromText } = require("./clusterSchema");

const SCORING_WEIGHTS = {
  EXACT_MATCH: 30,
  PARTIAL_MATCH: 2,
  SYNONYM_MATCH: 8,
  SHEET_NAME_BONUS: 1.5,
  NUMERIC_COLUMN_BONUS: 3,
  NUMERIC_COLUMN_PENALTY: -5,
  CLUSTER_MATCH: 20,
  ROLE_MATCH: 12,
  TYPE_MATCH: 8,
};

function columnLetterToIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
  }
  return column;
}

/* =========================================
   날짜 계산 유틸 (내부 종속성 정리 완료)
========================================= */
function _dateRelativeExpr(ctx, opts = {}) {
  const mode = (opts.mode || "calendar").toLowerCase();
  const base =
    opts.base != null
      ? typeof opts.base === "string"
        ? opts.base
        : String(opts.base)
      : "TODAY()";
  const n = Number(opts.offset_days || 0);
  if (mode === "eomonth") return `EOMONTH(${base}, ${n})`;
  if (mode === "workday") {
    const mask = opts.weekend_mask && `"${opts.weekend_mask}"`;
    const hol = opts.holidays ? rangeFromSpec(ctx, opts.holidays) : null;
    if (mask)
      return hol
        ? `WORKDAY.INTL(${base}, ${n}, ${mask}, ${hol})`
        : `WORKDAY.INTL(${base}, ${n}, ${mask})`;
    return hol ? `WORKDAY(${base}, ${n}, ${hol})` : `WORKDAY(${base}, ${n})`;
  }
  return n ? `(${base})+${n}` : `${base}`;
}

function buildDateWindowPairs(ctx, windowObj = {}) {
  const unit = (windowObj.unit || "days").toLowerCase();
  if (unit !== "days") return [];
  const size = Number(windowObj.size || 0);
  if (!size) return [];

  const rr = resolveHeaderRef(
    ctx,
    windowObj.header || windowObj.date_header || "날짜",
    windowObj.sheet,
  );
  if (!rr) return [];

  const upperInc = windowObj.include_upper !== false;
  const mode = windowObj.mode || "calendar";
  const holidays = windowObj.holidays;
  const weekend_mask = windowObj.weekend_mask;
  const upperBase = windowObj.base || "TODAY()";

  const startExpr = _dateRelativeExpr(ctx, {
    mode,
    base: upperBase,
    offset_days: -size,
    holidays,
    weekend_mask,
  });
  const endExpr = _dateRelativeExpr(ctx, {
    mode,
    base: upperBase,
    offset_days: 0,
    holidays,
    weekend_mask,
  });

  const ge = `">="&${startExpr}`;
  const le = upperInc ? `"<="&${endExpr}` : `"<"&${endExpr}`;
  return [rr.range, ge, rr.range, le];
}

/* =========================================
   내부 헬퍼(로컬 범위 해석기)
========================================= */
function resolveHeaderRef(ctx, headerText, sheetHint) {
  if (!ctx?.allSheetsData) return null;
  const term = expandTermsFromText(headerText);
  let sheets = ctx.allSheetsData;
  if (sheetHint) {
    sheets = Object.fromEntries(
      Object.entries(sheets).filter(([n]) => n === sheetHint),
    );
  }
  const col = findBestColumnAcrossSheets(sheets, term, "lookup");
  if (!col) return null;
  const s = `'${col.sheetName}'!`;
  const c = col.columnLetter;
  const st = col.startRow || 2;
  const en = col.lastDataRow || col.rowCount || 1048576;
  return {
    sheetName: col.sheetName,
    range: `${s}${c}${st}:${c}${en}`,
    cell: `${s}${c}${st}`,
  };
}

function rangeFromSpec(ctx, spec) {
  if (!spec) return null;
  if (typeof spec === "string") {
    const m = spec.match(/^\s*'?([^'!]+)'?\s*!\s*(.+)\s*$/);
    if (m) {
      const r = resolveHeaderRef(ctx, m[2].trim(), m[1].trim());
      return r ? r.range : null;
    }
    const r = resolveHeaderRef(ctx, spec, null);
    return r ? r.range : spec;
  }
  if (typeof spec === "object" && spec.header) {
    const r = resolveHeaderRef(ctx, spec.header, spec.sheet || null);
    return r ? r.range : null;
  }
  return null;
}

/* =========================================
   문자열/비교 유틸
========================================= */
function _quoteString(s) {
  return `"${String(s).replace(/"/g, '""')}"`;
}

/**
 * "A1", "A1:A10", "a1부터 a10까지" 같은 텍스트에서
 * 명시적인 셀/범위만 추출해서 "A1" 또는 "A1:A10" 형태로 반환
 * 못 찾으면 null
 */
function parseExplicitCellOrRange(text = "") {
  const upper = String(text).toUpperCase();

  // 1) 이미 A1:A10 형태로 들어온 경우
  const rangeMatch = upper.match(/[A-Z]+[0-9]+:[A-Z]+[0-9]+/);
  if (rangeMatch) return rangeMatch[0];

  // 2) 전체 열 범위 표기: B:B
  const fullColumnRangeMatch = upper.match(/\b([A-Z]{1,3}):([A-Z]{1,3})\b/);
  if (fullColumnRangeMatch) {
    return `${fullColumnRangeMatch[1]}:${fullColumnRangeMatch[2]}`;
  }

  // 3) "B열", "AA열" 같은 단일 열 표기
  const koreanColumnMatch = upper.match(/\b([A-Z]{1,3})\s*열\b/);
  if (koreanColumnMatch) {
    return `${koreanColumnMatch[1]}:${koreanColumnMatch[1]}`;
  }

  // 4) "A1부터 A10까지" / "a1부터a10까지" 처럼 셀 두 개만 있는 경우
  const cells = upper.match(/[A-Z]+[0-9]+/g);
  if (cells && cells.length >= 2) {
    return `${cells[0]}:${cells[1]}`;
  }

  // 5) "A1" 하나만 있는 경우
  if (cells && cells.length === 1) {
    return cells[0];
  }

  return null;
}

function formatValue(value, options = {}) {
  const { trim_text = true, coerce_number = true, forceText = false } = options;
  if (value == null) return '""';

  // --- Helpers: keep A1 refs / ranges unquoted ---
  const isA1RefOrRange = (s) => {
    const t = String(s || "").trim();
    // A1, $A$1, A1:B10, $A$1:$B$10
    if (/^\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?$/i.test(t))
      return true;
    // Sheet!A1 or 'Sheet Name'!A1 and ranges
    if (
      /^([^!'\s]+|'[^']+')!\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?$/i.test(
        t,
      )
    )
      return true;
    return false;
  };

  // --- Helper: ISO-like date string -> DATEVALUE("YYYY-MM-DD") ---
  const toIsoDateValueExpr = (s) => {
    const t = String(s || "").trim();
    if (!/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(t)) return null;
    const iso = t.replace(/[./]/g, "-");
    return `DATEVALUE(${_quoteString(iso)})`;
  };

  if (
    typeof value === "string" &&
    /^\s*(NOW\(\)|TODAY\(\)|DATE\(|EOMONTH\(|WORKDAY\()/.test(value)
  )
    return value.trim();

  // ✅ 셀/범위 참조는 따옴표로 감싸지 않는다. (예: J2, A1:A10, '나무'!H91:H177)
  if (typeof value === "string" && isA1RefOrRange(value)) return value.trim();

  if (
    !forceText &&
    coerce_number &&
    (typeof value === "number" ||
      (typeof value === "string" && /^-?\d+(\.\d+)?$/.test(value.trim())))
  )
    return String(Number(value));

  // ✅ 날짜 리터럴("2023-01-01")은 DATEVALUE로 변환 (forceText면 유지)
  if (!forceText && typeof value === "string") {
    const dv = toIsoDateValueExpr(value);
    if (dv) return dv;
  }

  if (typeof value === "string") {
    const s = trim_text ? value.trim() : value;
    return _quoteString(s);
  }

  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return _quoteString(String(value));
}

/* =========================================
   범위 유틸 (Vector Alignment 등)
========================================= */
function _isRangeString(s) {
  if (typeof s !== "string") return false;
  return /!/.test(s) && /:/.test(s);
}

function _toVectorExpr(exprOrRange) {
  if (!exprOrRange) return null;
  const s = String(exprOrRange);
  if (_isRangeString(s)) return `TOCOL(${s}, 1)`;
  return s;
}

function _rowsOf(exprOrRange) {
  const v = _toVectorExpr(exprOrRange);
  return v ? `ROWS(${v})` : "0";
}

function ALIGN_TO(left, right) {
  const lVec = _toVectorExpr(left);
  const rVec = _toVectorExpr(right);

  if (lVec && rVec) {
    const lenExpr = `MIN(${_rowsOf(lVec)}, ${_rowsOf(rVec)})`;
    return {
      leftVec: `TAKE(${lVec}, ${lenExpr})`,
      rightVec: `TAKE(${rVec}, ${lenExpr})`,
      rowsExpr: lenExpr,
    };
  }
  if (lVec && !rVec) {
    const lenExpr = _rowsOf(lVec);
    const scalar = String(right ?? '""');
    return {
      leftVec: lVec,
      rightVec: `EXPAND(${scalar}, ${lenExpr}, 1)`,
      rowsExpr: lenExpr,
    };
  }
  if (!lVec && rVec) {
    const lenExpr = _rowsOf(rVec);
    const scalar = String(left ?? '""');
    return {
      leftVec: `EXPAND(${scalar}, ${lenExpr}, 1)`,
      rightVec: rVec,
      rowsExpr: lenExpr,
    };
  }
  return {
    leftVec: String(left ?? '""'),
    rightVec: String(right ?? '""'),
    rowsExpr: "1",
  };
}

/* =========================================
   비교/검색 유틸
========================================= */
function _eq(left, right, { case_sensitive } = {}) {
  return case_sensitive ? `EXACT(${left}, ${right})` : `(${left}=${right})`;
}

function _contains(rangeExpr, needleExpr, { case_sensitive } = {}) {
  return case_sensitive
    ? `ISNUMBER(FIND(${needleExpr}, ${rangeExpr}))`
    : `ISNUMBER(SEARCH(${needleExpr}, ${rangeExpr}))`;
}

function _startsWith(rangeExpr, needleExpr, { case_sensitive } = {}) {
  return case_sensitive
    ? `EXACT(LEFT(${rangeExpr}, LEN(${needleExpr})), ${needleExpr})`
    : `LOWER(LEFT(${rangeExpr}, LEN(${needleExpr})))=LOWER(${needleExpr})`;
}

function _endsWith(rangeExpr, needleExpr, { case_sensitive } = {}) {
  return case_sensitive
    ? `EXACT(RIGHT(${rangeExpr}, LEN(${needleExpr})), ${needleExpr})`
    : `LOWER(RIGHT(${rangeExpr}, LEN(${needleExpr})))=LOWER(${needleExpr})`;
}

/* =========================================
   에러 처리 유틸
========================================= */
function wrapIfError(expr, ctxOrDefault) {
  const s = String(expr || "").trim();
  if (/^=?\s*ERROR\(/i.test(s)) return s.startsWith("=") ? s : "=" + s;
  const inner = s.startsWith("=") ? s.slice(1) : s;
  const v = ctxOrDefault?.policy?.value_if_error ?? "";
  return `IFERROR(${inner}, ${_quoteString(String(v))})`;
}

function wrapIfNA(expr, ctxOrDefault) {
  const v = ctxOrDefault?.policy?.value_if_not_found ?? "";
  return `IFNA(${expr}, ${_quoteString(String(v))})`;
}

function fallbackNotFoundArg(ctxOrDefault) {
  const v = ctxOrDefault?.policy?.value_if_not_found ?? "";
  return _quoteString(String(v));
}

/* =========================================
   캐시 IO (DISABLED)
========================================= */
function readCache(conversionType) {
  return { normalized: {}, direct: {}, fewShots: [] };
}

function writeCache(cacheData, conversionType) {
  return;
}

/* =========================================
   시트 열 탐색 & 스코어링
========================================= */
function isNumericLike(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "" || /[^\d.+\-eE]/.test(s)) return false;
  const n = Number(s);
  return Number.isFinite(n);
}

function indexToColumnLetter(idx) {
  let n = idx + 1,
    s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function inferDesiredRole(operation = "") {
  const op = String(operation || "").toLowerCase();

  if (["lookup", "xlookup"].includes(op)) return "lookup";
  if (["group", "group_by"].includes(op)) return "group";

  if (
    [
      "return",
      "average",
      "sum",
      "stdev",
      "min",
      "max",
      "averageifs",
      "sumifs",
      "countifs",
      "minifs",
      "maxifs",
      "count",
    ].includes(op)
  ) {
    return "metric";
  }

  return null;
}

function inferExpectedType(operation = "") {
  const op = String(operation || "").toLowerCase();
  if (
    [
      "average",
      "sum",
      "stdev",
      "min",
      "max",
      "averageifs",
      "sumifs",
      "countifs",
      "minifs",
      "maxifs",
    ].includes(op)
  ) {
    return "number";
  }
  return null;
}

function bestHeaderInSheet(sheetInfo, sheetName, termSet, operation) {
  const MINIMUM_SCORE_THRESHOLD = 12;
  const AMBIGUOUS_GAP_THRESHOLD = 8; // 초기값(보수적). 운영 피드백으로 튜닝

  let best = { header: "", score: -1, col: null };
  let runnerUp = { header: "", score: -1, col: null };

  const meta = sheetInfo.metaData || {};
  for (const [header, metaInfo] of Object.entries(meta)) {
    const termText = [...termSet].join(" ");
    const s = scoreColumn(sheetName, header, metaInfo, termSet, operation, {
      desiredCluster: inferClusterFromText(termText),
      desiredRole: inferDesiredRole(operation),
      expectedType: inferExpectedType(operation),
    });
    if (s > best.score) {
      runnerUp = best;
      best = { header, score: s, col: metaInfo };
    } else if (s > runnerUp.score) {
      runnerUp = { header, score: s, col: metaInfo };
    }
  }
  if (best.score < MINIMUM_SCORE_THRESHOLD)
    return {
      header: "",
      score: 0,
      col: null,
      runnerUp: null,
      gap: 0,
      isAmbiguous: false,
    };

  const gap = best.score - (runnerUp?.score ?? -1);
  const hasStrongMatch =
    best.score >= SCORING_WEIGHTS.EXACT_MATCH || gap >= AMBIGUOUS_GAP_THRESHOLD;

  const isAmbiguous = !!(
    runnerUp?.col &&
    !hasStrongMatch &&
    gap < AMBIGUOUS_GAP_THRESHOLD
  );

  return {
    ...best,
    runnerUp,
    gap,
    isAmbiguous,
    ambiguityReason: isAmbiguous
      ? `후보 열이 모호합니다: ${best.header} / ${runnerUp.header}`
      : "",
  };
}

function findBestColumnAcrossSheets(allSheetsData, termSet, operation) {
  let winner = null;
  for (const [sheetName, sheetInfo] of Object.entries(allSheetsData || {})) {
    const cand = bestHeaderInSheet(sheetInfo, sheetName, termSet, operation);
    if (!cand.col) continue;

    const colMeta = cand.col || {};
    const hit = {
      sheetName,
      header: cand.header,
      score: cand.score,
      columnLetter: colMeta.columnLetter,
      startRow: colMeta.startRow || sheetInfo.startRow,
      lastDataRow: colMeta.lastRow || sheetInfo.lastDataRow,
      // ambiguity info (Top-2)
      isAmbiguous: !!cand.isAmbiguous,
      ambiguityGap: cand.gap ?? 0,
      runnerUpHeader: cand.runnerUp?.header || null,
      runnerUpColumnLetter: cand.runnerUp?.col?.columnLetter || null,
      ambiguityReason: cand.ambiguityReason || "",
    };

    if (!winner || hit.score > winner.score) winner = hit;
  }
  return winner;
}

/* =========================================
   텍스트/스코어링 동의어
========================================= */
const SYNONYMS = {
  // 🔥 역할 기반 구조 (완전 범용화 핵심)

  person: [
    "이름",
    "성명",
    "직원명",
    "사원명",
    "환자명",
    "학생명",
    "선수명",
    "고객명",
    "회원명",
    "담당자",
    "name",
    "person",
  ],

  organization: [
    "부서",
    "팀",
    "소속",
    "조직",
    "진료과",
    "학과",
    "학부",
    "반",
    "구단",
    "학교",
    "기관",
    "department",
    "team",
    "group",
    "organization",
  ],

  date: [
    "날짜",
    "일자",
    "방문일",
    "진료일",
    "입학일",
    "입사일",
    "경기일",
    "등록일",
    "생성일",
    "수정일",
    "시작일",
    "종료일",
    "date",
    "day",
    "visit date",
    "hire date",
  ],

  amount: [
    "연봉",
    "급여",
    "월급",
    "임금",
    "보수",
    "금액",
    "가격",
    "단가",
    "매출",
    "매출액",
    "수납액",
    "결제금액",
    "총액",
    "salary",
    "pay",
    "amount",
    "price",
    "cost",
    "revenue",
  ],

  score: [
    "점수",
    "성적",
    "평점",
    "학점",
    "득점",
    "포인트",
    "골",
    "score",
    "grade",
    "gpa",
    "points",
    "goals",
  ],

  quantity: [
    "수량",
    "개수",
    "건수",
    "인원수",
    "환자수",
    "학생수",
    "경기수",
    "판매수량",
    "quantity",
    "count",
    "qty",
  ],

  grade: [
    "등급",
    "평가등급",
    "평가 등급",
    "학년",
    "레벨",
    "등수",
    "grade",
    "rating",
    "level",
    "rank",
  ],

  id: [
    "id",
    "아이디",
    "사번",
    "직원id",
    "번호",
    "고유번호",
    "employeeid",
    "emp id",
  ],

  status: ["상태", "진행상태", "처리상태", "완료", "미완료", "status", "state"],
};

function getSynonyms(...keys) {
  const out = [];
  const seen = new Set();

  for (const key of keys.flat()) {
    const list = SYNONYMS[key] || [];
    for (const v of list) {
      const s = String(v || "").trim();
      if (!s) continue;

      const norm = s.toLowerCase();
      if (seen.has(norm)) continue;

      seen.add(norm);
      out.push(s);
    }
  }

  return out;
}

function includesAlias(text, ...keys) {
  const raw = String(text || "").toLowerCase();
  const aliases = getSynonyms(...keys);

  return aliases.some((kw) => raw.includes(String(kw).toLowerCase()));
}

function textIncludesAnyAlias(text, ...keys) {
  const raw = String(text || "").toLowerCase();
  return getSynonyms(...keys).some((kw) =>
    raw.includes(String(kw).toLowerCase()),
  );
}

function norm(s = "") {
  return String(s)
    .toLowerCase()
    .replace(/\(.*?\)/g, "")
    .replace(/[^\p{Letter}\p{Number}]+/gu, "")
    .trim();
}

function expandTermsFromText(text = "") {
  const base = norm(text);
  const terms = new Set([base]);

  const clusterKey = inferClusterFromText(text);
  if (clusterKey && CLUSTER_DEFS[clusterKey]?.aliases) {
    CLUSTER_DEFS[clusterKey].aliases.forEach((v) => terms.add(norm(v)));
  }

  if (!clusterKey) {
    for (const list of Object.values(SYNONYMS)) {
      const norms = list.map(norm);
      if (norms.includes(base)) {
        // 전체 확장 대신 base 제외 상위 몇 개만 제한 사용
        for (const v of list.slice(0, 8)) {
          const nv = norm(v);
          if (nv && nv !== base) terms.add(nv);
        }
        break;
      }
    }
  }

  return terms;
}

function sheetNameScore(sheetName, termSet) {
  const s = norm(sheetName);
  let score = 0;
  for (const t of termSet) if (s.includes(t)) score += 1.5;
  return score;
}

function scoreColumn(
  sheetName,
  header,
  meta,
  termSet,
  operation,
  options = {},
) {
  const h = norm(header);
  if (termSet.has(h)) return SCORING_WEIGHTS.EXACT_MATCH;

  let score = 0;
  let lexicalScore = 0;
  let clusterScore = 0;
  let roleScore = 0;
  let typeScore = 0;
  let synonymScore = 0;
  let profileScore = 0;

  const desiredCluster = options.desiredCluster || null;
  const desiredRole = options.desiredRole || null;
  const expectedType = options.expectedType || null;

  if (
    desiredCluster &&
    meta?.clusterCandidate &&
    String(meta.clusterCandidate) === String(desiredCluster)
  ) {
    clusterScore += SCORING_WEIGHTS.CLUSTER_MATCH;
    lexicalScore = Math.max(0, lexicalScore - 1);
  }

  if (
    desiredRole &&
    meta?.inferredRole &&
    String(meta.inferredRole) === String(desiredRole)
  ) {
    roleScore += SCORING_WEIGHTS.ROLE_MATCH;
  }

  const dominantType =
    meta?.profileType || meta?.clusterType || meta?.dominantType || null;

  const profileType = String(meta?.profileType || "").toLowerCase();
  const uniqueRatio = Number(meta?.uniqueRatio ?? 1);
  const uniqueCount = Number(meta?.uniqueCount ?? 0);
  const numericRatio = Number(meta?.numericRatio ?? 0);
  const dateRatio =
    Number(meta?.dateRatio ?? 0) + Number(meta?.datetimeRatio ?? 0);

  const op = String(operation || "").toLowerCase();

  const wantsNumber =
    expectedType === "number" ||
    [
      "sum",
      "average",
      "avg",
      "mean",
      "min",
      "max",
      "median",
      "stdev",
      "rank",
      "rankcolumn",
      "topn",
      "topnrows",
    ].includes(op);

  const wantsDate =
    expectedType === "date" ||
    /(date|날짜|일자|최근|최신|이후|이전|입사|방문|경기|등록)/i.test(
      [...termSet].join(" "),
    );

  const wantsCategory =
    desiredRole === "group" ||
    ["group", "group_by", "countifs", "averageifs", "sumifs"].includes(op);

  if (wantsNumber) {
    if (profileType === "number" || numericRatio >= 0.8) profileScore += 6;
    if (profileType === "date" || dateRatio >= 0.6) profileScore -= 4;
    if (profileType === "text" && numericRatio < 0.3) profileScore -= 3;
  }

  if (wantsDate) {
    if (profileType === "date" || dateRatio >= 0.6) profileScore += 6;
    if (profileType === "number" && dateRatio < 0.2) profileScore -= 2;
  }

  if (wantsCategory) {
    if (profileType === "category") profileScore += 5;
    if (uniqueRatio <= 0.4 && uniqueCount > 1) profileScore += 3;
    if (profileType === "number" && uniqueRatio > 0.7) profileScore -= 3;
  }

  if (
    expectedType &&
    dominantType &&
    String(dominantType) === String(expectedType)
  ) {
    typeScore += SCORING_WEIGHTS.TYPE_MATCH;
  }

  if ([...termSet].some((t) => h.includes(t) || t.includes(h)))
    lexicalScore += SCORING_WEIGHTS.PARTIAL_MATCH;
  else {
    for (const list of Object.values(SYNONYMS)) {
      const nlist = list.map(norm);
      const termHit = [...termSet].some((t) => nlist.includes(t));
      if (termHit && nlist.some((a) => h === a)) {
        synonymScore += SCORING_WEIGHTS.SYNONYM_MATCH;
        break;
      }
    }
  }
  lexicalScore += sheetNameScore(sheetName, termSet);

  const needNumeric = [
    "average",
    "sum",
    "stdev",
    "min",
    "max",
    "averageifs",
    "sumifs",
    "countifs",
    "minifs",
    "maxifs",
  ];
  if (needNumeric.includes(operation)) {
    if (meta.numericRatio >= 0.8)
      typeScore += SCORING_WEIGHTS.NUMERIC_COLUMN_BONUS;
    else if (meta.numericRatio >= 0.4) typeScore += 1;
    else typeScore += SCORING_WEIGHTS.NUMERIC_COLUMN_PENALTY;

    if (meta?.clusterType === "ordered_text") typeScore -= 2;
    if (meta?.clusterType === "date") typeScore -= 1;
  }

  score =
    lexicalScore +
    clusterScore +
    roleScore +
    typeScore +
    synonymScore +
    profileScore;

  return score;
}

/* =========================================
   파일 전처리
========================================= */
function preprocessFileData(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const allSheetsData = buildAllSheetsData(workbook);
    return { allSheetsData };
  } catch (error) {
    console.error("파일 파싱 오류:", error);
    return { error: "파일 파싱에 실패했습니다." };
  }
}

function findBestSheetAndColumns(allSheetsData, searchTerms, options = {}) {
  const data = allSheetsData || {};
  const { return: retTerms, lookup: lookupTerms } = searchTerms || {};

  const bestReturn = retTerms
    ? findBestColumnAcrossSheets(data, expandTermsFromText(retTerms), "return")
    : null;

  const bestLookup = lookupTerms
    ? findBestColumnAcrossSheets(
        data,
        expandTermsFromText(lookupTerms),
        "lookup",
      )
    : null;

  let totalScore = (bestReturn?.score || 0) + (bestLookup?.score || 0);

  if (
    bestReturn &&
    bestLookup &&
    bestReturn.sheetName === bestLookup.sheetName
  ) {
    totalScore += options.sameSheetBonus ?? 0.2;
  }

  if (options.typeHints) {
    const adjust = (meta, expected, base) => {
      if (!meta?.dominantType) return base; // 아직 preprocess에서 dominantType을 안 채우고 있으니 no-op
      return meta.dominantType === expected ? base + 0.1 : base;
    };
    if (bestReturn)
      totalScore = adjust(bestReturn, options.typeHints.return, totalScore);
    if (bestLookup)
      totalScore = adjust(bestLookup, options.typeHints.lookup, totalScore);
  }

  return {
    sheetName: bestReturn?.sheetName || bestLookup?.sheetName || null,
    return: bestReturn,
    lookup: bestLookup,
    totalScore,
  };
}

/* =========================================
   Export
========================================= */
module.exports = {
  _quoteString,
  formatValue,
  _dateRelativeExpr,
  buildDateWindowPairs,
  resolveHeaderRef,
  rangeFromSpec,

  _isRangeString,
  _toVectorExpr,
  _rowsOf,
  ALIGN_TO,

  _eq,
  _contains,
  _startsWith,
  _endsWith,
  wrapIfError,
  wrapIfNA,
  fallbackNotFoundArg,

  parseExplicitCellOrRange,

  readCache,
  writeCache,

  preprocessFileData,
  findBestSheetAndColumns,
  findBestColumnAcrossSheets,
  bestHeaderInSheet,
  expandTermsFromText,
  norm,
  columnLetterToIndex,
  indexToColumnLetter,
  isNumericLike,
  SYNONYMS,
  getSynonyms,
  textIncludesAnyAlias,
  includesAlias,
};
