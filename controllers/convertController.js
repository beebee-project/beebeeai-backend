const { cleanAIResponse } = require("../utils/responseHelper");
const formulaUtils = require("../utils/formulaUtils");
const { getOrBuildAllSheetsData } = require("../utils/sheetPreprocessor");
const { buildIntentCacheKey } = require("../utils/intentCacheKeyBuilder");
const { makeSheetStateSig } = require("../utils/sheetStateSig");
const intentCache = require("../services/intentCache");
const { writeRequestLog } = require("../services/requestLogService");
const crypto = require("crypto");
const { classifyReason } = require("../utils/reasonClassifier");
const { validateFormula } = require("../utils/outputValidator");
const { buildDebugMeta } = require("../utils/debugMetaBuilder");
const {
  normalizeIntentSchema,
  normalizePolicy,
} = require("../utils/intentSchema");
const {
  resolveIntent,
  buildResolvedContext,
} = require("../utils/intentResolver");
const { detectFormulaCompatibility } = require("../utils/formulaCompatibility");
const { buildConditionMask } = require("../utils/conditionEngine");

// === 빌더 모음 ===
const logicalFunctionBuilder = require("../builders/logicalFunctions");
const mathStatsFunctionBuilder = require("../builders/mathStatsFunctions");
const dateFunctionBuilder = require("../builders/dateFunctions");
const referenceFunctionBuilder = require("../builders/referenceFunctions");
const textFunctionBuilder = require("../builders/textFunctions");
const arrayFunctionBuilder = require("../builders/arrayFunctions");
const direct = require("../builders/direct");

const { bumpUsage, assertCanUse } = require("../services/usageService");

const { shouldLogCache } = require("../utils/cacheLog");
const { appendFeedback } = require("../utils/feedbackStore");

/* ---------------------------------------------
 * Lazy singletons (OpenAI / GCS Bucket)
 * -------------------------------------------*/
let _openai = null;
function getOpenAI() {
  try {
    if (_openai) return _openai;
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) return null;
    const OpenAI = require("openai");
    _openai = new OpenAI({ apiKey });
    return _openai;
  } catch {
    return null;
  }
}

let _bucket = null;
function getBucket() {
  try {
    if (_bucket) return _bucket;
    const name = process.env.GCS_BUCKET_NAME;
    if (!name) return null;
    const { Storage } = require("@google-cloud/storage");
    const storage = new Storage();
    _bucket = storage.bucket(name);
    return _bucket;
  } catch {
    return null;
  }
}

/**

@typedef {Object} HeaderSpec

@property {string} [header] - 열 이름 (예: "매출액")

@property {string} [sheet] - 시트 이름

@property {string} [columnLetter] - 열 문자 (예: "B")
*/

/**

@typedef {Object} ConditionLeaf

@property {string|HeaderSpec} [target] - 조건의 기준 열 또는 셀

@property {string} [operator] - 비교 연산자 (예: "=", ">", "<", "contains")

@property {string|number|boolean|HeaderSpec} [value] - 비교 값
*/

/**

@typedef {Object} ConditionGroup

@property {"AND"|"OR"} logical_operator - 논리 연산자

@property {Array<ConditionLeaf|ConditionGroup>} conditions - 조건 리스트
*/

/**

@typedef {ConditionLeaf|ConditionGroup} ConditionNode
*/

/**

@typedef {Object} DateWindow

@property {"days"} [type] - 윈도우 단위

@property {number} [size] - 윈도우 크기 (예: 최근 7일)

@property {string} [date_header] - 날짜 기준 열 이름
*/

/**

@typedef {Object} RowSelector

@property {string|HeaderSpec} [hint] - 기준 열 또는 셀

@property {string|number|boolean} [value] - 선택할 값

@property {string} [sheet] - 시트 이름
*/

/**

@typedef {Object} Intent

@property {string} operation - 수행할 연산 (예: "sum", "xlookup", "filter")

@property {"excel"|"sheets"} [engine] - 실행 대상 엔진

@property {"strict"|"normal"} [mode] - 정책 모드

@property {string|number|boolean} [value_if_not_found] - 찾지 못했을 때 기본값

@property {string|number|boolean} [value_if_error] - 오류 시 기본값

@property {string} [header_hint] - 주 대상 열 힌트

@property {string} [lookup_hint] - 조회 기준 열 힌트

@property {string} [return_hint] - 반환 대상 열 힌트

@property {Array<ConditionNode>} [conditions] - 조건 리스트

@property {ConditionNode} [condition] - 단일 조건 (IF 등)

@property {DateWindow} [window] - 기간 조건

@property {RowSelector} [row_selector] - 특정 행 선택용

@property {string} [group_by] - 그룹 기준 열

@property {string|number} [lookup_value] - 조회 값

@property {string|number|boolean} [value_if_true] - 조건 만족 시 값

@property {string|number|boolean} [value_if_false] - 조건 불만족 시 값

@property {Array<string|number>} [in_values] - 다중 비교 값

@property {string} [delimiter] - 텍스트 구분자

@property {Array<string>} [delimiters] - 다중 구분자

@property {boolean} [ignore_empty] - 빈 텍스트 무시 여부

@property {boolean} [remove_empty_text] - 빈 문자열 제거 여부

@property {string} [message] - 사용자 메시지 또는 오류 문구
*/

/**

@typedef {Object} SheetMeta

@property {number} rowCount - 총 행 수

@property {number} startRow - 데이터 시작 행

@property {number} lastDataRow - 데이터 마지막 행

@property {Object<string, {columnLetter: string, numericRatio: number}>} metaData - 열 메타 정보
*/

/**

@typedef {Object} ColumnRange

@property {string} sheetName - 시트 이름

@property {string} columnLetter - 열 문자 (예: "B")

@property {string} header - 열 이름

@property {number} startRow - 시작 행

@property {number} lastDataRow - 마지막 행

@property {string} range - 실제 Excel 범위 (예: "Sheet1!B2:B100")
*/

/**

@typedef {Object} Context

@property {Intent} intent - 현재 요청의 의도 객체

@property {"excel"|"sheets"} engine - 엔진 종류

@property {Object} policy - 에러 처리 및 정책

@property {Object} formatOptions - 값 포맷팅 정책

@property {Object<string, SheetMeta>} [allSheetsData] - 시트별 전처리 데이터

@property {ColumnRange} [bestReturn] - 자동 탐색된 반환 열

@property {ColumnRange} [bestLookup] - 자동 탐색된 조회 열

@property {Object} formulaBuilder - 빌더 함수 모음
*/

/* ---------------------------------------------
 * Controller-level 기본 정책/옵션
 * -------------------------------------------*/
const DEFAULT_ENGINE = "excel";
const DEFAULT_POLICY = {
  mode: "loose",
  value_if_not_found: "",
  value_if_error: "",
};
const DEFAULT_FORMAT_OPTIONS = {
  case_sensitive: false,
  trim_text: true,
  coerce_number: true,
};

// function summarizeIntentForCache(intent = {}) {
//   const op = (intent.operation || "").toLowerCase();

//   const headerHints = [
//     intent.header_hint,
//     intent.return_hint,
//     intent.lookup_hint,
//   ]
//     .filter(Boolean)
//     .map((s) => String(s).trim().toLowerCase())
//     .sort()
//     .join("|");

//   // 조건(target + operator 위주로 요약, 값은 캐시에 크게 중요하지 않다고 가정)
//   const conds = (intent.conditions || []).map((c) => {
//     if (c && typeof c === "object" && c.logical_operator) {
//       return `G:${c.logical_operator}`;
//     }
//     const t =
//       typeof c?.target === "string"
//         ? c.target.toLowerCase()
//         : c?.target?.header?.toLowerCase?.() || "";
//     const op = c?.operator || "=";
//     return `C:${t}:${op}`;
//   });

//   const windowSig = intent.window
//     ? [
//         intent.window.type || "",
//         intent.window.size || "",
//         (intent.window.date_header || "").toLowerCase(),
//       ].join(":")
//     : "";

//   return [op, headerHints, conds.sort().join("&"), windowSig].join("||");
// }

// function buildCacheKey({ message, fileHash, allSheetsData, intent }) {
//   const normalizedMessage = String(message || "")
//     .trim()
//     .toLowerCase()
//     .replace(/\s+/g, " ");

//   const headers = [];
//   if (allSheetsData) {
//     Object.values(allSheetsData).forEach((sheetInfo) => {
//       Object.keys(sheetInfo.metaData || {}).forEach((h) => headers.push(h));
//     });
//   }
//   headers.sort();

//   const intentSig = summarizeIntentForCache(intent || {});

//   return [
//     normalizedMessage,
//     intentSig || "no-intent",
//     fileHash || "nofile",
//     headers.join("|") || "noheaders",
//   ].join("||");
// }

function shouldUseDirectBuilder(intent = {}, ctx = {}) {
  if (typeof direct?.isDirectEligible === "function") {
    return direct.isDirectEligible(intent, ctx);
  }

  const raw = String(intent?.raw_message || "");
  const explicit =
    formulaUtils.parseExplicitCellOrRange(raw) ||
    intent?.range ||
    intent?.target_cell;

  const hasSheetMeta =
    !!ctx?.allSheetsData && Object.keys(ctx.allSheetsData).length > 0;

  const headerDriven = Boolean(
    intent?.header_hint ||
    intent?.return_hint ||
    intent?.lookup_hint ||
    intent?.group_by ||
    (Array.isArray(intent?.return_fields) && intent.return_fields.length) ||
    (Array.isArray(intent?.filters) && intent.filters.length) ||
    (Array.isArray(intent?.conditions) && intent.conditions.length) ||
    intent?.lookup?.key_header,
  );

  if (hasSheetMeta) return false;
  if (headerDriven) return false;
  if (!explicit) return false;

  return true;
}

/* ---------------------------------------------
 * 로컬 의도 추론 (LLM 미사용 시 폴백)
 * -------------------------------------------*/
function _deduceOp(text = "") {
  const s = String(text).toLowerCase();
  if (/(average|avg|mean|평균)/.test(s)) return "average";
  if (/(sum|total|합계|총합|합\b)/.test(s)) return "sum";
  if (/(count|개수|갯수|건수|수량|카운트)/.test(s)) return "count";
  if (/(xlookup|lookup|찾아|조회|검색|참조)/.test(s)) return "xlookup";
  if (/(filter|필터)/.test(s)) return "filter";
  if (/\b(if|조건|만약)\b/.test(s)) return "if";
  if (/(median|중앙값|중간값|가운데\s*값|중앙\s*(연봉|급여|값|금액)?)/.test(s))
    return "median";
  if (/(stdev|표준편차)/.test(s)) return "stdev_s";
  if (/(var|분산)/.test(s)) return "var_s";
  if (/(sortby|정렬)/.test(s)) return "sortby";
  return "formula";
}

function _extractDeptConditions(message = "") {
  const out = [];
  const m = String(message || "");

  const deptNames = ["영업", "마케팅", "개발", "인사", "재무", "총무"];
  for (const d of deptNames) {
    if (new RegExp(`${d}\\s*부서|${d}`).test(m)) {
      out.push(d);
    }
  }

  return [...new Set(out)];
}

function _extractGradeCondition(message = "") {
  const m = String(message || "");
  const hit = m.match(/평가\s*등급\s*([ABCDFS][\+\-]?)/i);
  if (!hit) return null;
  return hit[1].toUpperCase();
}

function _extractSalaryThreshold(message = "") {
  const m = String(message || "");
  const hit = m.match(
    /(연봉|급여)\s*([0-9]+(?:[.,][0-9]+)?)\s*(이상|이하|초과|미만)?/,
  );
  if (!hit) return null;

  const raw = hit[2].replace(/,/g, "");
  const dir = hit[3] || "이상";

  let operator = ">=";
  if (/이하/.test(dir)) operator = "<=";
  else if (/초과/.test(dir)) operator = ">";
  else if (/미만/.test(dir)) operator = "<";

  return { value: raw, operator };
}

const CANONICAL_FIELD_ALIASES = [
  {
    canonical: "직원 ID",
    aliases: ["직원 id", "직원ID", "사번", "직원번호", "id"],
  },
  { canonical: "이름", aliases: ["이름", "성명", "name"] },
  { canonical: "부서", aliases: ["부서", "소속", "팀"] },
  { canonical: "직급", aliases: ["직급", "직함", "직책", "title"] },
  { canonical: "연봉", aliases: ["연봉", "급여", "salary"] },
  {
    canonical: "입사일",
    aliases: ["입사일", "등록일", "채용일", "입사 날짜", "등록 날짜"],
  },
  { canonical: "평가 등급", aliases: ["평가 등급", "평가", "등급", "rating"] },
];

function _escapeRegExp(s = "") {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function _findCanonicalFieldsInText(text = "") {
  const src = String(text || "");
  const found = [];

  for (const spec of CANONICAL_FIELD_ALIASES) {
    const hit = spec.aliases.some((alias) =>
      new RegExp(
        `(^|[^가-힣A-Za-z0-9])${_escapeRegExp(alias)}([^가-힣A-Za-z0-9]|$)`,
        "i",
      ).test(src),
    );
    if (hit) found.push(spec.canonical);
  }

  return [...new Set(found)];
}

function _extractReturnFieldsFromMessage(text = "") {
  const found = _findCanonicalFieldsInText(text);
  return found.filter((x) => x !== "직원 ID");
}

function _extractLookupFieldFromMessage(text = "") {
  const src = String(text || "");

  const byPatterns = [
    /(직원\s*id|직원ID|사번|직원번호|id)\s*(으로|로|기준으로|기준)/i,
    /(이름|성명|name)\s*(으로|로|기준으로|기준)/i,
    /(부서|소속|팀)\s*(으로|로|기준으로|기준)/i,
    /(직급|직함|직책)\s*(으로|로|기준으로|기준)/i,
    /(연봉|급여|salary)\s*(으로|로|기준으로|기준)/i,
  ];

  for (const p of byPatterns) {
    const m = src.match(p);
    if (!m) continue;
    const found = _findCanonicalFieldsInText(m[1]);
    if (found.length) return found[0];
  }

  const equalityPatterns = [
    /(직원\s*id|직원ID|사번|직원번호|id)\s*가\s*/i,
    /(이름|성명|name)\s*이\s*/i,
    /(부서|소속|팀)\s*가\s*/i,
    /(직급|직함|직책)\s*이\s*/i,
    /(연봉|급여|salary)\s*이\s*/i,
  ];

  for (const p of equalityPatterns) {
    const m = src.match(p);
    if (!m) continue;
    const found = _findCanonicalFieldsInText(m[1]);
    if (found.length) return found[0];
  }

  return null;
}

function _extractLookupValueFromMessage(text = "") {
  const src = String(text || "").trim();

  // 1) 명시적 셀 참조
  const cellRef = src.match(/\b([A-Z]{1,3}\d{1,7})\b/);
  if (cellRef) return cellRef[1].toUpperCase();

  // 2) 따옴표로 감싼 명시 값
  const quoted = src.match(/["']([^"']+)["']/);
  if (quoted) {
    const v = quoted[1].trim();
    return v || null;
  }

  // 3) "직원 ID가 1001", "이름이 홍길동" 같은 equality 구조만 허용
  const equality = src.match(
    /(직원\s*id|직원ID|사번|직원번호|id|이름|성명|name|부서|소속|팀|직급|직함|직책|연봉|급여)\s*(?:가|이)\s*([A-Za-z0-9가-힣._-]+)/i,
  );

  if (equality) {
    const raw = String(equality[2] || "").trim();

    // 동사/어미 잘못 캡처 방지
    if (/^(가져와줘|보여줘|찾아줘|출력해줘|조회해줘|반환해줘)$/i.test(raw)) {
      return null;
    }

    // 너무 짧은 꼬리값 방지
    if (raw.length <= 1) return null;

    return raw;
  }

  // 4) "으로/로" 구조는 lookup field 힌트일 뿐, value가 아님
  //    예: "사번으로 이름 가져와줘" → value 없음
  return null;
}

function _looksLikeStructuredLookup(text = "") {
  const src = String(text || "");

  if (/(찾아|조회|검색|lookup|xlookup|참조)/i.test(src)) return true;
  if (
    /(가져와|보여줘|출력|반환)/i.test(src) &&
    /(으로|로|기준으로|기준|가|이)/.test(src)
  )
    return true;

  return false;
}

/* ---------------------------------------------
 * buildLocalIntentFromText(text)
 * -------------------------------------------
 * 역할:
 *  - LLM(OpenAI)이 없을 때 fallback으로 동작하는
 *    규칙 기반 Intent 추론기.
 *
 * 기능:
 *  1) operation 키워드 감지 (sum, count, lookup, filter, if 등)
 *  2) 간단한 열 힌트/조건 힌트 추출
 *  3) 최소 Intent 스키마 구조 반환
 *
 * 결과:
 *  - LLM이 없을 때도 어느 정도 작동 가능한 기본 Intent 객체 생성
 * -------------------------------------------*/
function buildLocalIntentFromText(text = "") {
  const original = String(text || "");
  const s = original.toLowerCase().trim();

  const op = _deduceOp(s);

  /** @type {Intent} */
  const intent = { operation: op };

  // ✅ B(중앙값) 우선 해결:
  // "중앙값" 요청인데 header_hint가 비면 bestReturn이 연봉이 아닌 숫자열로 잡힐 수 있음.
  // 로컬 intent 경로에서는 확실히 연봉을 타게 유도한다.
  if (intent.operation === "median") {
    if (!intent.header_hint && !intent.return_hint) {
      intent.header_hint = "연봉";
    }
  }

  // ✅ 1. Lookup / 조회 구조 감지
  // - "사번으로 이름 가져와줘"
  // - "직원 ID로 부서 보여줘"
  // - "직원 ID가 1001인 이름 찾아줘"
  // - "홍길동의 연봉"
  const lookupMatch = s.match(
    /([가-힣a-z0-9]+)[의\s]*(매출|점수|금액|이름|값|수량|가격|부서|직급|연봉|급여)/i,
  );

  if (op.includes("lookup") || _looksLikeStructuredLookup(original)) {
    intent.operation = "xlookup";

    const detectedLookup = _extractLookupFieldFromMessage(original);
    const detectedReturns = _extractReturnFieldsFromMessage(original);
    const detectedValue = _extractLookupValueFromMessage(original);

    if (lookupMatch) {
      const lhsFields = _findCanonicalFieldsInText(lookupMatch[1]);
      const rhsFields = _findCanonicalFieldsInText(lookupMatch[2]);

      if (lhsFields.length && !intent.lookup_hint) {
        intent.lookup_hint = lhsFields[0];
      }
      if (rhsFields.length) {
        if (rhsFields.length >= 2) {
          intent.return_fields = [...new Set(rhsFields)];
          delete intent.return_hint;
        } else if (!intent.return_hint) {
          intent.return_hint = rhsFields[0];
        }
      }
    }

    if (!intent.lookup_hint && detectedLookup) {
      intent.lookup_hint = detectedLookup;
    }

    const returnCandidates = detectedReturns.filter(
      (x) => x && x !== intent.lookup_hint,
    );

    if (returnCandidates.length >= 2) {
      intent.return_fields = [...new Set(returnCandidates)];
      delete intent.return_hint;
    } else if (
      returnCandidates.length === 1 &&
      !intent.return_hint &&
      !(Array.isArray(intent.return_fields) && intent.return_fields.length)
    ) {
      intent.return_hint = returnCandidates[0];
    }

    if (detectedValue != null && detectedValue !== "") {
      intent.lookup_value = detectedValue;
    } else {
      delete intent.lookup_value;
    }

    // fallback: 이름만 언급됐는데 lookup도 return도 비어 있으면
    if (!intent.lookup_hint && /(이름|성명|name)/i.test(original)) {
      intent.lookup_hint = "이름";
    }

    // latest duplicate 보강
    if (
      /(중복된\s*경우|중복일\s*경우|가장\s*최근|최근\s*입사|최신|latest)/i.test(
        original,
      )
    ) {
      intent.duplicate_rule = "latest";
      if (!intent.date_header) intent.date_header = "입사일";
    }

    // not-found fallback 보강
    if (
      /(존재하지\s*않는|없는)/.test(original) &&
      intent.value_if_not_found == null
    ) {
      intent.value_if_not_found = "";
    }

    // lookup과 return이 같은 값이면 return 제거
    if (
      intent.lookup_hint &&
      intent.return_hint &&
      String(intent.lookup_hint).trim() === String(intent.return_hint).trim()
    ) {
      delete intent.return_hint;
    }

    if (Array.isArray(intent.return_fields) && intent.lookup_hint) {
      intent.return_fields = intent.return_fields.filter(
        (x) => String(x).trim() !== String(intent.lookup_hint).trim(),
      );
      if (!intent.return_fields.length) delete intent.return_fields;
    }

    // 값이 없는 구조형 lookup 문장은 formula 생성을 서두르지 않도록 표시
    if (
      intent.operation === "xlookup" &&
      !intent.lookup_value &&
      !intent.lookup?.value &&
      !intent.lookup?.value_ref
    ) {
      intent.needs_lookup_value = true;
    }

    return intent;
  }

  // ✅ 텍스트 공백 제거 패턴
  if (
    /(공백|띄어쓰기|앞뒤\s*공백|trim)/i.test(original) &&
    /(이름|성명)/.test(original)
  ) {
    intent.operation = "trim";
    intent.header_hint = "이름";
    intent.target_header = "이름";
    intent.scope = "all";
    return intent;
  }

  // ✅ 2. 집계 / 그룹 집계 패턴
  // 예: "부서별 직원 수", "평가 등급별 평균 연봉", "부서별 평가 등급 A 직원 수"
  if (
    /(sum|합계|total|평균|average|count|개수|갯수|인원수|직원\s*수|최고|최저|중앙값|중간값|중앙\s*(연봉|급여|값|금액)?|정렬|순으로)/.test(
      s,
    )
  ) {
    const groupBy = _detectGroupByFromMessage(original);
    const aggOp = _detectAggregateOpFromMessage(original, intent.operation);
    const headerHint = _detectHeaderHintFromMessage(original);
    const sortOrder = _detectSortOrderFromMessage(original);

    if (groupBy) intent.group_by = groupBy;
    intent.operation = aggOp;

    if (!intent.header_hint && !intent.return_hint) {
      if (aggOp !== "count" && headerHint) {
        intent.header_hint = headerHint;
      }
    }

    const conditions = [];
    const conditionGroups = [];

    const grade = _extractGradeCondition(original);
    if (grade) {
      conditions.push({
        target: "평가 등급",
        operator: "=",
        value: grade,
      });
    }

    const salaryCond = _extractSalaryThreshold(original);
    if (salaryCond) {
      conditions.push({
        target: "연봉",
        operator: salaryCond.operator,
        value: salaryCond.value,
        value_type: "number",
      });
    }

    const depts = _extractDeptConditions(original);

    // "영업 또는 마케팅" → OR 그룹
    if (/(또는|or)/i.test(original) && depts.length >= 2) {
      conditionGroups.push({
        logical_operator: "OR",
        conditions: depts.map((d) => ({
          target: "부서",
          operator: "=",
          value: d,
        })),
      });
    } else if (depts.length === 1) {
      conditions.push({
        target: "부서",
        operator: "=",
        value: depts[0],
      });
    }

    if (conditions.length) {
      intent.conditions = conditions;
    }
    if (conditionGroups.length) {
      intent.condition_groups = conditionGroups;
    }

    if (sortOrder) {
      intent.sorted = true;
      intent.sort_order = sortOrder;
    }

    return intent;
  }

  // ✅ 3. 필터 조건 패턴
  // 예: "매출이 100만원 이상인 행", "이름이 홍길동인 데이터"
  if (/(filter|필터|조건|만족|해당)/.test(s)) {
    intent.operation = "filter";
    const cond = {};
    const headerMatch = s.match(/(매출|금액|점수|나이|기간|날짜)/);
    if (headerMatch) cond.target = headerMatch[1];
    if (/(이상|greater|over|>)\b/.test(s)) cond.operator = ">=";
    else if (/(이하|under|<|작은)\b/.test(s)) cond.operator = "<=";
    else if (/(같|=|equal)/.test(s)) cond.operator = "=";
    const numMatch = s.match(/([0-9]+[.,]?[0-9]*)/);
    if (numMatch) cond.value = numMatch[1];
    if (Object.keys(cond).length) intent.conditions = [cond];
    return intent;
  }

  // ✅ 3-1. 평균 이상/이하 IF 패턴
  if (
    /(연봉|급여)/.test(original) &&
    /(평균\s*(이상|이하|초과|미만)|average)/i.test(original) &&
    /이면|아니면/.test(original)
  ) {
    intent.operation = "if";
    intent.scope = "all";

    let op = ">=";
    if (/평균\s*이하/.test(original)) op = "<=";
    else if (/평균\s*초과/.test(original)) op = ">";
    else if (/평균\s*미만/.test(original)) op = "<";

    intent.condition = {
      target: { header: "연봉" },
      operator: op,
      value: {
        operation: "average",
        header_hint: "연봉",
      },
    };

    // 따옴표가 없어도 "상/하" 추출
    if (/상/.test(original)) intent.value_if_true = "상";
    if (/하/.test(original)) intent.value_if_false = "하";

    // 혹시 다른 표현도 허용
    if (!intent.value_if_true) intent.value_if_true = "상";
    if (!intent.value_if_false) intent.value_if_false = "하";

    return intent;
  }

  // ✅ 평균 이상/이하 IF 패턴
  if (
    /(연봉|급여)/.test(original) &&
    /(평균\s*(이상|이하|초과|미만)|average)/i.test(original) &&
    /이면|아니면/.test(original)
  ) {
    intent.operation = "if";
    intent.scope = "all";

    let op = ">=";
    if (/평균\s*이하/.test(original)) op = "<=";
    else if (/평균\s*초과/.test(original)) op = ">";
    else if (/평균\s*미만/.test(original)) op = "<";

    intent.condition = {
      target: { header: "연봉" },
      operator: op,
      value: {
        operation: "average",
        header_hint: "연봉",
      },
    };

    intent.value_if_true = "상";
    intent.value_if_false = "하";
    return intent;
  }

  // ✅ 4. IF 조건형 문장
  // 예: "매출이 100 이상이면 '우수', 아니면 '보통'"
  if (/\bif\b|조건|이면|아니면|참|거짓/.test(s)) {
    intent.operation = "if";
    const cond = {};
    const headerMatch = s.match(/(매출|점수|나이|금액|수량)/);
    if (headerMatch) cond.target = headerMatch[1];
    if (/(이상|greater|over|>)\b/.test(s)) cond.operator = ">=";
    else if (/(이하|under|<|작은)\b/.test(s)) cond.operator = "<=";
    const numMatch = s.match(/([0-9]+[.,]?[0-9]*)/);
    if (numMatch) cond.value = numMatch[1];
    const labelMatches = s.match(/['"](.*?)['"]/g);
    if (labelMatches && labelMatches.length >= 2) {
      intent.value_if_true = labelMatches[0].replace(/['"]/g, "");
      intent.value_if_false = labelMatches[1].replace(/['"]/g, "");
    }
    intent.condition = cond;
    return intent;
  }

  // ✅ 5. 날짜/최근 기간 패턴
  // 예: "최근 7일 매출", "지난달 평균 매출"
  if (/최근|지난|이번|오늘|yesterday|today|month|week|day/.test(s)) {
    const numMatch = s.match(
      /([0-9]+)\s*(일|day|days|주|week|weeks|달|month|months)/,
    );
    const size = numMatch ? parseInt(numMatch[1], 10) : 7;
    intent.window = { type: "days", size, date_header: "날짜" };
    if (/매출|sales/.test(s)) intent.header_hint = "매출액";
    if (/평균/.test(s)) intent.operation = "averageifs";
    else intent.operation = "sumifs";
    return intent;
  }

  // ✅ 6. 기본 fallback
  return intent;
}

/* ---------------------------------------------
 * normalizeLookupIntent(intent)
 * -------------------------------------------
 * 역할:
 *  - LLM 또는 로컬 룰 기반 Intent 중
 *    lookup / xlookup 계열의 입력 필드를 표준 구조로 보정한다.
 *
 * 표준화 내용:
 *  1) LLM이 준 lookup_key / return 필드를 lookup_array / return_array로 변환
 *  2) lookup_value, lookup_array, return_array를 보장
 *  3) referenceFunctions 등 빌더가 기대하는 intent.lookup / intent.return 구조를 생성
 *
 * 결과:
 *  - 모든 xlookup Intent는 아래 필드를 최소 포함하게 된다.
 *      intent.lookup_value
 *      intent.lookup_array = { sheet, header }
 *      intent.return_array = { sheet, header }
 *      intent.lookup = { value, header, sheet }
 *      intent.return = { header, sheet }
 * -------------------------------------------*/
function normalizeLookupIntent(intent) {
  if (!intent || !intent.operation) return intent;

  const op = String(intent.operation).toLowerCase();
  if (op !== "xlookup" && op !== "lookup") return intent;

  // ✅ 1. LLM 출력 보정: lookup_key → lookup_array 변환
  if (intent.lookup_key) {
    if (intent.lookup_value == null)
      intent.lookup_value = intent.lookup_key.value;
    if (!intent.lookup_array) {
      intent.lookup_array = {
        sheet: intent.lookup_key.sheet,
        header: intent.lookup_key.header,
      };
    }
  }

  // ✅ 2. return → return_array 변환
  if (intent.return && !intent.return_array) {
    intent.return_array = {
      sheet: intent.return.sheet,
      header: intent.return.header,
    };
  }

  // ✅ 3. 중첩 구조 통일 (referenceFunctions 호환용)
  intent.lookup = intent.lookup || {};
  if (intent.lookup_value != null && intent.lookup.value == null) {
    intent.lookup.value = intent.lookup_value;
  }
  if (intent.lookup_array) {
    if (!intent.lookup.header)
      intent.lookup.header = intent.lookup_array.header;
    if (!intent.lookup.sheet) intent.lookup.sheet = intent.lookup_array.sheet;
  }

  intent.return = intent.return || {};
  if (intent.return_array) {
    if (!intent.return.header)
      intent.return.header = intent.return_array.header;
    if (!intent.return.sheet) intent.return.sheet = intent.return_array.sheet;
  }

  if (Array.isArray(intent.return_fields) && intent.return_fields.length) {
    intent.return_fields = [...new Set(intent.return_fields.map(String))];
  }

  return intent;
}

function buildCtx(rawCtx) {
  const rawIntent = rawCtx.intent || {};
  const message = rawCtx.message || rawIntent.raw_message || "";

  const intent = normalizeIntentSchema(rawIntent, message);
  const { engine, policy, formatOptions } = normalizePolicy(intent);

  const baseCtx = {
    ...rawCtx,
    intent,
    engine,
    policy,
    formatOptions,
  };

  const resolved = resolveIntent(baseCtx);
  return buildResolvedContext(baseCtx, resolved);
}

/* ---------------------------------------------
 * formulaBuilder 어셈블
 *  - 빌더들이 기대하는 헬퍼만 노출
 * -------------------------------------------*/
const formulaBuilder = {
  _formatValue: (val, opts = {}) => formulaUtils.formatValue(val, { ...opts }), // 정책은 상위에서 주입

  _buildConditionPairs: function (ctx) {
    const { intent, allSheetsData } = ctx;
    if (!allSheetsData) return [];
    if (!intent?.conditions?.length) return [];

    return intent.conditions
      .map((c) => {
        // 1) 어떤 문자열을 헤더 후보로 쓸지 정리
        let headerText = "";

        if (typeof c?.target === "string") {
          headerText = c.target;
        } else if (c?.target && typeof c.target === "object") {
          // HeaderSpec 형태 { header, sheet, ... } 지원
          headerText = c.target.header || "";
        } else if (c?.hint) {
          // 혹시 과거 포맷과의 호환을 위해 hint도 fallback으로 사용
          headerText = c.hint;
        }

        if (!headerText) return null;

        const term = formulaUtils.expandTermsFromText(headerText);
        const best = formulaUtils.findBestColumnAcrossSheets(
          allSheetsData,
          term,
          "lookup",
        );
        if (!best) return null;

        // ✅ 불확실(Top2 gap 좁음)하면 "그럴듯하게 틀림" 방지를 위해 즉시 중단
        if (best.isAmbiguous) {
          const candA = best.header || "후보1";
          const candB = best.runnerUpHeader || "후보2";
          ctx.__errorFormula = `=ERROR("조건 열이 모호합니다: '${candA}' 또는 '${candB}' 중 선택이 필요합니다.")`;
          return null;
        }

        const range = `'${best.sheetName}'!${best.columnLetter}${best.startRow}:${best.columnLetter}${best.lastDataRow}`;

        const op = String(c.operator || "=").trim();
        const rawVal = c.value;

        // ✅ 값이 비어있으면 조건으로 만들지 않는다 ("" 조건 생성 방지)
        if (
          rawVal == null ||
          (typeof rawVal === "string" && rawVal.trim() === "")
        ) {
          return null;
        }

        // 값도 반드시 포매터를 통과시켜 따옴표/숫자 처리
        const val = formulaBuilder._formatValue(rawVal);

        // COUNTIFS/SUMIFS/AVERAGEIFS 기준:
        // - 숫자 비교:  "<=100" 형태
        // - 날짜/텍스트 비교(>=,<= 등): "<="&DATEVALUE("2023-01-01") 처럼 연결
        // - contains/starts_with/ends_with: 와일드카드
        const cmpOps = new Set([">", ">=", "<", "<=", "<>"]);
        if (cmpOps.has(op)) {
          if (rawVal != null && !isNaN(rawVal))
            return `${range}, "${op}${rawVal}"`;
          return `${range}, "${op}"&${val}`;
        }
        if (/^contains$/i.test(op)) return `${range}, "*"&${val}&"*"`;
        if (/^starts?_with$/i.test(op)) return `${range}, ${val}&"*"`;
        if (/^ends?_with$/i.test(op)) return `${range}, "*"&${val}`;

        // 기본(=)
        return `${range}, ${val}`;
      })
      .filter(Boolean);
  },

  _buildConditionMask: function (ctx) {
    return buildConditionMask(ctx, (v, o) =>
      formulaUtils.formatValue(v, {
        ...(ctx?.formatOptions || {}),
        ...(o || {}),
      }),
    );
  },
};
Object.assign(formulaBuilder, logicalFunctionBuilder);
Object.assign(formulaBuilder, mathStatsFunctionBuilder);
Object.assign(formulaBuilder, dateFunctionBuilder);
Object.assign(formulaBuilder, referenceFunctionBuilder);
Object.assign(formulaBuilder, textFunctionBuilder);
Object.assign(formulaBuilder, arrayFunctionBuilder);

/* ---------------------------------------------
 * OP 해석 (별칭 → 실제 구현 함수키)
 * convert()에서만 사용. handleConversion은 직접 키 호출.
 * -------------------------------------------*/
const OP_ALIASES = {
  if: "if",
  ifs: "ifs",
  textjoin: "textjoin",
  text_join: "textjoin",

  xlookup: "xlookup",
  lookup: "xlookup",

  // 합계
  sum: "sum",
  sumifs: "sum",

  // 평균
  average: "average",
  avg: "average",
  averageifs: "average",

  // 개수
  count: "count",
  countifs: "count",

  // 통계
  stdev: "stdev_s",
  var: "var_s",

  median: "median",
  med: "median",

  // ✅ 행 반환(최고/최저 직원 정보)
  maxrow: "maxrow",
  minrow: "minrow",
  argmax: "maxrow",
  argmin: "minrow",
  top1: "maxrow",
  bottom1: "minrow",
  topnrows: "topnrows",
  rankcolumn: "rankcolumn",

  sortby: "sortby",
  regexmatch: "regexmatch",
  textsplit: "textsplit",
};

// ✅ LLM이 평균으로 오판하는 케이스를 중앙값에 한해 안전하게 보정
function applyMedianOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;
  if (!/(median|중앙값|중간값|가운데\s*값)/i.test(msg)) return intent;

  const op = String(intent.operation || "").toLowerCase();
  if (!op || ["average", "avg", "mean"].includes(op)) {
    intent.operation = "median";
  }
  if (!intent.header_hint && !intent.return_hint) {
    intent.header_hint = "연봉";
  }
  return intent;
}

function applyDateBoundaryOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;
  if (!Array.isArray(intent.conditions)) return intent;

  const yearAfterMatch = msg.match(/(20\d{2})년\s*이후/);
  if (!yearAfterMatch) return intent;

  const y = yearAfterMatch[1];
  intent.conditions = intent.conditions.map((c) => {
    if (!c || typeof c !== "object") return c;
    if (!/(입사일|날짜)/.test(String(c.target || ""))) return c;
    return {
      ...c,
      operator: ">=",
      value: `${y}-01-01`,
    };
  });
  return intent;
}

function applyExtremeRowOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  // ✅ Top N / N명 / Top3 / 상위 3명 문장은 extreme-row(1건)가 아니라
  //    뒤 단계의 topnrows가 처리하도록 여기서는 건드리지 않는다.
  const hasExplicitTopN =
    /\btop\s*\d+/i.test(msg) ||
    /(상위|하위)\s*[2-9]\d*\s*명?/.test(msg) ||
    /\d+\s*명(?:의|을|만|씩|중)?/.test(msg);

  if (hasExplicitTopN) return intent;

  // ✅ 1) 날짜 extreme-row: 최근 입사 / 가장 오래 근무
  const wantsDateRow =
    /(이름|성명|직원|정보)/.test(msg) && /(입사|입사일|근무)/.test(msg);

  if (wantsDateRow) {
    const isRecentHire =
      /(가장\s*최근|최근\s*입사|최신|most\s*recent|latest)/i.test(msg) &&
      /(입사|입사일)/i.test(msg);

    const isOldestTenure =
      /(가장\s*오래|오래\s*근무|최장\s*근무|earliest|oldest)/i.test(msg) &&
      /(근무|입사|입사일)/i.test(msg);

    if (isRecentHire) {
      intent.operation = "maxrow";
      intent.header_hint = "입사일";
      if (!intent.return_headers && !intent.select_headers) {
        intent.return_headers = ["이름"];
      }
      return intent;
    }

    if (isOldestTenure) {
      intent.operation = "minrow";
      intent.header_hint = "입사일";
      if (!intent.return_headers && !intent.select_headers) {
        intent.return_headers = ["이름"];
      }
      return intent;
    }
  }

  const wantsRowFields =
    /(이름|성명|부서|직급|정보|직원)/.test(msg) && /(연봉|salary)/i.test(msg);
  if (!wantsRowFields) return intent;

  const isMax =
    /(가장\s*높|최고|최대|top|highest|max)/i.test(msg) &&
    !/(가장\s*낮|최저|최소|bottom|lowest|min)/i.test(msg);
  const isMin = /(가장\s*낮|최저|최소|bottom|lowest|min)/i.test(msg);

  if (isMax) intent.operation = "maxrow";
  else if (isMin) intent.operation = "minrow";
  else return intent;

  if (!intent.header_hint && !intent.return_hint) intent.header_hint = "연봉";
  if (!intent.return_headers && !intent.select_headers) {
    intent.return_headers = ["이름", "부서", "직급", "연봉"];
  }
  return intent;
}

function applyRecentTopNOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  const nMatch =
    msg.match(/\btop\s*(\d+)\b/i) ||
    msg.match(/상위\s*(\d+)\s*명?/) ||
    msg.match(/하위\s*(\d+)\s*명?/) ||
    msg.match(/(\d+)\s*명/);

  const takeN = nMatch ? Number(nMatch[1]) : null;
  if (!Number.isFinite(takeN) || takeN <= 0) return intent;

  const wantsList = /(보여줘|목록|리스트|명|행|직원|사람)/.test(msg);
  const hasRankingCue =
    /(top|상위|하위|높은\s*순|낮은\s*순|가장\s*높|가장\s*낮|최고|최저|가장\s*빠른|가장\s*늦은|가장\s*최근|최근\s*순|최신\s*순|오래된\s*순|오래\s*된\s*순|내림차순|오름차순)/i.test(
      msg,
    );

  if (!(wantsList && hasRankingCue)) return intent;

  const currentOp = String(intent.operation || "").toLowerCase();
  const explicitTopNSignal =
    /\btop\s*\d+\b/i.test(msg) ||
    /(상위|하위)\s*\d+\s*명?/.test(msg) ||
    /\d+\s*명(?:의|을|만|씩|중)?/.test(msg);

  // monthcount/yearcount는 유지
  // maxrow/minrow는 "명시적 Top N" 문장이라면 topnrows가 덮어쓰게 허용
  if (
    ["monthcount", "yearcount"].includes(currentOp) ||
    ((currentOp === "maxrow" || currentOp === "minrow") && !explicitTopNSignal)
  ) {
    return intent;
  }

  const headerHint = _detectHeaderHintFromMessage(msg);
  const isHireDateTopN = /(입사|입사일|근무)/.test(msg);
  const isSalaryTopN = /(연봉|salary)/i.test(msg);

  const resolvedHeader =
    (isSalaryTopN && "연봉") || (isHireDateTopN && "입사일") || headerHint;

  if (!resolvedHeader) return intent;

  let sortOrder = _detectSortOrderFromMessage(msg);
  if (!sortOrder) {
    if (/(하위|낮은\s*순|가장\s*빠른|오래된\s*순|오래\s*된\s*순)/i.test(msg)) {
      sortOrder = "asc";
    } else {
      sortOrder = "desc";
    }
  }

  intent.operation = "topnrows";
  intent.header_hint = resolvedHeader;
  intent.sort_order = sortOrder;
  intent.take_n = takeN;

  const deptMatch = msg.match(/([가-힣A-Za-z0-9]+)\s*부서/);
  if (deptMatch) {
    _appendCondition(intent, {
      target: "부서",
      operator: "=",
      value: deptMatch[1],
    });
  }

  const gradeMatch = msg.match(/평가\s*등급\s*([ABCDFS][\+\-]?)/i);
  if (gradeMatch) {
    _appendCondition(intent, {
      target: "평가 등급",
      operator: "=",
      value: gradeMatch[1].toUpperCase(),
    });
  }

  // ✅ Top N 문장은 기존 return_headers를 그대로 믿지 말고 재정규화
  const headers = [];

  const explicitNameOnly =
    /(이름만|성명만)/.test(msg) ||
    (/(이름|성명)/.test(msg) &&
      !/(이름\s*(과|와|,|및|하고)\s*(연봉|부서|직급|입사일|직원\s*id|사번|id))/.test(
        msg,
      ) &&
      !/((연봉|부서|직급|입사일|직원\s*id|사번|id)\s*(과|와|,|및|하고)\s*(이름|성명))/.test(
        msg,
      ) &&
      !/(포함|같이|함께)/.test(msg));

  const wantsId = /직원\s*id|사번|id/i.test(msg);
  const wantsName =
    /(이름|성명)/.test(msg) || /직원\s*\d+\s*명|직원|사람/.test(msg);

  // 정렬 기준으로 등장한 "연봉"은 반환열에 자동 포함하지 않는다.
  const wantsSalaryField =
    /(이름|성명|직원\s*id|사번|id|부서|직급|입사일)\s*(과|와|,|및|하고)\s*연봉/i.test(
      msg,
    ) ||
    /연봉\s*(과|와|,|및|하고)\s*(이름|성명|직원\s*id|사번|id|부서|직급|입사일)/i.test(
      msg,
    ) ||
    /(연봉\s*포함|연봉도|연봉까지|이름과\s*연봉|이름\s*,\s*연봉)/i.test(msg);

  const wantsHireDateField =
    /(이름\s*과\s*입사일|입사일\s*과\s*이름|입사일\s*포함)/.test(msg);
  const wantsDeptField = /(이름\s*과\s*부서|부서\s*와\s*이름|부서\s*포함)/.test(
    msg,
  );
  const wantsTitleField =
    /(이름\s*과\s*직급|직급\s*과\s*이름|직급\s*포함)/.test(msg);

  if (explicitNameOnly) {
    intent.return_headers = ["이름"];
  } else {
    if (wantsId) headers.push("직원ID");
    if (wantsName || !wantsId) headers.push("이름");
    if (wantsDeptField) headers.push("부서");
    if (wantsTitleField) headers.push("직급");
    if (wantsSalaryField) headers.push("연봉");
    if (wantsHireDateField) headers.push("입사일");

    intent.return_headers = [...new Set(headers.length ? headers : ["이름"])];
  }

  delete intent.select_headers;

  return intent;
}

function applyMonthCountOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  const wantsMonthCount =
    /(월별|월\s*단위)/.test(msg) &&
    /(입사|입사일)/.test(msg) &&
    /(직원\s*수|입사자\s*수|개수|인원수|표)/.test(msg);

  if (!wantsMonthCount) return intent;

  intent.operation = "monthcount";
  intent.header_hint = "입사일";
  intent.return_hint = "입사일";
  return intent;
}

function applyYearCountOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  const wantsYearCount =
    /(연도별|년도별|연\s*단위)/.test(msg) &&
    /(입사|입사일)/.test(msg) &&
    /(직원\s*수|입사자\s*수|개수|인원수|표)/.test(msg);

  if (!wantsYearCount) return intent;

  intent.operation = "yearcount";
  intent.header_hint = "입사일";
  intent.return_hint = "입사일";
  return intent;
}

function applyUniqueSortOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  const wantsUnique = /(중복\s*없이|중복\s*제거|unique)/i.test(msg);
  if (!wantsUnique) return intent;

  intent.operation = "unique";

  // "가나다순" 등 정렬 키워드가 있으면 SORT(UNIQUE())로 가도록 플래그
  if (/(가나다|정렬|오름차순|asc)/i.test(msg)) {
    intent.sorted = true;
    intent.sort_order = "asc";
  }

  // bestReturn이 null이면 UNIQUE 빌더가 =ERROR(...) 반환하므로 최소 힌트 보강
  if (!intent.header_hint && !intent.return_hint) {
    if (/부서/.test(msg)) intent.header_hint = "부서";
  }
  return intent;
}

function applySortListOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  const hasSortCue =
    /(높은\s*순|낮은\s*순|내림차순|오름차순|정렬|순으로)/i.test(msg);
  const hasSalaryCue = /(연봉|salary)/i.test(msg);
  const hasListCue = /(보여줘|목록|리스트|직원|사람|이름|성명)/i.test(msg);

  if (!(hasSortCue && hasSalaryCue && hasListCue)) return intent;

  const currentOp = String(intent.operation || "").toLowerCase();
  if (
    ["topnrows", "maxrow", "minrow", "rankcolumn", "unique"].includes(currentOp)
  ) {
    return intent;
  }

  intent.operation = "sortby";
  intent.header_hint = "연봉";
  intent.lookup_hint = "연봉";
  intent.sort_by = "연봉";
  intent.sort_order = /(낮은\s*순|오름차순|작은\s*순)/i.test(msg)
    ? "asc"
    : "desc";

  const deptMatch = msg.match(/([가-힣A-Za-z0-9]+)\s*부서/);
  if (deptMatch) {
    _appendCondition(intent, {
      target: "부서",
      operator: "=",
      value: deptMatch[1],
    });
  }

  const gteMatch = msg.match(/연봉\s*(\d+(?:\.\d+)?)\s*(이상|초과)/);
  if (gteMatch) {
    _appendCondition(intent, {
      target: "연봉",
      operator: gteMatch[2] === "초과" ? ">" : ">=",
      value: Number(gteMatch[1]),
    });
  }

  const lteMatch = msg.match(/연봉\s*(\d+(?:\.\d+)?)\s*(이하|미만)/);
  if (lteMatch) {
    _appendCondition(intent, {
      target: "연봉",
      operator: lteMatch[2] === "미만" ? "<" : "<=",
      value: Number(lteMatch[1]),
    });
  }

  const explicitNameOnly =
    /(이름만|성명만)/.test(msg) ||
    (/(이름|성명)/.test(msg) &&
      !/(이름\s*(과|와|,|및|하고)\s*(연봉|부서|직급|입사일|직원\s*id|사번|id))/.test(
        msg,
      ) &&
      !/((연봉|부서|직급|입사일|직원\s*id|사번|id)\s*(과|와|,|및|하고)\s*(이름|성명))/.test(
        msg,
      ) &&
      !/(포함|같이|함께)/.test(msg));

  const explicitNameSalary =
    /(이름|성명).*(연봉)|(연봉).*(이름|성명)/i.test(msg) ||
    /(연봉\s*포함|연봉도|연봉까지)/.test(msg);

  if (explicitNameOnly) {
    intent.return_headers = ["이름"];
  } else if (explicitNameSalary) {
    intent.return_headers = ["이름", "연봉"];
  } else {
    intent.return_headers = ["이름"];
  }

  if (intent.return_headers.length === 1) {
    intent.return_hint = intent.return_headers[0];
  } else {
    delete intent.return_hint;
  }
  delete intent.select_headers;

  return intent;
}

function applyFilteredSortOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  const currentOp = String(intent.operation || "").toLowerCase();
  if (currentOp === "sortby") return intent;
  if (
    ["topnrows", "maxrow", "minrow", "rankcolumn", "unique"].includes(currentOp)
  ) {
    return intent;
  }

  const hasSortCue = /(높은\s*순|낮은\s*순|내림차순|오름차순|정렬|순으로)/.test(
    msg,
  );
  const sortHeader = _detectHeaderHintFromMessage(msg);
  const hasEmployeeListCue = /(직원|사람|행|목록|리스트|보여줘)/.test(msg);
  const hasFilterCue = /(부서|평가\s*등급|이상|이하|초과|미만)/.test(msg);
  if (!(hasSortCue && sortHeader && hasEmployeeListCue && hasFilterCue))
    return intent;

  intent.operation = "sortby";
  intent.lookup_hint = sortHeader;
  intent.header_hint = sortHeader;
  intent.sort_order = /(낮은\s*순|오름차순|작은\s*순)/i.test(msg)
    ? "asc"
    : "desc";

  const deptMatch = msg.match(/([가-힣A-Za-z0-9]+)\s*부서/);
  if (deptMatch) {
    _appendCondition(intent, {
      target: "부서",
      operator: "=",
      value: deptMatch[1],
    });
  }
  const gradeMatch = msg.match(/평가\s*등급\s*([ABCDFS][\+\-]?)/i);
  if (gradeMatch) {
    _appendCondition(intent, {
      target: "평가 등급",
      operator: "=",
      value: gradeMatch[1].toUpperCase(),
    });
  }
  const salaryThreshold = msg.match(
    /연봉\s*([0-9][0-9,]*)\s*(이상|이하|초과|미만)/,
  );
  if (salaryThreshold) {
    const opMap = { 이상: ">=", 이하: "<=", 초과: ">", 미만: "<" };
    _appendCondition(intent, {
      target: "연봉",
      operator: opMap[salaryThreshold[2]] || ">=",
      value: String(salaryThreshold[1]).replace(/,/g, ""),
    });
  }

  const explicitNameOnly = /(이름만|성명만|이름\s*만\s*보여)/.test(msg);
  const explicitNameSalary =
    /(이름\s*과\s*연봉|연봉\s*과\s*이름|이름\s*와\s*연봉|연봉\s*와\s*이름|연봉\s*포함)/.test(
      msg,
    );
  if (explicitNameOnly) intent.return_headers = ["이름"];
  else if (explicitNameSalary) intent.return_headers = ["이름", "연봉"];
  else intent.return_headers = ["이름", "부서", "직급", sortHeader];

  return intent;
}

function applyRankColumnOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;
  if (!/(순위\s*열|순위열|등수\s*열|랭크\s*열|순위\s*만들)/.test(msg))
    return intent;

  const headerHint =
    _detectHeaderHintFromMessage(msg) || intent.header_hint || "연봉";
  intent.operation = "rankcolumn";
  intent.header_hint = headerHint;
  intent.return_hint = headerHint;
  intent.sort_order = /(낮은\s*순|오름차순|작은\s*값)/i.test(msg)
    ? "asc"
    : "desc";
  return intent;
}

function applyDuplicateLatestMetricOverride(message, intent) {
  const msg = String(message || "");

  if (!/중복/.test(msg)) return intent;
  if (!/가장\s*최근|최신/.test(msg)) return intent;
  if (!/(연봉|급여|이름|부서|직급)/.test(msg)) return intent;
  if (!/(직원\s*id|사번|직원번호)/i.test(msg)) return intent;

  intent.operation = "duplicate_latest_metric";
  intent.lookup_hint = "직원 ID";
  intent.date_header = "입사일";

  if (/(연봉|급여)/.test(msg)) intent.return_hint = "연봉";
  else if (/이름/.test(msg)) intent.return_hint = "이름";
  else if (/부서/.test(msg)) intent.return_hint = "부서";
  else if (/직급/.test(msg)) intent.return_hint = "직급";

  return intent;
}

function applyRankThresholdCountOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  if (!/(직급)/.test(msg)) return intent;
  if (!/(개수|갯수|인원수|직원\s*수|count)/.test(msg)) return intent;
  if (!/(이상)/.test(msg)) return intent;

  const rankOrder = [
    "사원",
    "주임",
    "대리",
    "과장",
    "차장",
    "부장",
    "이사",
    "상무",
    "전무",
    "대표",
  ];
  const hit = rankOrder.find((r) => msg.includes(r));
  if (!hit) return intent;

  const idx = rankOrder.indexOf(hit);
  if (idx < 0) return intent;

  const allowed = rankOrder.slice(idx);

  intent.operation = "count";
  intent.condition_groups = [
    {
      logical_operator: "OR",
      conditions: allowed.map((r) => ({
        target: "직급",
        operator: "=",
        value: r,
      })),
    },
  ];

  // 중복 방지: 단순 conditions는 제거
  delete intent.conditions;
  return intent;
}

function _detectGroupByFromMessage(msg = "") {
  const s = String(msg || "");
  if (/부서별|부서\s*기준|각\s*부서/.test(s)) return "부서";
  if (/직급별|직급\s*기준|각\s*직급/.test(s)) return "직급";
  if (/평가\s*등급별|등급별|평가별/.test(s)) return "평가 등급";
  return null;
}

function _detectHeaderHintFromMessage(msg = "") {
  const s = String(msg || "");
  if (/(연봉|salary)/i.test(s)) return "연봉";
  if (/(입사일|입사\s*날짜)/.test(s)) return "입사일";
  if (/(평가\s*등급|등급)/.test(s)) return "평가 등급";
  if (/직급/.test(s)) return "직급";
  if (/부서/.test(s)) return "부서";
  return null;
}

function _detectAggregateOpFromMessage(msg = "", fallbackOp = "") {
  const s = String(msg || "");
  if (/(평균|average|avg|mean)/i.test(s)) return "average";
  if (/(합계|총합|sum|total)/i.test(s)) return "sum";
  if (/(개수|갯수|건수|인원수|직원\s*수|count)/i.test(s)) return "count";
  if (/(최고|최대|가장\s*높|max|highest)/i.test(s)) return "max";
  if (/(최저|최소|가장\s*낮|min|lowest)/i.test(s)) return "min";
  if (/(중앙값|중간값|가운데\s*값|median|중앙\s*(연봉|급여|값|금액)?)/i.test(s))
    return "median";
  return fallbackOp || "formula";
}

function _detectSortOrderFromMessage(msg = "") {
  const s = String(msg || "");
  if (/(적은\s*순|낮은\s*순|오름차순|asc|작은\s*순)/i.test(s)) return "asc";
  if (/(많은\s*순|높은\s*순|내림차순|desc|큰\s*순)/i.test(s)) return "desc";
  return null;
}

function _appendCondition(intent, cond) {
  if (!intent || !cond || typeof cond !== "object") return;
  if (!intent.conditions) intent.conditions = [];

  const incomingTarget = String(cond.target || cond.header || "")
    .trim()
    .toLowerCase();
  const incomingOp = String(cond.operator || "=")
    .trim()
    .toLowerCase();
  const incomingValue = String(cond.value ?? "")
    .trim()
    .toLowerCase();

  const exists = intent.conditions.some((c) => {
    if (!c || typeof c !== "object" || c.logical_operator) return false;
    const target = String(c.target || c.header || "")
      .trim()
      .toLowerCase();
    const op = String(c.operator || "=")
      .trim()
      .toLowerCase();
    const value = String(c.value ?? "")
      .trim()
      .toLowerCase();
    return (
      target === incomingTarget && op === incomingOp && value === incomingValue
    );
  });

  if (!exists) {
    intent.conditions.push(cond);
  }
}

function _extractDeptsFromMessage(msg = "") {
  const s = String(msg || "");
  const candidates = ["영업", "마케팅", "개발", "인사", "재무", "총무"];
  return candidates.filter((d) => new RegExp(`${d}\\s*부서|${d}`).test(s));
}

function _extractGradeFromMessage(msg = "") {
  const m = String(msg || "").match(/평가\s*등급\s*([ABCDFS][\+\-]?)/i);
  return m ? m[1].toUpperCase() : null;
}

function _extractNumericThresholdFromMessage(msg = "") {
  const m = String(msg || "").match(
    /(연봉|급여)\s*([0-9][0-9,]*)\s*(이상|이하|초과|미만)/,
  );
  if (!m) return null;

  const opMap = { 이상: ">=", 이하: "<=", 초과: ">", 미만: "<" };
  return {
    target: /급여/.test(m[1]) ? "연봉" : "연봉",
    operator: opMap[m[3]] || ">=",
    value: m[2].replace(/,/g, ""),
    value_type: "number",
  };
}

function applyTrimColumnOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  if (!/(공백|띄어쓰기|앞뒤\s*공백|trim)/i.test(msg)) return intent;
  if (!/(이름|성명)/.test(msg)) return intent;

  intent.operation = "trim";
  intent.header_hint = "이름";
  intent.target_header = "이름";
  intent.scope = "all";

  delete intent.delimiter;
  delete intent.delimiters;
  delete intent.values;
  delete intent.a;
  delete intent.b;

  return intent;
}

function applyAverageThresholdIfOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  if (!/(연봉|급여)/.test(msg)) return intent;
  if (!/(평균\s*(이상|이하|초과|미만)|average)/i.test(msg)) return intent;
  if (!/이면|아니면|출력/.test(msg)) return intent;

  let op = ">=";
  if (/평균\s*이하/.test(msg)) op = "<=";
  else if (/평균\s*초과/.test(msg)) op = ">";
  else if (/평균\s*미만/.test(msg)) op = "<";

  intent.operation = "if";
  intent.scope = "all";
  intent.condition = {
    target: { header: "연봉" },
    operator: op,
    value: {
      operation: "average",
      header_hint: "연봉",
    },
  };

  intent.value_if_true = "상";
  intent.value_if_false = "하";

  return intent;
}

function applyGroupedAggregateOverride(message, intent) {
  const msg = String(message || "");
  if (!intent || typeof intent !== "object") return intent;

  const groupBy = _detectGroupByFromMessage(msg);
  const aggOp = _detectAggregateOpFromMessage(msg, intent.operation);
  const headerHint = _detectHeaderHintFromMessage(msg);
  const sortOrder = _detectSortOrderFromMessage(msg);

  const looksGrouped =
    !!groupBy &&
    /(개수|갯수|인원수|직원\s*수|평균|합계|총합|최고|최저|중앙값|중간값|중앙\s*(연봉|급여|값|금액)?|정렬|순으로|많은\s*순|높은\s*순|낮은\s*순)/.test(
      msg,
    );

  if (!looksGrouped) return intent;

  // 1) 핵심 집계 구조
  intent.operation = aggOp;
  if (groupBy) {
    intent.group_by = groupBy;
    if (
      intent.group_by &&
      typeof intent.group_by === "object" &&
      intent.group_by.header
    ) {
      intent.group_by = intent.group_by.header;
    }
  }

  // 2) 집계 대상 열 보정
  // count는 header_hint가 없어도 되지만,
  // average/max/min/median/sum은 연봉류가 없으면 흔들릴 수 있어 기본 보정
  if (!intent.header_hint && !intent.return_hint) {
    if (aggOp !== "count") {
      intent.header_hint = /(연봉|salary)/i.test(msg)
        ? "연봉"
        : headerHint || "연봉";
    }
  }

  // 3) "평가 등급 A 직원 수" 같은 조건 추론
  const gradeMatch = msg.match(/평가\s*등급\s*([ABCDFS][\+\-]?)/i);
  if (gradeMatch) {
    _appendCondition(intent, {
      target: "평가 등급",
      operator: "=",
      value: gradeMatch[1].toUpperCase(),
    });
  }
  const depts = _extractDeptsFromMessage(msg);
  if (/(또는|or)/i.test(msg) && depts.length >= 2) {
    intent.condition_groups = intent.condition_groups || [];

    const existsSameOrGroup = intent.condition_groups.some((g) => {
      if (!g || String(g.logical_operator || "").toUpperCase() !== "OR")
        return false;
      const vals = (g.conditions || []).map((c) => String(c.value || ""));
      return depts.every((d) => vals.includes(d));
    });

    if (!existsSameOrGroup) {
      intent.condition_groups.push({
        logical_operator: "OR",
        conditions: depts.map((d) => ({
          target: "부서",
          operator: "=",
          value: d,
        })),
      });
    }
  } else if (depts.length === 1) {
    _appendCondition(intent, {
      target: "부서",
      operator: "=",
      value: depts[0],
    });
  }

  const numericThreshold = _extractNumericThresholdFromMessage(msg);
  if (numericThreshold) {
    _appendCondition(intent, numericThreshold);
  } else {
    const bareGradeMatch = msg.match(
      /(?:^|\s)([ABCDFS][\+\-]?)(?:\s*등급|\s*직원|\s*인원|\s*$)/i,
    );
    if (bareGradeMatch && /등급|직원|인원/.test(msg)) {
      _appendCondition(intent, {
        target: "평가 등급",
        operator: "=",
        value: bareGradeMatch[1].toUpperCase(),
      });
    }
  }

  // 4) 정렬 보정
  if (sortOrder) {
    intent.sorted = true;
    intent.sort_order = sortOrder;
  }

  // 5) count 계열은 sortby가 아니라 count 집계로 되돌린다
  if (
    String(intent.operation || "").toLowerCase() === "sortby" &&
    /(개수|갯수|인원수|직원\s*수)/.test(msg)
  ) {
    intent.operation = "count";
  }

  return intent;
}

function resolveOp(op) {
  if (!op) return null;
  const k = String(op).toLowerCase().replace(/[ \-]/g, "");
  const base = OP_ALIASES[k] || k;
  return typeof formulaBuilder[base] === "function" ? base : null;
}

/* ---------------------------------------------
 * 파일 전처리 유틸
 * -------------------------------------------*/
async function loadAndPreprocessFromBucketIfPossible(user, fileName) {
  const logLP = shouldLogCache();
  if (logLP) console.log("[loadAndPreprocess] user?.id:", user?.id);
  if (logLP) console.log("[loadAndPreprocess] fileName:", fileName);

  const bucket = getBucket();
  if (logLP) console.log("[loadAndPreprocess] bucket exists?:", !!bucket);

  if (!bucket || !user || !fileName) {
    if (logLP)
      console.log("[loadAndPreprocess] early return (no bucket/user/fileName)");
    return { isFileAttached: false, preprocessed: null };
  }

  if (logLP)
    console.log("[loadAndPreprocess] user.uploadedFiles:", user?.uploadedFiles);
  const fileInfo = user.uploadedFiles?.find((f) => f.originalName === fileName);
  if (logLP) console.log("[loadAndPreprocess] fileInfo:", fileInfo);

  if (!fileInfo) {
    if (logLP)
      console.log("[loadAndPreprocess] early return (fileInfo not found)");
    return { isFileAttached: false, preprocessed: null };
  }

  const file = bucket.file(fileInfo.gcsName);
  const [buffer] = await file.download();
  if (logLP)
    console.log("[loadAndPreprocess] downloaded buffer length:", buffer.length);

  const { fileHash, allSheetsData } = await getOrBuildAllSheetsData(buffer);
  if (logLP) {
    console.log(
      "[loadAndPreprocess] got allSheetsData keys:",
      Object.keys(allSheetsData || {}),
    );
  }

  return {
    isFileAttached: true,
    preprocessed: { fileHash, allSheetsData },
  };
}

/* ---------------------------------------------
 * LLM 의도 추출 (OpenAI 있을 때만)
 *  - formulaUtils에 의존하지 않도록 컨트롤러에 포함
 * -------------------------------------------*/
function getSystemPrompt() {
  return `
You are an intent extractor for a formula generator (Excel / Google Sheets).

Your job:
  - Read the user's request (in Korean or English).
  - Return ONLY ONE JSON object named "intent".
  - This JSON will later be used to build an Excel/Sheets formula.
  - DO NOT output any formula. JSON ONLY.

Core fields (always consider):
  - operation: string
      The main action. Examples:
        "sum", "sumifs", "average", "averageifs", "countifs",
        "lookup", "xlookup", "filter", "if", "sortby",
        "textjoin", "textsplit", "regexmatch", "regexreplace".
      Choose the most appropriate single operation.

  - engine (optional): "excel" | "sheets"
      Only set this if the user explicitly mentions the target (e.g. Google Sheets).

  - mode (optional): "strict" | "normal"
      Use "strict" only if the user explicitly asks for exact / case-sensitive behavior.

Lookup-related hints (when the user wants to find values by key):
  - lookup_hint (optional): string
      Natural language description of the key column.
      e.g. "고객 ID", "이메일", "상품코드"
  - return_hint (optional): string
      Natural language description of the value to return.
      e.g. "매출액", "고객 이름", "재고 수량"
  - header_hint (optional): string
      General target column when not a typical lookup.
  - lookup_value (optional):
      Value or concept used to look up (e.g. a specific customer name).

Conditions (for sumifs / averageifs / countifs / filter / if):
  - conditions (optional): array of condition nodes.
      A condition node can be:
        { "target": "매출액", "operator": ">", "value": 1000000 }
      or
        {
          "logical_operator": "AND" | "OR",
          "conditions": [ ... nested condition nodes ... ]
        }
  - condition (optional): single condition node (for "if" style operations).

Date window (recent N days, weeks, etc.):
  - window (optional):
      {
        "type": "days",
        "size": 7,
        "date_header": "날짜"
      }

Row selection (select a specific row by key):
  - row_selector (optional):
      {
        "hint": "고객 ID",
        "value": 12345,
        "sheet": "Sheet1"
      }

Aggregation / grouping (sum by branch, average by category, etc.):
  - group_by (optional): string
      e.g. "지점명", "카테고리"

Text operations:
  - delimiter (optional): string
      For textjoin or simple splitting.
  - delimiters (optional): string[]
      For multiple split delimiters.
  - ignore_empty (optional): boolean
  - remove_empty_text (optional): boolean

IF / mapping:
  - value_if_true (optional)
  - value_if_false (optional)
  - in_values (optional): array of values, for "IN" style checks.
  - message (optional): user-facing message or error text if needed.

General rules:
  - Prefer header-based hints (lookup_hint, return_hint, header_hint, group_by)
    instead of hard-coded cell references or ranges.
  - Use Korean header names if the user uses Korean column names.
  - If you are unsure about a field, OMIT it rather than guessing.
  - The output MUST be valid JSON. No comments, no trailing commas, no extra text.
  - Prefer header-based hints (lookup_hint, return_hint, header_hint, group_by)
    instead of hard-coded cell references or ranges.
  - Do NOT invent sheet names, column letters, or A1-style ranges.
    Never output things like "Sheet1!B2:B100" or column letters like "A", "B", "C".
  - Use Korean header names if the user uses Korean column names.
  - If you are unsure about a field, OMIT it rather than guessing.
  - The output MUST be valid JSON. No comments, no trailing commas, no extra text.

  Examples (important):

1) Simple SUM with condition
User: "서울 지점의 매출 합계를 구해줘"
Intent:
{
  "intent": {
    "operation": "sum",
    "header_hint": "매출액",
    "conditions": [
      { "target": "지점", "operator": "=", "value": "서울" }
    ]
  }
}

2) XLOOKUP style
User: "김선수의 포지션을 찾아줘"
Intent:
{
  "intent": {
    "operation": "lookup",
    "lookup_hint": "선수명",
    "return_hint": "포지션",
    "lookup_value": "김선수"
  }
}

3) AVERAGE with recent N days window
User: "최근 7일간 매출 평균"
Intent:
{
  "intent": {
    "operation": "average",
    "header_hint": "매출액",
    "window": {
      "type": "days",
      "size": 7,
      "date_header": "날짜"
    }
  }
}

Only follow the JSON structure shown above. For each new user request, return exactly one JSON object named "intent".
`.trim();
}

function buildFewShotBlock(fewShots = []) {
  const good = (fewShots || []).filter(
    (fs) => fs && fs.isHelpful !== false && fs.intent && fs.message,
  );

  // 최근 5개 정도만 사용
  const selected = good.slice(-5);

  if (!selected.length) {
    return "No additional labeled examples are available.";
  }

  return selected
    .map((ex, idx) => {
      return [
        `Example ${idx + 1}:`,
        `User: "${ex.message}"`,
        `Intent JSON:`,
        JSON.stringify(ex.intent, null, 2),
      ].join("\n");
    })
    .join("\n\n");
}

async function extractIntentWithLLM(
  openai,
  message,
  metaHintForModel,
  fewShots = [],
) {
  const fewShotText = buildFewShotBlock(fewShots);

  const userPrompt =
    `You are given some past labeled examples (user message + intent JSON).\n` +
    `Use them as guidance for the style and structure of the "intent" you should produce.\n\n` +
    `=== Labeled Examples ===\n` +
    `${fewShotText}\n\n` +
    `=== Current Task ===\n` +
    `Analyze the user request and return a structured JSON intent.\n` +
    `Data Schema: ${metaHintForModel}\n` +
    `User Request: "${message}"`;

  const completion = await openai.chat.completions.create({
    model: "gpt-4o-mini",
    temperature: 0,
    response_format: { type: "json_object" },
    messages: [
      { role: "system", content: getSystemPrompt() },
      { role: "user", content: userPrompt },
    ],
  });
  const raw = completion?.choices?.[0]?.message?.content || "{}";
  const cleaned = cleanAIResponse(raw);
  try {
    return JSON.parse(cleaned);
  } catch {
    return {};
  }
}

function shouldCountConversion(result) {
  if (typeof result !== "string") return false;
  const t = result.trim();
  if (!t) return false;

  // ❌ 에러 수식은 카운트 제외
  if (/^=ERROR\s*\(/i.test(t)) return false;

  // ✅ 정상 Excel/Sheets 수식
  if (t.startsWith("=")) return true;

  // ✅ SQL 등 텍스트 결과도 허용하고 싶으면 (현재 프론트 isFormula 기준과 맞춤)
  if (/^(SELECT|WITH)\b/i.test(t)) return true;

  // ✅ Notion/기타 텍스트 포맷(현재 프론트에서 prop( 포함이면 코드블록 처리)
  if (t.includes("prop(")) return true;

  return false;
}

/* ---------------------------------------------
 * 메인 컨버전 핸들러
 * -------------------------------------------*/
exports.handleConversion = async (req, res, next) => {
  // ===== DEBUG STATE =====
  let _dbgCompatibility = null;
  let _dbgCtx = null;
  // ---- debug-safe holders (so logging never crashes) ----
  let _dbgMessage = null;
  let _dbgIntent = null;
  let _dbgIntentCacheKey = null;
  let _dbgCacheHit = null;

  // ---- timing (ms) ----
  const _t0 = process.hrtime.bigint();
  let _tPreStart = null,
    _tPreEnd = null;
  let _tIntentStart = null,
    _tIntentEnd = null;
  let _tBuildStart = null,
    _tBuildEnd = null;

  function _ms(a, b) {
    if (!a || !b) return null;
    return Number(b - a) / 1e6;
  }

  function _shouldLogTiming() {
    // Dev: always
    if (process.env.NODE_ENV !== "production") return true;
    // Prod: sample (default 1%)
    const rate = Number(process.env.CONVERT_TIMING_LOG_RATE || "0.01");
    if (!(rate > 0)) return false;
    return Math.random() < rate;
  }

  const traceId = crypto.randomUUID
    ? crypto.randomUUID()
    : crypto.randomBytes(16).toString("hex");
  const startedAt = Date.now();

  try {
    const {
      message,
      fileName,
      conversionType = "Excel/Google Sheets",
    } = req.body || {};
    _dbgMessage = message || null;
    if (!message || !conversionType) {
      return res.status(400).json({ result: "요청 정보가 부족합니다." });
    }

    // ✅ 변환 한도 체크 (FREE면 10회)
    if (req.user?.id) {
      try {
        await assertCanUse(req.user.id, "formulaConversions", 1);
      } catch (e) {
        return res.status(e.status || 429).json({
          error: "Usage limit exceeded",
          code: e.code || "LIMIT_EXCEEDED",
          ...e.meta,
        });
      }
    }

    // 1) 파일 전처리(옵션)
    _tPreStart = process.hrtime.bigint();
    const { isFileAttached, preprocessed } =
      await loadAndPreprocessFromBucketIfPossible(req.user, fileName);
    _tPreEnd = process.hrtime.bigint();

    const fileHash = preprocessed?.fileHash || null;
    const allSheetsData = preprocessed?.allSheetsData || null;
    const sheetStateSig = makeSheetStateSig(allSheetsData);

    // 2) 메타 힌트(LLM용)
    let metaHintForModel = "No file data provided.";
    if (isFileAttached && allSheetsData) {
      const allHeaders = new Set();
      Object.values(allSheetsData).forEach((sheetInfo) => {
        Object.keys(sheetInfo.metaData || {}).forEach((h) =>
          allHeaders.add(`'${h}'`),
        );
      });
      metaHintForModel = `The file contains columns like: [${Array.from(
        allHeaders,
      ).join(", ")}]`;
    }

    // 3) 의도 추출 (OpenAI 있으면 LLM, 없으면 로컬)
    let intent = buildLocalIntentFromText(message);
    intent = normalizeIntentSchema(intent, message);

    const openai = getOpenAI();

    _tIntentStart = process.hrtime.bigint();

    // ---------------------------------------------
    // Intent Cache (SKELETON) - default OFF
    // ---------------------------------------------
    // NOTE:
    // - Cache stores INTENT ONLY (never formula/script).
    // - Key must include context to avoid cross-file leakage.
    //
    // Enable later by setting: INTENT_CACHE_ENABLED=1
    // ---------------------------------------------
    const modelName = "gpt-4o-mini";
    const intentSchemaVersion = "intent-v1";
    const builderType = conversionType || "Excel/Google Sheets";
    const userKey = req.user?.id ? `u:${req.user.id}` : `anon:${req.ip}`;
    const targetRangeSig = null; // TODO: wire from UI when available (e.g. selected column/range)

    let cacheHit = false;
    if (intentCache.isEnabled()) {
      const { key } = buildIntentCacheKey({
        version: 1,
        builderType,
        model: modelName,
        schemaVersion: intentSchemaVersion,
        userKey,
        prompt: message,
        sheetStateSig: `${fileHash || "nofile"}|${sheetStateSig}`,
        targetRangeSig,
      });
      _dbgIntentCacheKey = key;

      const cached = await intentCache.get(key);

      if (cached && cached.intent && typeof cached.intent === "object") {
        _dbgCacheHit = true;
        if (shouldLogCache()) {
          console.log("[intentCache] HIT", key.slice(0, 8));
        }
        intent = { ...intent, ...cached.intent };
      } else {
        _dbgCacheHit = false;
        if (shouldLogCache()) {
          console.log("[intentCache] MISS", key.slice(0, 8));
        }
      }
    }

    // ✅ LLM 호출 (single place)
    const skipLLMOnHit = process.env.INTENT_CACHE_SKIP_LLM_ON_HIT === "1";
    if (openai && !(skipLLMOnHit && _dbgCacheHit === true)) {
      const llm = await extractIntentWithLLM(
        openai,
        message,
        metaHintForModel,
        [], // fewShots disabled (no persistence)
      );

      if (llm && typeof llm === "object") {
        if (llm.intent && typeof llm.intent === "object") {
          intent = { ...intent, ...llm.intent };
        } else {
          intent = { ...intent, ...llm };
        }
      }
    }

    intent = normalizeIntentSchema(intent, message);
    intent = normalizeLookupIntent(intent);
    intent = applyMedianOverride(message, intent);
    intent = applyDateBoundaryOverride(message, intent);
    intent = applyExtremeRowOverride(message, intent);
    intent = applyRecentTopNOverride(message, intent);
    intent = applyMonthCountOverride(message, intent);
    intent = applyYearCountOverride(message, intent);
    intent = applyUniqueSortOverride(message, intent);
    intent = applySortListOverride(message, intent);
    intent = applyFilteredSortOverride(message, intent);
    intent = applyRankColumnOverride(message, intent);
    intent = applyGroupedAggregateOverride(message, intent);
    intent = applyDuplicateLatestMetricOverride(message, intent);
    intent = applyRankThresholdCountOverride(message, intent);
    intent = applyTrimColumnOverride(message, intent);
    intent = applyAverageThresholdIfOverride(message, intent);
    intent.raw_message = message;
    _dbgIntent = intent;

    _tIntentEnd = process.hrtime.bigint();

    // store intent only (SKELETON)
    if (intentCache.isEnabled() && _dbgIntentCacheKey) {
      await intentCache.set(
        _dbgIntentCacheKey,
        {
          intent,
          meta: {
            model: modelName,
            schema: intentSchemaVersion,
          },
        },
        600, // 10 min TTL (tune later)
      );
    }

    // 4) 컨텍스트 구성 + 자동 열 매핑
    _tBuildStart = process.hrtime.bigint();
    let context = buildCtx({
      intent,
      message,
      formulaBuilder,
      allSheetsData,
    });
    _dbgCtx = context;
    if (isFileAttached && allSheetsData) {
      const hasHints = !!(
        intent.return_hint ||
        intent.header_hint ||
        intent.lookup_hint
      );

      if (hasHints) {
        const searchTerms = {
          return: intent.return_hint || intent.header_hint || "",
          lookup: intent.lookup_hint || "",
        };

        const joint = formulaUtils.findBestSheetAndColumns(
          allSheetsData,
          searchTerms,
          {
            sameSheetBonus: 0.5,
          },
        );

        const bestReturn = joint.return;
        const bestLookup = joint.lookup;

        if (!bestReturn && (intent.header_hint || intent.return_hint)) {
          return res.json({
            result: `=ERROR("필요한 열을 파일에서 찾을 수 없습니다.")`,
          });
        }

        Object.assign(context, {
          bestReturn,
          bestLookup,
          allSheetsData,
        });
        _dbgCtx = context;
      } else {
        Object.assign(context, { allSheetsData });
      }
    }

    // 5) direct(파일無) 빠른 경로
    if (
      !isFileAttached &&
      direct?.canHandleWithoutFile?.(intent, context) &&
      shouldUseDirectBuilder(intent, context)
    ) {
      const f = direct.buildFormula(intent, context);
      if (f) {
        // ✅ 6-1: 출력 검증(Direct도 동일 적용)
        const v = validateFormula(f);
        const safeOut = v.ok
          ? f
          : `=ERROR("결과 검증에 실패했습니다. (direct) 다시 시도해 주세요.")`;
        const directCompatibility = detectFormulaCompatibility(safeOut || "");
        _dbgCompatibility = directCompatibility;
        if (req.user?.id && shouldCountConversion(f)) {
          await bumpUsage(req.user.id, "formulaConversions", 1);
        }
        const rawReason = "OK";
        const reasonNorm = classifyReason({
          reason: rawReason,
          prompt: message,
          result: safeOut,
        });
        await writeRequestLog({
          traceId,
          userId: req.user?.id,
          route: "/convert",
          engine: "formula",
          status: "success",
          reason: reasonNorm,
          isFallback: false,
          prompt: message,
          latencyMs: Date.now() - startedAt,
          debugMeta: buildDebugMeta({
            rawReason,
            cacheHit: _dbgCacheHit,
            intentOp: intent?.operation,
            intentCacheKey: _dbgIntentCacheKey,
            validator: v,
            timing: {
              preprocess: _ms(_tPreStart, _tPreEnd),
              intent: _ms(_tIntentStart, _tIntentEnd),
              build: _ms(_tBuildStart, _tBuildEnd),
              total: Date.now() - startedAt,
            },
            extra: {
              compatibility: directCompatibility,
              resolvedBaseSheet: context?.resolved?.baseSheet || null,
              resolvedReturnHeaders: (
                context?.resolved?.returnColumns || []
              ).map((x) => x.header),
              resolvedLookupHeader:
                context?.resolved?.lookupColumn?.header || null,
              resolvedGroupHeader:
                context?.resolved?.groupColumn?.header || null,
            },
          }),
        });
        return res.json({
          result: safeOut,
          compatibility: directCompatibility,
        });
      }
    }

    // 6) 빌더 호출
    const opKey = resolveOp(intent.operation);
    const builder = opKey && formulaBuilder[opKey];

    let finalFormula;
    if (!builder) {
      finalFormula = `=ERROR("지원하지 않는 작업입니다: ${
        intent.operation || "none"
      }")`;
    } else {
      finalFormula = builder.call(
        formulaBuilder,
        context,
        formulaBuilder._formatValue,
        formulaBuilder._buildConditionPairs,
        formulaBuilder._buildConditionMask,
      );
    }
    _tBuildEnd = process.hrtime.bigint();

    if (req.user?.id && shouldCountConversion(finalFormula)) {
      await bumpUsage(req.user.id, "formulaConversions", 1);
    }
    const rawReason = shouldCountConversion(finalFormula)
      ? "OK"
      : String(finalFormula || "").startsWith("=ERROR(")
        ? "ERROR_FORMULA"
        : "UNKNOWN";

    const reasonNorm = classifyReason({
      reason: rawReason,
      prompt: message,
      result: finalFormula,
    });

    // ✅ 6-1: 최종 출력 검증(깨진 수식/따옴표/괄호 불일치 차단)
    const v = validateFormula(finalFormula);
    const safeFinal = v.ok
      ? finalFormula
      : `=ERROR("결과 검증에 실패했습니다. 입력을 더 구체적으로 작성해 주세요.")`;

    const compatibility = detectFormulaCompatibility(safeFinal || "");
    _dbgCompatibility = compatibility;

    await writeRequestLog({
      traceId,
      userId: req.user?.id,
      route: "/convert",
      engine: "formula",
      status: shouldCountConversion(safeFinal) ? "success" : "fail",
      reason: reasonNorm,
      isFallback: v.ok ? false : true,
      prompt: message,
      latencyMs: Date.now() - startedAt,
      debugMeta: buildDebugMeta({
        rawReason,
        cacheHit: _dbgCacheHit,
        intentOp: intent?.operation,
        intentCacheKey: _dbgIntentCacheKey,
        validator: v,
        timing: {
          preprocess: _ms(_tPreStart, _tPreEnd),
          intent: _ms(_tIntentStart, _tIntentEnd),
          build: _ms(_tBuildStart, _tBuildEnd),
          total: Date.now() - startedAt,
        },
        extra: {
          compatibility,
          resolvedBaseSheet: context?.resolved?.baseSheet || null,
          resolvedReturnHeaders: (context?.resolved?.returnColumns || []).map(
            (x) => x.header,
          ),
          resolvedLookupHeader: context?.resolved?.lookupColumn?.header || null,
          resolvedGroupHeader: context?.resolved?.groupColumn?.header || null,
        },
      }),
    });
    const finalCompatibility = detectFormulaCompatibility(
      safeFinal || finalFormula || "",
    );

    return res.json({
      result: safeFinal,
      compatibility: finalCompatibility,
    });
  } catch (err) {
    const rawReason = "EXCEPTION";
    const reasonNorm = classifyReason({
      reason: rawReason,
      prompt: _dbgMessage,
      error: err,
    });
    await writeRequestLog({
      traceId,
      userId: req.user?.id,
      route: "/convert",
      engine: "formula",
      status: "fail",
      reason: reasonNorm,
      isFallback: false,
      prompt: _dbgMessage || "",
      latencyMs: Date.now() - startedAt,
      debugMeta: buildDebugMeta({
        rawReason,
        cacheHit: _dbgCacheHit,
        intentOp: _dbgIntent?.operation,
        intentCacheKey: _dbgIntentCacheKey,
        validator: null,
        timing: {
          preprocess: _ms(_tPreStart, _tPreEnd),
          intent: _ms(_tIntentStart, _tIntentEnd),
          build: _ms(_tBuildStart, _tBuildEnd),
          total: Date.now() - startedAt,
        },
        extra: {
          error: err?.message,
          stack: err?.stack?.slice?.(0, 500),
          compatibility: _dbgCompatibility || null,
          resolvedBaseSheet: _dbgCtx?.resolved?.baseSheet || null,
        },
      }),
    });
    console.error("[handleConversion][error]", err);
    next(err);
  } finally {
    if (_shouldLogTiming()) {
      const tTotal = _ms(_t0, process.hrtime.bigint());
      const tPre = _ms(_tPreStart, _tPreEnd);
      const tIntent = _ms(_tIntentStart, _tIntentEnd);
      const tBuild = _ms(_tBuildStart, _tBuildEnd);
      const user = req.user?.id ? String(req.user.id) : "anon";
      const file =
        req.body && req.body.fileName ? String(req.body.fileName) : "-";
      const cache =
        _dbgCacheHit === true ? "HIT" : _dbgCacheHit === false ? "MISS" : "NA";
      console.log(
        `[convert.timing] total=${tTotal?.toFixed?.(
          1,
        )}ms preprocess=${tPre?.toFixed?.(1)}ms intent=${tIntent?.toFixed?.(
          1,
        )}ms build=${tBuild?.toFixed?.(
          1,
        )}ms cache=${cache} user=${user} file=${file}`,
      );
    }
    // ✅ 절대 크래시 나지 않는 디버그 로그
    if (process.env.NODE_ENV !== "production") {
      console.log("[INTENT_DEBUG] message:", _dbgMessage);
      console.log(
        "[INTENT_DEBUG] intent:",
        JSON.stringify(_dbgIntent, null, 2),
      );
    }
  }
};

/* ---------------------------------------------
 * 피드백 핸들러
 * -------------------------------------------*/
exports.handleFeedback = async (req, res, next) => {
  try {
    const {
      message,
      result, // 프론트 표준: result로 보냄
      formula, // 호환
      isHelpful, // true=정확함, false=수정 필요
      reason, // ✅ 단일 필드
      conversionType = "Excel/Google Sheets",
      fileName,
    } = req.body || {};

    const msg = typeof message === "string" ? message.trim() : "";
    const out =
      (typeof result === "string" && result.trim()) ||
      (typeof formula === "string" && formula.trim()) ||
      "";
    const why = typeof reason === "string" ? reason.trim() : "";

    if (!msg || !out) {
      return res
        .status(400)
        .json({ error: "질문 내용과 결과가 모두 필요합니다." });
    }
    // ✅ '수정 필요'(isHelpful=false)면 reason 필수
    if (isHelpful === false && !why) {
      return res.status(400).json({
        error: "어떤 부분이 수정이 필요한지 알려주시면 도움이 됩니다.",
      });
    }

    const saved = await appendFeedback({
      ts: new Date().toISOString(),
      userId: req.user?.id ? String(req.user.id) : null,
      ip: req.ip || null,
      conversionType,
      fileName: fileName || null,
      message: msg,
      result: out,
      isHelpful: isHelpful === true ? true : isHelpful === false ? false : null,
      reason: why || null,
    });

    return res.status(200).json({ message: "피드백이 저장되었습니다.", saved });
  } catch (error) {
    next(error);
  }
};

/* ---------------------------------------------
 * 테스트/내부용 convert (LLM 미사용 경량)
 * -------------------------------------------*/
async function convert(nl, options = {}, meta = {}) {
  // 1) Intent 생성 (로컬 룰 or meta.intent 오버라이드)
  const baseIntent = meta.intent ? meta.intent : buildLocalIntentFromText(nl);
  let intent = normalizeLookupIntent(baseIntent);

  if (
    intent?.operation === "xlookup" &&
    intent?.needs_lookup_value === true &&
    intent?.lookup_value == null &&
    intent?.lookup?.value == null &&
    intent?.lookup?.value_ref == null
  ) {
    return '=ERROR("조회값이 없습니다. 예: 직원 ID가 1001인 이름 찾아줘")';
  }

  intent = applyMedianOverride(nl, intent);
  intent = applyExtremeRowOverride(nl, intent);
  intent = applyDateBoundaryOverride(nl, intent);
  intent = applyRecentTopNOverride(nl, intent);
  intent = applyMonthCountOverride(nl, intent);
  intent = applyYearCountOverride(nl, intent);
  intent = applyUniqueSortOverride(nl, intent);
  intent = applySortListOverride(nl, intent);
  intent = applyFilteredSortOverride(nl, intent);
  intent = applyRankColumnOverride(nl, intent);
  intent = applyGroupedAggregateOverride(nl, intent);
  intent = applyDuplicateLatestMetricOverride(nl, intent);
  intent = applyRankThresholdCountOverride(nl, intent);
  intent = applyTrimColumnOverride(nl, intent);
  intent = applyAverageThresholdIfOverride(nl, intent);

  // 2) 기본 컨텍스트 재료
  const engine = options.engine || DEFAULT_ENGINE;
  const policy = options.policy || DEFAULT_POLICY;
  const allSheetsData = meta.allSheetsData || null;

  /** @type {any} */
  let mergedMeta = {
    message: nl,
    engine,
    policy,
    intent,
    ...meta,
  };

  // 3) allSheetsData가 있으면, 자동 열 매핑(bestReturn / bestLookup) 시도
  if (allSheetsData) {
    const hasHints = !!(
      intent.return_hint ||
      intent.header_hint ||
      intent.lookup_hint
    );

    if (
      hasHints &&
      typeof formulaUtils.findBestSheetAndColumns === "function"
    ) {
      const searchTerms = {
        return: intent.return_hint || intent.header_hint || "",
        lookup: intent.lookup_hint || "",
      };

      const joint = formulaUtils.findBestSheetAndColumns(
        allSheetsData,
        searchTerms,
        {
          sameSheetBonus: 0.5,
        },
      );

      const bestReturn = joint?.return || null;
      const bestLookup = joint?.lookup || null;

      // bestReturn이 없는데도 sum/average 같은 집계 op를 요청하면
      // 테스트에서는 그냥 ERROR 문자열을 받게 해도 됨
      if (!bestReturn && (intent.header_hint || intent.return_hint)) {
        return '=ERROR("필요한 열을 파일에서 찾을 수 없습니다.")';
      }

      mergedMeta = {
        ...mergedMeta,
        allSheetsData,
        bestReturn,
        bestLookup,
      };
    } else {
      mergedMeta = { ...mergedMeta, allSheetsData };
    }
  }

  // 4) 정책/포맷 옵션 정규화
  const ctx = buildCtx(mergedMeta);

  // 5) 실제 빌더 호출
  const op = resolveOp(ctx.intent?.operation);
  if (!op) return '=ERROR("알 수 없는 operation 입니다.")';

  const built = formulaBuilder[op](
    ctx,
    (v, o) =>
      formulaUtils.formatValue(v, { ...ctx.formatOptions, ...(o || {}) }),
    formulaBuilder._buildConditionPairs,
    formulaBuilder._buildConditionMask,
  );
  // ✅ 조건 매칭 불확실로 인해 중단 요청이 들어온 경우
  if (ctx.__errorFormula) return ctx.__errorFormula;
  return built;
}

module.exports.convert = convert;
