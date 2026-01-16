const { cleanAIResponse } = require("../utils/responseHelper");
const formulaUtils = require("../utils/formulaUtils");
const { getOrBuildAllSheetsData } = require("../utils/sheetPreprocessor");
const { buildIntentCacheKey } = require("../utils/intentCacheKeyBuilder");
const { makeSheetStateSig } = require("../utils/sheetStateSig");
const intentCache = require("../services/intentCache");

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
  if (/(median|중앙값)/.test(s)) return "median";
  if (/(stdev|표준편차)/.test(s)) return "stdev_s";
  if (/(var|분산)/.test(s)) return "var_s";
  if (/(sortby|정렬)/.test(s)) return "sortby";
  return "formula";
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

  // ✅ 1. Lookup / 조회 패턴 감지
  // 예: "홍길동의 매출", "이름으로 점수 찾기"
  const lookupMatch = s.match(
    /([가-힣a-z0-9]+)[의\s]*(매출|점수|금액|이름|값|수량|가격)/i
  );
  if (op.includes("lookup") || /찾|조회|검색|lookup/.test(s)) {
    intent.operation = "xlookup";
    if (lookupMatch) {
      intent.lookup_hint = lookupMatch[1];
      intent.return_hint = lookupMatch[2];
    } else {
      // "OO의 매출" 패턴이 없으면 기본 힌트 추정
      if (/매출|sales?/.test(s)) intent.return_hint = "매출액";
      if (/이름|name/.test(s)) intent.lookup_hint = "이름";
    }
    return intent;
  }

  // ✅ 2. 조건부 합계 / 평균 / 개수 패턴
  // 예: "지점별 매출 합계", "서울 지역 평균 매출", "카테고리별 개수"
  if (/(sum|합계|total|평균|average|count|개수)/.test(s)) {
    intent.conditions = [];
    // 단순 조건 키워드 추출 ("서울", "지점", "카테고리")
    const condMatch = s.match(
      /(지역|지점|카테고리|부서|분류|도시|날짜|월|년도|날)/
    );
    if (condMatch) {
      intent.conditions.push({
        target: condMatch[1],
        operator: "=",
        value: "", // 나중에 LLM이 보완
      });
    }

    // 대상 열 추론
    if (/매출|sales?/.test(s)) intent.header_hint = "매출액";
    if (/평균/.test(s)) intent.operation = "averageifs";
    if (/합계|sum|total/.test(s)) intent.operation = "sumifs";
    if (/count|개수|갯수/.test(s)) intent.operation = "countifs";
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
      /([0-9]+)\s*(일|day|days|주|week|weeks|달|month|months)/
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

  return intent;
}

/* ---------------------------------------------
 * normalizeIntentPolicy(intent)
 * -------------------------------------------
 * 역할:
 *  - Intent 객체에 지정된 engine / mode / 기본값 정책을
 *    시스템 전역 기본값(DEFAULT_*) 기준으로 정규화한다.
 *
 * 처리 내용:
 *  1) engine: "excel" / "sheets" 이외의 값 → "excel"로 통일
 *  2) policy: mode / value_if_not_found / value_if_error 확정
 *  3) formatOptions: policy.mode에 따라 포맷팅 동작 제어
 *      - strict 모드 → 대소문자 구분, trim 비활성, 숫자 강제 변환 비활성
 *      - loose 모드 → 기본값(대소문자 무시, trim 활성, 숫자 변환)
 *
 * 결과:
 *  - { engine, policy, formatOptions } 반환
 * -------------------------------------------*/
function normalizeIntentPolicy(intent = {}) {
  const engineRaw =
    intent.engine || intent.platform || intent.target_engine || "";
  const engine =
    engineRaw === "sheets" || engineRaw === "googlesheets" ? "sheets" : "excel";

  const policy = {
    mode: intent.mode === "strict" ? "strict" : DEFAULT_POLICY.mode,
    value_if_not_found:
      intent.value_if_not_found != null
        ? String(intent.value_if_not_found)
        : DEFAULT_POLICY.value_if_not_found,
    value_if_error:
      intent.value_if_error != null
        ? String(intent.value_if_error)
        : DEFAULT_POLICY.value_if_error,
  };

  const formatOptions = { ...DEFAULT_FORMAT_OPTIONS };
  if (policy.mode === "strict") {
    formatOptions.case_sensitive = true;
    formatOptions.trim_text = false;
    formatOptions.coerce_number = false;
  }

  return { engine, policy, formatOptions };
}

/* ---------------------------------------------
 * buildCtx(rawCtx)
 * -------------------------------------------
 * 역할:
 *  - Intent 및 환경 정보를 기반으로 최종 Context 객체를 구성한다.
 *  - Builder 함수들이 참조할 모든 전역 설정을 이 단계에서 주입.
 *
 * 입력(rawCtx) 예시:
 *  {
 *    message: "홍길동의 매출 조회",
 *    intent: { operation: "xlookup", lookup_hint: "이름", return_hint: "매출" },
 *    engine: "excel",
 *    policy: { mode: "loose", value_if_not_found: "", value_if_error: "" },
 *    allSheetsData: {...},
 *    bestReturn: {...},
 *    bestLookup: {...}
 *  }
 *
 * 결과(Context):
 *  {
 *    intent,
 *    engine,
 *    policy,
 *    formatOptions,
 *    allSheetsData,
 *    bestReturn,
 *    bestLookup,
 *    formulaBuilder
 *  }
 * -------------------------------------------*/
function buildCtx(rawCtx) {
  const intent = rawCtx.intent || {};
  const { engine, policy, formatOptions } = normalizeIntentPolicy(intent);
  return { ...rawCtx, engine, policy, formatOptions };
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
          "lookup"
        );
        if (!best) return null;

        const range = `'${best.sheetName}'!${best.columnLetter}${best.startRow}:${best.columnLetter}${best.lastDataRow}`;

        const op = c.operator || "=";
        const rawVal = c.value;

        // 값도 반드시 포매터를 통과시켜 따옴표/숫자 처리
        const val = formulaBuilder._formatValue(rawVal);

        if (op !== "=" && rawVal != null && !isNaN(rawVal)) {
          return `${range}, "${op}${rawVal}"`;
        }
        return `${range}, ${val}`;
      })
      .filter(Boolean);
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

  sortby: "sortby",
  regexmatch: "regexmatch",
  textsplit: "textsplit",
};

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
  console.log("[loadAndPreprocess] user?.id:", user?.id);
  console.log("[loadAndPreprocess] fileName:", fileName);

  const bucket = getBucket();
  console.log("[loadAndPreprocess] bucket exists?:", !!bucket);

  if (!bucket || !user || !fileName) {
    console.log("[loadAndPreprocess] early return (no bucket/user/fileName)");
    return { isFileAttached: false, preprocessed: null };
  }

  console.log("[loadAndPreprocess] user.uploadedFiles:", user?.uploadedFiles);
  const fileInfo = user.uploadedFiles?.find((f) => f.originalName === fileName);
  console.log("[loadAndPreprocess] fileInfo:", fileInfo);

  if (!fileInfo) {
    console.log("[loadAndPreprocess] early return (fileInfo not found)");
    return { isFileAttached: false, preprocessed: null };
  }

  const file = bucket.file(fileInfo.gcsName);
  const [buffer] = await file.download();
  console.log("[loadAndPreprocess] downloaded buffer length:", buffer.length);

  const { fileHash, allSheetsData } = await getOrBuildAllSheetsData(buffer);
  console.log(
    "[loadAndPreprocess] got allSheetsData keys:",
    Object.keys(allSheetsData || {})
  );

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
    (fs) => fs && fs.isHelpful !== false && fs.intent && fs.message
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
  fewShots = []
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
          allHeaders.add(`'${h}'`)
        );
      });
      metaHintForModel = `The file contains columns like: [${Array.from(
        allHeaders
      ).join(", ")}]`;
    }

    // 3) 의도 추출 (OpenAI 있으면 LLM, 없으면 로컬)
    let intent = buildLocalIntentFromText(message);
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
        [] // fewShots disabled (no persistence)
      );

      if (llm && typeof llm === "object") {
        if (llm.intent && typeof llm.intent === "object") {
          intent = { ...intent, ...llm.intent };
        } else {
          intent = { ...intent, ...llm };
        }
      }
    }

    intent = normalizeLookupIntent(intent);
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
        600 // 10 min TTL (tune later)
      );
    }

    // 4) 컨텍스트 구성 + 자동 열 매핑
    _tBuildStart = process.hrtime.bigint();
    const context = { intent, formulaBuilder };
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
          }
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
      } else {
        Object.assign(context, { allSheetsData });
      }
    }

    // 5) direct(파일無) 빠른 경로
    if (!isFileAttached && direct?.canHandleWithoutFile?.(intent)) {
      const f = direct.buildFormula(intent);
      if (f) {
        if (req.user?.id && shouldCountConversion(f)) {
          await bumpUsage(req.user.id, "formulaConversions", 1);
        }
        return res.json({ result: f });
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
        formulaBuilder._buildConditionPairs
      );
    }
    _tBuildEnd = process.hrtime.bigint();

    if (req.user?.id && shouldCountConversion(finalFormula)) {
      await bumpUsage(req.user.id, "formulaConversions", 1);
    }
    return res.json({ result: finalFormula });
  } catch (err) {
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
          1
        )}ms preprocess=${tPre?.toFixed?.(1)}ms intent=${tIntent?.toFixed?.(
          1
        )}ms build=${tBuild?.toFixed?.(
          1
        )}ms cache=${cache} user=${user} file=${file}`
      );
    }
    // ✅ 절대 크래시 나지 않는 디버그 로그
    if (process.env.NODE_ENV !== "production") {
      console.log("[INTENT_DEBUG] message:", _dbgMessage);
      console.log(
        "[INTENT_DEBUG] intent:",
        JSON.stringify(_dbgIntent, null, 2)
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
      // 프론트 표준: reason 단일 필드로 통일
      reason,
      // 기존 필드들(호환): result/formula 둘 중 하나로 들어올 수 있음
      formula,
      result,
      // 원문 질문
      message,
      // "정확함/수정 필요" (boolean)
      isHelpful,
      // context (있으면 저장)
      conversionType = "Excel/Google Sheets",
      fileName,
    } = req.body || {};

    const finalMessage = typeof message === "string" ? message.trim() : "";
    const finalResult =
      (typeof formula === "string" && formula.trim()) ||
      (typeof result === "string" && result.trim()) ||
      "";
    const finalReason = typeof reason === "string" ? reason.trim() : "";

    if (!finalMessage || !finalResult) {
      return res.status(400).json({ error: "message and result are required" });
    }

    // ✅ 규칙: '수정 필요'(isHelpful=false)면 reason 필수
    if (isHelpful === false && !finalReason) {
      return res
        .status(400)
        .json({ error: "reason is required when isHelpful is false" });
    }

    const event = {
      ts: new Date().toISOString(),
      userId: req.user?.id ? String(req.user.id) : null,
      ip: req.ip || null,
      conversionType,
      fileName: fileName || null,
      message: finalMessage,
      result: finalResult,
      isHelpful: isHelpful === true ? true : isHelpful === false ? false : null,
      // reason 단일 저장 (정확함일 때는 null 가능)
      reason: finalReason || null,
    };

    appendFeedback(event);

    return res.status(200).json({ message: "피드백이 저장되었습니다." });
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
  const intent = normalizeLookupIntent(baseIntent);

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
        }
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

  return formulaBuilder[op](
    ctx,
    (v, o) =>
      formulaUtils.formatValue(v, { ...ctx.formatOptions, ...(o || {}) }),
    formulaBuilder._buildConditionPairs
  );
}

module.exports.convert = convert;
