const fs = require("fs");
const path = require("path");
const { cleanAIResponse } = require("../utils/responseHelper");
const formulaUtils = require("../utils/formulaUtils");
const { getOrBuildAllSheetsData } = require("../utils/sheetPreprocessor");
const { downloadToBuffer } = require("../utils/storage");
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
const {
  detectFormulaCompatibility,
  shouldAttemptCompatibilityFallback,
} = require("../utils/formulaCompatibility");
const { tryGenerateFallbackFormula } = require("../utils/formulaFallback");
const { buildConditionMask } = require("../utils/conditionEngine");

// === 빌더 모음 ===
const logicalFunctionBuilder = require("../builders/logicalFunctions");
const mathStatsFunctionBuilder = require("../builders/mathStatsFunctions");
const dateFunctionBuilder = require("../builders/dateFunctions");
const referenceFunctionBuilder = require("../builders/referenceFunctions");
const textFunctionBuilder = require("../builders/textFunctions");
const arrayFunctionBuilder = require("../builders/arrayFunctions");
const direct = require("../builders/direct");
const { shouldUseDirectBuilder: shouldUseDirectGate } = direct;

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

function sendFormulaResponse(res, payload = {}) {
  const excelFormula = payload.excelFormula ?? null;
  const sheetsFormula = Object.prototype.hasOwnProperty.call(
    payload,
    "sheetsFormula",
  )
    ? payload.sheetsFormula
    : excelFormula;
  const compatibility = payload.compatibility || {
    level: "common",
    blockers: [],
  };
  const debugMeta = payload.debugMeta || null;

  return res.json({
    excelFormula,
    sheetsFormula,
    compatibility,
    debugMeta,
  });
}

const INTENT_VERSION = "v2";
const CLUSTER_VERSION = "v1";

function getResolverMode(ctx = {}) {
  if (ctx?.allSheetsData && Object.keys(ctx.allSheetsData).length > 0) {
    return "sheet+cluster";
  }
  return "direct";
}

function shouldUseDirectBuilder(intent = {}, ctx = {}) {
  const raw = String(intent?.raw_message || ctx?.message || "");
  if (
    /["'][^"']+["']/.test(raw) &&
    /(포함|시작|끝나는|끝\s*나는|contains|starts?\s*with|ends?\s*with)/i.test(
      raw,
    )
  ) {
    return false;
  }
  return shouldUseDirectGate(intent, ctx);
}

/* ---------------------------------------------
 * 로컬 의도 추론 (LLM 미사용 시 폴백)
 * -------------------------------------------*/
function _deduceOp(text = "") {
  const s = String(text).toLowerCase();

  if (
    /(가장\s*(최근|오래|빠른|늦은|높은|낮은)|최근\s*순|오래된\s*순|높은\s*순|낮은\s*순|상위\s*\d+|하위\s*\d+|top\s*\d+|bottom\s*\d+)/i.test(
      s,
    ) &&
    /(직원|사람|행|레코드|항목|이름|목록|보여|가져와|출력)/i.test(s)
  ) {
    if (/(가장\s*(최근|늦은|높은)|최신|highest|latest)/i.test(s)) {
      return "maxrow";
    }
    if (/(가장\s*(오래|빠른|낮은)|oldest|earliest|lowest)/i.test(s)) {
      return "minrow";
    }
    return "topnrows";
  }
  if (/(중복\s*없이|중복\s*제거|unique|uniq)/i.test(s)) return "unique";
  if (/(세로로\s*(합치|합쳐|붙이|붙여)|vstack)/.test(s)) return "vstack";
  if (/(한\s*열로\s*펴|한열로\s*펴|tocol|세로로\s*펼쳐|flatten)/.test(s))
    return "tocol";
  if (/(각\s*행의\s*합계|행별\s*합계|각\s*행\s*합계|byrow)/.test(s))
    return "byrow";
  if (/(최근\s*순.*\d+|오래된\s*순.*\d+|최신.*\d+)/.test(s)) return "topnrows";
  if (/(목록을\s*보여줘|목록을\s*뽑아줘|리스트)/.test(s)) return "filter";
  if (/(상위\s*\d+명|top\s*\d+)/i.test(s)) return "topnrows";
  if (/(가장\s*높은|최고).*(가져와|보여|출력|알려|get|show|return)/i.test(s))
    return "maxrow";
  if (/(가장\s*낮은|최저).*(가져와|보여|출력|알려|get|show|return)/i.test(s))
    return "minrow";
  if (/(average|avg|mean|평균)/.test(s)) return "average";
  if (/(sum|total|합계|총합|합\b)/.test(s)) return "sum";
  if (/(count|개수|갯수|건수|수량|카운트|인원수|몇\s*명|명수)/.test(s)) {
    return "count";
  }
  if (
    /(?:[가-힣A-Za-z0-9_()\/ -]+)\s*(?:수|명)(?:\s|$|[을를이가은는의,.;!?])/.test(
      s,
    )
  ) {
    return "count";
  }
  if (/(xlookup|lookup|찾아|조회|검색|참조)/.test(s)) return "xlookup";
  if (/(filter|필터)/.test(s)) return "filter";
  if (/\b(if|조건|만약)\b/.test(s)) return "if";
  if (/(median|중앙값|중간값|가운데\s*값)/.test(s)) return "median";
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

  const metricHint = _extractMetricHintFromMessage(original);
  const lookupPhrases = _extractLookupPhrasesFromMessage(original);
  const explicitHeaderHint = _detectHeaderHintFromMessage(original);
  const headerHint =
    explicitHeaderHint || metricHint || lookupPhrases.returnHint || null;
  const groupBy = _detectGroupByFromMessage(original);
  let aggOp = _detectAggregateOpFromMessage(original, intent.operation);
  const sortOrder = _detectSortOrderFromMessage(original);
  const isUniqueListRequest = _looksLikeUniqueListRequest(original);
  const topNLimit = _extractTopNLimit(original);
  const wantsSortedList = _looksLikeSortedListRequest(original);
  const wantsRankColumn = _looksLikeRankColumnRequest(original);
  const wantsGroupedAggregate = _looksLikeGroupedAggregateRequest(original);
  const wantsGroupedSort = _looksLikeGroupedSortRequest(original);
  const sortPhraseHint = _extractSortHintFromMessage(original);
  const inferredSortHeader = _inferSortHeaderHint(
    original,
    sortPhraseHint || headerHint,
  );
  const inferredSortOrder = _inferSortOrderFromMessage(original);

  const explicitCellOrRange = formulaUtils.parseExplicitCellOrRange(original);
  const explicitCellMatch = original.match(/\b([A-Z]{1,3}\d{1,7})\b/);

  const quotedTextMatch = original.match(/["']([^"']+)["']/);
  const quotedText = quotedTextMatch ? quotedTextMatch[1] : null;
  const requestedReturnFields = _extractRequestedReturnFields(original);
  const wantsNameList =
    /(이름\s*목록|직원들의\s*이름|직원\s*이름\s*목록)/.test(original) ||
    ((/이름/.test(original) || /성명/.test(original)) &&
      /(목록|리스트|보여줘|뽑아줘|가져와줘)/.test(original));
  const hasContains =
    /포함/.test(original) || /contains|include/i.test(original);
  const hasStartsWith =
    /시작/.test(original) || /starts?\s*with/i.test(original);
  const hasEndsWith =
    /끝/.test(original) || /ends?\s*with|끝나는/i.test(original);

  const hasLookupCue =
    op.includes("lookup") || /찾|조회|검색|lookup/i.test(original);
  const hasGroup = Boolean(groupBy);
  const hasMetric = Boolean(aggOp && aggOp !== "formula");

  const filterSpecs = _extractRoleFilterSpecsFromMessage(original, headerHint);
  const legacyConditions = _filterSpecsToLegacyConditions(filterSpecs);
  const hasFilterSpecs = filterSpecs.length > 0;
  const hasDateFilter = filterSpecs.some((x) => x.role === "date_filter");

  if (wantsNameList) {
    intent.return_role = "entity_name";
    intent.return_fields = [];
  }
  if (hasMetric) intent.metric_role = "measure";
  if (groupBy) intent.group_role = "dimension";
  // 날짜 조건 + count 문구인데 aggregate op가 비어 있으면 count로 승격
  if (
    (!aggOp || aggOp === "formula") &&
    hasDateFilter &&
    _looksLikeCountRequest(original)
  ) {
    aggOp = "count";
  }
  if (hasLookupCue) {
    intent.lookup_role = "key";
    intent.return_role = "value";
  }
  if (filterSpecs.some((x) => x.role === "date_filter")) {
    intent.date_role = "date";
  }
  if (sortOrder) {
    intent.sort_spec = {
      order: sortOrder,
      role: hasMetric ? "measure_sort" : "generic_sort",
      header_hint: headerHint || null,
    };
  }
  if (hasFilterSpecs) {
    intent.filter_specs = filterSpecs;
  }

  const uniqueTarget =
    headerHint ||
    requestedReturnFields[0] ||
    _extractListTargetFromMessage(original) ||
    _extractSortHintFromMessage(original);

  if (isUniqueListRequest && uniqueTarget) {
    intent.operation = "unique";
    intent.return_fields = [uniqueTarget];
    intent.return_role = "list_item";
    intent.sorted = true;
    intent.sort = true;
    intent.sort_order = sortOrder || "asc";
    return intent;
  }

  if ((wantsGroupedAggregate || wantsGroupedSort) && groupBy) {
    intent.group_by = groupBy;
    intent.group_role = "dimension";

    // 집계 연산이 있으면 그걸 우선, 없고 직원 수/인원수 계열이면 count
    if (aggOp && aggOp !== "formula") {
      intent.operation = aggOp;
    } else if (/(인원수|개수|건수|수량)/.test(original)) {
      intent.operation = "count";
    } else {
      intent.operation = "count";
    }

    if (headerHint && intent.operation !== "count") {
      intent.header_hint = headerHint;
    }

    if (legacyConditions.length) {
      intent.conditions = legacyConditions;
    }

    if (wantsGroupedSort || sortOrder) {
      intent.sorted = true;
      intent.sort = true;
      intent.sort_order = inferredSortOrder || sortOrder || "desc";
    }

    return intent;
  }

  if (wantsRankColumn && inferredSortHeader) {
    intent.operation = "rankcolumn";
    intent.header_hint = inferredSortHeader;
    intent.sort_order = inferredSortOrder;
    intent.sorted = true;
    return intent;
  }

  if (topNLimit > 0 && inferredSortHeader) {
    intent.operation = "topnrows";
    intent.limit = topNLimit;
    intent.take_n = topNLimit;
    intent.header_hint = inferredSortHeader;
    intent.sort_order = inferredSortOrder;
    intent.sorted = true;

    if (requestedReturnFields.length) {
      intent.return_fields = requestedReturnFields;
    } else if (wantsNameList) {
      intent.return_fields = [];
      intent.return_role = "entity_name";
    }

    if (legacyConditions.length) {
      intent.conditions = legacyConditions;
    }
    return intent;
  }

  if (wantsSortedList && inferredSortHeader) {
    intent.operation = "sortby";
    intent.header_hint = inferredSortHeader;
    intent.sort_order = inferredSortOrder;
    intent.sorted = true;

    if (requestedReturnFields.length) {
      intent.return_fields = requestedReturnFields;
    } else if (wantsNameList) {
      intent.return_fields = [];
      intent.return_role = "entity_name";
    } else if (
      /(직원|사람|행|레코드|항목)/.test(original) ||
      (!requestedReturnFields.length &&
        !intent.return_role &&
        intent.operation === "sortby")
    ) {
      intent.return_role = "row";
      intent.return_fields = [];
    }

    if (legacyConditions.length) {
      intent.conditions = legacyConditions;
    }
    return intent;
  }

  if (/정렬|높은 순|낮은 순|오름차순|내림차순/.test(original)) {
    intent.operation = "sortby";
    if (!intent.return_fields && /이름/.test(original)) {
      intent.return_fields = [];
    }
    if (!intent.header_hint && inferredSortHeader) {
      intent.header_hint = inferredSortHeader;
    }
    if (sortOrder) {
      intent.sorted = true;
      intent.sort_order = sortOrder;
    }
    return intent;
  }

  if (hasLookupCue) {
    intent.operation = "xlookup";

    if (lookupPhrases.returnHint || headerHint) {
      intent.return_hint = lookupPhrases.returnHint || headerHint;
    }
    if (lookupPhrases.lookupHint) {
      intent.lookup_hint = lookupPhrases.lookupHint;
    }

    if (explicitCellOrRange?.ref) {
      intent.lookup_value = explicitCellOrRange.ref;
      intent.lookup = {
        ...(intent.lookup || {}),
        value_ref: explicitCellOrRange.ref,
      };
    } else if (explicitCellMatch) {
      const ref = explicitCellMatch[1].toUpperCase();
      intent.lookup_value = ref;
      intent.lookup = {
        ...(intent.lookup || {}),
        value_ref: ref,
      };
    }

    if (
      /(존재하지\s*않는|없는)/.test(original) &&
      intent.value_if_not_found == null
    ) {
      intent.value_if_not_found = "";
    }

    return intent;
  }

  // row-return 계열은 aggregate(max/min)로 덮어쓰지 않음
  if (_isRowEntityOp(op)) {
    intent.operation = op;

    const rankMetricHint = _extractRankMetricHint(original);
    if (rankMetricHint) {
      intent.header_hint = rankMetricHint;
    }

    if (requestedReturnFields.length) {
      intent.return_fields = requestedReturnFields;
    }

    if (!intent.header_hint && headerHint) {
      intent.header_hint = headerHint;
    }

    if (legacyConditions.length) {
      intent.conditions = legacyConditions;
    }

    if (sortOrder) {
      intent.sorted = true;
      intent.sort_order = sortOrder;
    }

    return intent;
  }

  if (
    hasGroup ||
    hasMetric ||
    (hasDateFilter && _looksLikeCountRequest(original))
  ) {
    intent.operation = aggOp || intent.operation || "formula";

    if (groupBy) {
      intent.group_by = groupBy;
    }
    if (headerHint && aggOp !== "count") {
      intent.header_hint = headerHint;
    }
    if (legacyConditions.length) {
      intent.conditions = legacyConditions;
    }
    if (sortOrder) {
      intent.sorted = true;
      intent.sort_order = sortOrder;
    }
    return intent;
  }

  if (wantsNameList) {
    intent.operation = "filter";
    intent.return_fields = [];

    intent.return_role = "entity_name";
    if (legacyConditions.length) {
      intent.conditions = legacyConditions;
    }

    return intent;
  }

  if (/(filter|필터|조건|만족|해당)/.test(s) || hasFilterSpecs) {
    intent.operation = "filter";
    if (legacyConditions.length) {
      intent.conditions = legacyConditions;
    }
    return intent;
  }

  if (/\bif\b|조건|이면|아니면|참|거짓/.test(s) && legacyConditions.length) {
    intent.operation = "if";
    intent.condition = legacyConditions[0];

    const labelMatches = s.match(/['"](.*?)['"]/g);
    if (labelMatches && labelMatches.length >= 2) {
      intent.value_if_true = labelMatches[0].replace(/['"]/g, "");
      intent.value_if_false = labelMatches[1].replace(/['"]/g, "");
    }
    return intent;
  }

  return intent;
}

function applyStructuralOverrides(intent) {
  if (!intent || typeof intent !== "object") return intent;

  if (Array.isArray(intent.return_fields) && intent.return_fields.length) {
    intent.return_fields = [...new Set(intent.return_fields.map(String))];
    delete intent.return_hint;
  }

  const hasLookup =
    intent.lookup_value != null ||
    !!intent.lookup_hint ||
    !!intent.lookup?.value ||
    !!intent.lookup?.header ||
    !!intent.lookup?.key_header;

  const hasGroup = !!intent.group_by;
  const hasMetric = !!intent.header_hint || !!intent.return_hint;
  const op = String(intent.operation || "").toLowerCase();
  const raw = String(intent.raw_message || "").trim();
  const cellLookupRef = raw
    .match(/\b([A-Z]{1,3}\d{1,7})\b/i)?.[1]
    ?.toUpperCase();
  const looksLikeCellLookup =
    !!cellLookupRef &&
    /(에\s*있는|의|로|기준|값)/.test(raw) &&
    /(가져와줘|찾아줘|조회해줘|보여줘|알려줘|return|lookup|get|show)/i.test(
      raw,
    );

  const rawLookupReturnFields = _extractRequestedReturnFields(raw);

  if (
    looksLikeCellLookup &&
    rawLookupReturnFields.length >= 1 &&
    !/(순으로|정렬|순위|랭크|등수|상위|하위|top|bottom)/i.test(raw)
  ) {
    intent.operation = "xlookup";
    intent.lookup_value = cellLookupRef;
    intent.lookup = {
      ...(intent.lookup || {}),
      value_ref: cellLookupRef,
      value: cellLookupRef,
      match_mode: "exact",
    };
    intent.lookup_role = intent.lookup_role || "entity_id";
    if (!intent.lookup_hint) {
      if (/(직원\s*ID|사번|employee\s*id)/i.test(raw)) {
        intent.lookup_hint = "직원 ID";
      } else if (/(직원|사람|인원|대상)/i.test(raw)) {
        intent.lookup_hint = "ID";
      }
    }
    intent.return_fields = rawLookupReturnFields;
    intent.return_role = "value";
    delete intent.return_hint;
    delete intent.group_by;
  }

  const hasAggregateBoundCondition =
    /([가-힣A-Za-z0-9_()\/ -]+?)\s*(?:이|가)?\s*(평균|average|avg|mean)\s*(이상|이하|초과|미만)/i.test(
      raw,
    );

  if (
    hasAggregateBoundCondition &&
    /(?:직원|행|레코드|항목).*(?:수|개수|건수|몇\s*명)|(?:수|개수|건수)$/i.test(
      raw,
    )
  ) {
    intent.operation = "count";
  }

  const aggregateBoundMatch = raw.match(
    /([가-힣A-Za-z0-9_]+)\s*(?:이|가)?\s*(평균|average|avg|mean)\s*(이상|이하|초과|미만)/i,
  );

  if (aggregateBoundMatch) {
    const header = aggregateBoundMatch[1];
    const agg = aggregateBoundMatch[2];
    const dir = aggregateBoundMatch[3];

    intent.filters = intent.filters || [];

    intent.filters.push({
      header,
      operator:
        dir === "이상"
          ? ">="
          : dir === "초과"
            ? ">"
            : dir === "이하"
              ? "<="
              : "<",
      value: agg,
      value_type: "aggregate",
      aggregate: agg,
      role: "aggregate_filter",
    });
  }

  const wantsUniqueList = _looksLikeUniqueListRequest(raw);
  const requestedReturnFields =
    Array.isArray(intent.return_fields) && intent.return_fields.length
      ? intent.return_fields
      : _extractRequestedReturnFields(raw);
  const uniqueTarget =
    intent.header_hint ||
    requestedReturnFields[0] ||
    _extractListTargetFromMessage(raw) ||
    _extractSortHintFromMessage(raw);
  const topNLimit = _extractTopNLimit(raw);
  const wantsSortedList = _looksLikeSortedListRequest(raw);
  const wantsRankColumn = _looksLikeRankColumnRequest(raw);
  const wantsGroupedAggregate = _looksLikeGroupedAggregateRequest(raw);
  const wantsGroupedSort = _looksLikeGroupedSortRequest(raw);
  const inferredSortHeader = _inferSortHeaderHint(raw, intent.header_hint);
  const inferredSortOrder = _inferSortOrderFromMessage(raw);

  if (
    wantsUniqueList &&
    (op === "sortby" ||
      op === "filter" ||
      op === "formula" ||
      op === "unique") &&
    uniqueTarget
  ) {
    intent.operation = "unique";
    intent.return_fields = [uniqueTarget];
    intent.sorted = true;
    intent.sort = true;
    intent.sort_order = intent.sort_order || "asc";
    delete intent.return_hint;
  }

  const wantsRowExtreme =
    /(가장\s*(최근|오래|빠른|늦은|높은|낮은)|최신|oldest|earliest|latest|highest|lowest)/i.test(
      raw,
    ) &&
    /(직원|사람|행|레코드|항목|이름|목록|보여줘|가져와줘|출력해줘|알려줘)/i.test(
      raw,
    );

  if (wantsRowExtreme && inferredSortHeader) {
    const asc = /가장\s*(오래|빠른|낮은)|oldest|earliest|lowest/i.test(raw);
    intent.operation = asc ? "minrow" : "maxrow";
    intent.header_hint = inferredSortHeader;
    intent.sort_order = asc ? "asc" : "desc";
    intent.sorted = true;
    intent.row_return = {
      mode: asc ? "min" : "max",
      sortBy: inferredSortHeader,
      take: 1,
    };

    const rawRequested = _extractRequestedReturnFields(raw);
    if (rawRequested.length) {
      intent.return_fields = rawRequested;
    } else if (/(이름|성명)/.test(raw)) {
      intent.return_role = "entity_name";
      intent.return_fields = [];
    } else if (/(\d+)\s*(?:명|개|건)/.test(raw)) {
      intent.operation = "topnrows";
      const n = _extractTopNLimit(raw);
      intent.limit = n || intent.limit || 1;
      intent.take_n = n || intent.take_n || 1;
      intent.return_role = "row";
      intent.return_fields = [];
      intent.return_all_columns = true;
    }
  }

  if (topNLimit > 0 && inferredSortHeader) {
    intent.operation = "topnrows";
    intent.limit = intent.limit || topNLimit;
    intent.take_n = intent.take_n || topNLimit;
    intent.header_hint = intent.header_hint || inferredSortHeader;
    intent.sort_order = intent.sort_order || inferredSortOrder;
    intent.sorted = true;
    intent.row_return = {
      mode: inferredSortOrder === "asc" ? "min" : "max",
      sortBy: intent.header_hint || inferredSortHeader,
      take: intent.take_n || intent.limit || topNLimit,
    };

    if (requestedReturnFields.length) {
      intent.return_fields = requestedReturnFields;
    }
  } else if (
    wantsSortedList &&
    inferredSortHeader &&
    (op === "formula" || op === "filter")
  ) {
    intent.operation = "sortby";
    intent.header_hint = intent.header_hint || inferredSortHeader;
    intent.sort_order = intent.sort_order || inferredSortOrder;
    intent.sorted = true;
  }

  const rawRequestedReturnFields = _extractRequestedReturnFields(raw);
  const hasOnlySortHintAsReturn =
    Array.isArray(intent.return_fields) &&
    intent.return_fields.length === 1 &&
    inferredSortHeader &&
    String(intent.return_fields[0]).trim() ===
      String(inferredSortHeader).trim();

  if (
    String(intent.operation || "").toLowerCase() === "sortby" &&
    Array.isArray(intent.return_fields) &&
    intent.return_fields.length >= 2 &&
    String(intent.return_role || "").toLowerCase() === "entity_name"
  ) {
    delete intent.return_role;
  }

  // sortby인데 명시 반환열이 없고 "행/대상 목록" 요청이면 전체 행 반환으로 보정
  if (
    String(intent.operation || "").toLowerCase() === "sortby" &&
    wantsSortedList &&
    inferredSortHeader &&
    (!rawRequestedReturnFields.length || hasOnlySortHintAsReturn) &&
    String(intent.return_role || "").toLowerCase() !== "entity_name"
  ) {
    intent.return_role = "row";
    intent.return_fields = [];
    intent.return_all_columns = true;
    delete intent.return_hint;
    intent.header_hint = intent.header_hint || inferredSortHeader;
    intent.sort_order = intent.sort_order || inferredSortOrder || "desc";
    intent.sorted = true;
  }

  if (
    wantsRankColumn &&
    inferredSortHeader &&
    (op === "formula" || op === "sortby" || op === "filter")
  ) {
    intent.operation = "rankcolumn";
    intent.header_hint = intent.header_hint || inferredSortHeader;
    intent.sort_order = intent.sort_order || inferredSortOrder;
    intent.sorted = true;
  }

  if ((wantsGroupedAggregate || wantsGroupedSort) && intent.group_by) {
    if (op === "sortby" || op === "formula" || op === "filter") {
      if (/(인원수|개수|건수|수량)/.test(raw)) {
        intent.operation = "count";
      } else if (
        !intent.operation ||
        intent.operation === "formula" ||
        intent.operation === "sortby"
      ) {
        intent.operation = intent.header_hint ? "average" : "count";
      }
    }
    if (wantsGroupedSort) {
      intent.sorted = true;
      intent.sort = true;
      intent.sort_order = intent.sort_order || inferredSortOrder || "desc";
    }
  }

  const wantsRowReturnVerb =
    /(가져와줘|보여줘|출력해줘|알려줘|get|show|return)/i.test(raw);
  const hasMultiReturnFields = requestedReturnFields.length >= 2;

  if (
    (op === "max" || op === "min") &&
    wantsRowReturnVerb &&
    hasMultiReturnFields
  ) {
    intent.operation = op === "max" ? "maxrow" : "minrow";
    intent.return_fields = requestedReturnFields;
  }

  // explicit-range 구조 연산은 LLM 결과보다 우선.
  const explicitRanges =
    raw.match(/[A-Z]+[0-9]+:[A-Z]+[0-9]+|[A-Z]+:[A-Z]+/gi) || [];
  const explicitSingle =
    raw.match(/[A-Z]+[0-9]+:[A-Z]+[0-9]+|[A-Z]+:[A-Z]+/i)?.[0] || null;

  if (/(세로로\s*(합치|합쳐|붙이|붙여)|vstack)/i.test(raw)) {
    intent.operation = "vstack";
    if (explicitRanges.length >= 2) intent.ranges = explicitRanges;
    return intent;
  }

  if (/(한\s*열로\s*펴|한열로\s*펴|tocol|세로로\s*펼쳐|flatten)/i.test(raw)) {
    intent.operation = "tocol";
    if (explicitSingle) intent.range = explicitSingle;
    return intent;
  }

  if (/(각\s*행의\s*합계|행별\s*합계|각\s*행\s*합계|byrow)/i.test(raw)) {
    intent.operation = "byrow";
    intent.aggregate = intent.aggregate || "sum";
    if (explicitSingle) {
      intent.range = explicitSingle;
    } else {
      intent.require_explicit_range = true;
    }
    return intent;
  }

  if (op === "xlookup" && hasGroup && !hasLookup) {
    intent.operation = hasMetric ? "sum" : "count";
  }

  if (op === "xlookup" && hasLookup && hasGroup && !hasMetric) {
    delete intent.group_by;
  }

  if (op === "xlookup" && !hasLookup && raw) {
    const hasDateCue = /(날짜|일자|date|time)/i.test(raw);
    const hasRecentCue = /(가장\s*최근|최근|최신|latest|most\s*recent)/i.test(
      raw,
    );
    const hasOldestCue =
      /(가장\s*오래|오래\s*근무|최장\s*근무|earliest|oldest)/i.test(raw);
    const wantsRowEntity = /(정보|행|레코드|row|record)/i.test(raw);

    if (hasDateCue && wantsRowEntity && (hasRecentCue || hasOldestCue)) {
      intent.operation = hasRecentCue ? "maxrow" : "minrow";

      if (
        !Array.isArray(intent.return_fields) ||
        intent.return_fields.length === 0
      ) {
        intent.return_fields = [];
      }

      delete intent.lookup_hint;
      delete intent.lookup_value;
      if (intent.lookup) delete intent.lookup;
      if (intent.lookup_array) delete intent.lookup_array;
    }
  }

  return intent;
}

/* ---------------------------------------------
 * normalizeLookupIntent(intent)
 * -------------------------------------------
 * 역할:
 *  - LLM 또는 로컬 룰 기반 Intent 중
 *    lookup / xlookup 계열의 입력 필드를 표준 구조로 보정.
 *
 * 표준화 내용:
 *  1) LLM이 준 lookup_key / return 필드를 lookup_array / return_array로 변환
 *  2) lookup_value, lookup_array, return_array를 보장
 *  3) referenceFunctions 등 빌더가 기대하는 intent.lookup / intent.return 구조를 생성
 *
 * 결과:
 *  - 모든 xlookup Intent는 아래 필드를 최소 포함.
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

  if (Array.isArray(intent.return_fields) && intent.return_fields.length > 1) {
    intent.return_arrays = intent.return_fields.map((header) => ({
      header,
      sheet: intent.return?.sheet || intent.return_array?.sheet || undefined,
    }));
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
  unique: "unique",
  regexmatch: "regexmatch",
  textsplit: "textsplit",
};

function _detectGroupByFromMessage(msg = "") {
  const s = String(msg || "");
  const m = s.match(/([가-힣A-Za-z0-9_]+)\s*별/);
  if (m) return m[1];
  return null;
}

function _detectHeaderHintFromMessage(msg = "") {
  const s = String(msg || "");
  return null;
}

function _detectAggregateOpFromMessage(msg = "", fallbackOp = "") {
  const s = String(msg || "");
  if (/(평균|average|avg|mean)/i.test(s)) return "average";
  if (/(합계|총합|sum|total)/i.test(s)) return "sum";
  if (/(개수|갯수|건수|인원수|수량|count)/i.test(s)) return "count";
  if (/(최고|최대|가장\s*높|max|highest)/i.test(s)) return "max";
  if (/(최저|최소|가장\s*낮|min|lowest)/i.test(s)) return "min";
  if (/(중앙값|중간값|가운데\s*값|median)/i.test(s)) return "median";
  return fallbackOp || "formula";
}

function _detectSortOrderFromMessage(msg = "") {
  const s = String(msg || "");
  if (/(적은\s*순|낮은\s*순|오름차순|asc|작은\s*순)/i.test(s)) return "asc";
  if (/(많은\s*순|높은\s*순|내림차순|desc|큰\s*순)/i.test(s)) return "desc";
  return null;
}

function _looksLikeUniqueListRequest(msg = "") {
  const s = String(msg || "");
  return (
    /(중복\s*없이|중복\s*제거|unique|uniq)/i.test(s) &&
    /(목록|리스트|뽑|보여|정렬|가나다순|sort)/i.test(s)
  );
}

function _extractTopNLimit(msg = "") {
  const s = String(msg || "");
  const m =
    s.match(/(?:상위|하위|top|bottom)\s*(\d+)/i) ||
    s.match(/(\d+)\s*(?:명|개|건)(?:\s|$|[을를이가은는의,.;!?])/i) ||
    s.match(/(?:최근|최신|오래된|높은|낮은)\s*순(?:으로)?\s*(\d+)/i);
  return m ? Number(m[1]) : 0;
}

function _looksLikeSortedListRequest(msg = "") {
  const s = String(msg || "");
  return (
    /(순으로|정렬|오름차순|내림차순|상위|하위|top|bottom|최근|최신|오래된)/i.test(
      s,
    ) && /(목록|리스트|보여줘|뽑아줘|가져와줘|출력해줘)/i.test(s)
  );
}

function _looksLikeRankColumnRequest(msg = "") {
  const s = String(msg || "");
  return /(순위|랭크|등수|rank)/i.test(s);
}

function _looksLikeGroupedAggregateRequest(msg = "") {
  const s = String(msg || "");
  return (
    /(별|별로)/.test(s) &&
    /(표로|계산|구해|인원|수|합계|평균|최고|최저|중앙)/.test(s)
  );
}

function _looksLikeGroupedSortRequest(msg = "") {
  const s = String(msg || "");
  return (
    /(별|별로)/.test(s) &&
    /(정렬|순으로|높은\s*순|낮은\s*순|많은\s*순|적은\s*순)/.test(s)
  );
}

function _extractRankMetricHint(msg = "") {
  const raw = String(msg || "").trim();
  if (!raw) return null;

  // 날짜/근무 기간 문맥의 최신/오래됨/빠름/늦음은 날짜열 기준 정렬
  if (
    /(입사|근무|재직|날짜|일자|date)/i.test(raw) &&
    /(가장\s*(최근|오래|빠른|늦은)|최근\s*순|오래된\s*순|최신|oldest|earliest|latest)/i.test(
      raw,
    )
  ) {
    return "날짜";
  }

  // "연봉 순위 열", "매출 순위 컬럼" 등 rank column 기준 추출
  const rankColMetric = raw.match(
    /([가-힣A-Za-z0-9_()\/ -]+?)\s*(?:순위|랭크|등수|rank)\s*(?:열|컬럼|column)?/i,
  );
  if (rankColMetric) {
    const phrase = _cleanHintPhrase(rankColMetric[1])
      .replace(/^(?:전체|모든)\s*/u, "")
      .replace(/\s*(?:직원|사람|행|레코드|항목)\s*(?:의)?$/u, "")
      .trim();
    if (phrase) return phrase;
  }

  const patterns = [
    /([가-힣A-Za-z0-9_()\/ -]+?)\s*(?:이|가)?\s*가장\s*(?:높은|낮은)/,
    /([가-힣A-Za-z0-9_()\/ -]+?)\s*(?:상위|하위|top|bottom)\s*\d*/i,
    /(.+?)\s*(?:높은\s*순|낮은\s*순|많은\s*순|적은\s*순|오름차순|내림차순)/,
  ];

  for (const p of patterns) {
    const m = raw.match(p);
    const phrase = _cleanHintPhrase(m?.[1] || "")
      .split(/에서|중에서|among|within/i)
      .pop()
      .replace(/^[가-힣A-Za-z0-9_]+\s*부서\s*/u, "")
      .replace(/^[가-힣A-Za-z0-9_]+\s*팀\s*/u, "")
      .replace(
        /\d+(?:,\d{3})*(?:\.\d+)?\s*(?:이상|이하|초과|미만|부터|까지|>=|<=|>|<).*$/u,
        "",
      )
      .replace(/\s*(?:직원|사람|행|레코드|항목)\s*$/u, "")
      .replace(/\s*(?:이|가|은|는|을|를|의)$/u, "")
      .trim();

    if (phrase && phrase.length <= 40) return phrase;
  }

  return null;
}

function _inferSortHeaderHint(msg = "", fallback = null) {
  return _extractRankMetricHint(msg) || fallback || null;
}

function _inferSortOrderFromMessage(msg = "") {
  const s = String(msg || "");
  if (/(가장\s*(오래|빠른)|오래된\s*순|예전|earliest|oldest)/i.test(s)) {
    return "asc";
  }
  if (/(가장\s*(최근|늦은)|최근\s*순|최신|latest|most\s*recent)/i.test(s)) {
    return "desc";
  }
  if (
    /(하위|낮은\s*순|작은\s*순|오름차순|asc|오래된|예전|earliest|oldest)/i.test(
      s,
    )
  )
    return "asc";
  if (
    /(상위|높은\s*순|큰\s*순|내림차순|desc|최근|최신|latest|most\s*recent)/i.test(
      s,
    )
  )
    return "desc";
  return _detectSortOrderFromMessage(s) || "desc";
}

function _isRowEntityOp(op = "") {
  const k = String(op || "").toLowerCase();
  return ["maxrow", "minrow", "topnrows"].includes(k);
}

function _cleanReturnFieldToken(token = "") {
  let s = String(token || "").trim();
  if (!s) return "";

  // 일반적인 군더더기 제거 (도메인 특정 아님)
  s = s.replace(/^(?:행|레코드|항목)\s*의\s*/u, "");
  s = s.replace(/\s*(?:을|를|은|는|이|가)$/u, "");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

function _extractRequestedReturnFields(msg = "") {
  const raw = String(msg || "");
  if (!raw) return [];

  // 정렬/필터 후 "직원/행/항목을 보여줘"는 특정 반환열 요청이 아니라 row-return 요청이다.
  if (
    /(순으로|정렬|오름차순|내림차순|높은\s*순|낮은\s*순|상위|하위|top|bottom)/i.test(
      raw,
    ) &&
    /(행|레코드|항목|목록|리스트|직원|사람)/i.test(raw) &&
    !/(?:^|[,，및그리고와과\s])(이름|성명|직급|부서|ID|아이디|연봉|급여|나이|입사일)(?:\s*(?:,|및|그리고|와|과)|\s*(?:을|를)?\s*(?:보여줘|출력해줘|가져와줘|알려줘))/i.test(
      raw,
    )
  ) {
    return [];
  }

  // 셀 참조 lookup + 다중 반환열
  const cellLookupFieldEnum = raw.match(
    /\b[A-Z]{1,3}\d{1,7}\b.*?(?:의|에\s*있는)\s+([가-힣A-Za-z0-9_ ]+?)\s*(?:,|및|그리고|와|과)\s*([가-힣A-Za-z0-9_ ]+?)\s*(?:을|를)?\s*(?:동시에\s*)?(?:가져와줘|보여줘|출력해줘|알려줘|get|show|return)/i,
  );
  if (cellLookupFieldEnum) {
    return [
      ...new Set(
        [cellLookupFieldEnum[1], cellLookupFieldEnum[2]]
          .map((x) =>
            _cleanReturnFieldToken(x)
              .replace(/^(?:직원|사람|항목|레코드)\s*의\s*/u, "")
              .replace(/\s*동시에\s*$/u, "")
              .trim(),
          )
          .filter(Boolean),
      ),
    ];
  }

  const entityFieldAfterSort = raw.match(
    /(?:높은\s*순|낮은\s*순|오름차순|내림차순|정렬|순으로)\s*(?:으로)?\s*(?:직원|사람|항목|레코드)?\s*([가-힣A-Za-z0-9_ ]+?)\s*(?:,|및|그리고|와|과)\s*([가-힣A-Za-z0-9_ ]+?)\s*(?:을|를)?\s*(?:보여줘|출력해줘|가져와줘|알려줘|get|show|return)/i,
  );
  if (entityFieldAfterSort) {
    return [
      ...new Set(
        [entityFieldAfterSort[1], entityFieldAfterSort[2]]
          .map(_cleanReturnFieldToken)
          .filter(Boolean),
      ),
    ];
  }

  // "직원 이름과 연봉"처럼 대상어+열거가 섞인 요청 보정
  const entityFieldEnum = raw.match(
    /(?:직원|사람|항목|레코드)?\s*([가-힣A-Za-z0-9_ ]+?)\s*(?:,|및|그리고|와|과)\s*([가-힣A-Za-z0-9_ ]+?)\s*(?:을|를)?\s*(?:보여줘|출력해줘|가져와줘|알려줘|get|show|return)/i,
  );
  if (entityFieldEnum) {
    return [
      ...new Set(
        [entityFieldEnum[1], entityFieldEnum[2]]
          .map(_cleanReturnFieldToken)
          .filter(Boolean),
      ),
    ];
  }

  // “... 이름, 부서, 직급, 연봉을 가져와줘/보여줘 ...” 같은 열거형 요청을 일반적으로 추출
  const capture =
    raw.match(
      /(?:의|에 해당하는)\s+([^.?]+?)\s*(?:을|를)?\s*(?:가져와줘|보여줘|출력해줘|알려줘|get|show|return)/i,
    ) ||
    raw.match(
      /([^.?]+?)\s*(?:을|를)?\s*(?:가져와줘|보여줘|출력해줘|알려줘|get|show|return)/i,
    );

  const zone = capture?.[1] || "";
  if (!zone) return [];

  const normalized = zone.replace(/\s*(?:,|및|그리고|와|과)\s*/g, ",");
  const fields = normalized
    .split(",")
    .map(_cleanReturnFieldToken)
    .filter(Boolean)
    .filter(
      (x) =>
        !/^(?:가장\s*높은|가장\s*낮은|최고|최저|상위|하위|직원|목록)$/u.test(x),
    );

  return [...new Set(fields)];
}

function _cleanHintPhrase(s = "") {
  return String(s || "")
    .replace(/[\"']/g, "")
    .replace(/\b(을|를|은|는|이|가|의|로|으로)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function _extractMetricHintFromMessage(msg = "") {
  const raw = String(msg || "").trim();
  if (!raw) return null;

  const patterns = [
    /(.+?)\s*(?:평균|합계|총합|최고값|최저값|최고|최저|중앙값|중간값)\b/,
    /(.+?)\s*(?:average|sum|total|max|min|median)\b/i,
  ];

  for (const p of patterns) {
    const m = raw.match(p);
    const phrase = _cleanHintPhrase(m?.[1] || "");
    if (phrase && phrase.length <= 40) return phrase;
  }
  return null;
}

function _extractLookupPhrasesFromMessage(msg = "") {
  const raw = String(msg || "").trim();
  if (!raw) return { lookupHint: null, returnHint: null };

  // "... X로 Y를 찾아줘"
  let m = raw.match(
    /(.+?)로\s+(.+?)\s*(?:을|를)?\s*(?:찾아줘|조회해줘|검색해줘|lookup)/i,
  );
  if (m) {
    return {
      lookupHint: _cleanHintPhrase(m[1]),
      returnHint: _cleanHintPhrase(m[2]),
    };
  }

  // "... X의 Y를 보여줘"
  m = raw.match(/(.+?)의\s+(.+?)\s*(?:을|를)?\s*(?:보여줘|가져와줘|알려줘)/i);
  if (m) {
    return {
      lookupHint: _cleanHintPhrase(m[1]),
      returnHint: _cleanHintPhrase(m[2]),
    };
  }

  return { lookupHint: null, returnHint: null };
}

function _extractSortHintFromMessage(msg = "") {
  const raw = String(msg || "").trim();
  if (!raw) return null;

  const patterns = [
    /(.+?)\s*(?:높은\s*순|낮은\s*순|많은\s*순|적은\s*순|오름차순|내림차순)/,
    /(.+?)\s*(?:recent|latest|oldest|asc|desc)/i,
  ];

  for (const p of patterns) {
    const m = raw.match(p);
    const phrase = _cleanHintPhrase(m?.[1] || "");
    if (phrase && phrase.length <= 40) return phrase;
  }
  return null;
}

function _extractListTargetFromMessage(msg = "") {
  const raw = String(msg || "");
  const m = raw.match(/([가-힣A-Za-z0-9_()\/ -]+?)\s*목록/u);
  if (!m?.[1]) return null;
  return _cleanHintPhrase(m[1])
    .replace(/중복\s*없이/g, "")
    .replace(/가나다순으로?/g, "")
    .replace(/정렬해줘/g, "")
    .trim();
}

function _pushUniqueFilterSpec(out, spec) {
  if (!spec || typeof spec !== "object") return;
  const key = JSON.stringify([
    spec.role || "",
    spec.header_hint || "",
    spec.operator || "",
    spec.value ?? "",
    spec.value_type || "",
    spec.min ?? "",
    spec.max ?? "",
  ]);
  if (!out.__seen) out.__seen = new Set();
  if (out.__seen.has(key)) return;
  out.__seen.add(key);
  out.push(spec);
}

function _finalizeFilterSpecs(out) {
  if (out && out.__seen) delete out.__seen;
  return Array.isArray(out) ? out : [];
}

function _cleanCategoryValue(v) {
  let s = String(v || "").trim();
  if (!s) return s;

  // 조사/연결어 제거
  s = s.replace(/(이면서|이고|이며|이고도)$/u, "");
  s = s.replace(/(인|인데|인 직원|인 사원)$/u, "");
  s = s.replace(/(이|가|은|는|을|를|의)$/u, "");

  return s.trim();
}

function _looksLikeCountRequest(msg = "") {
  const s = String(msg || "");
  return /(인원수|개수|갯수|건수|수량|카운트|몇\s*명)/.test(s);
}

function _extractRoleFilterSpecsFromMessage(msg = "", headerHint = null) {
  const original = String(msg || "");
  const out = [];

  // 범용 카테고리 조건: "개발 부서", "영업 부서에서", "마케팅 팀"
  for (const m of original.matchAll(
    /([가-힣A-Za-z0-9_]+)\s*(부서|팀)(?:에서|의|이면서|이고|이며|인)?/gu,
  )) {
    const value = _cleanCategoryValue(m[1]);
    const header = String(m[2] || "").trim();
    if (!value || !header) continue;

    if (/^(직원|사람|항목|레코드|대상|값)$/u.test(value)) continue;

    _pushUniqueFilterSpec(out, {
      role: "category_filter",
      header_hint: header,
      operator: "=",
      value,
      value_type: "text",
    });
  }

  const quotedTextMatch = original.match(/["']([^"']+)["']/);
  const quotedText = quotedTextMatch ? quotedTextMatch[1].trim() : null;

  // 순서형 텍스트 조건: "직급이 대리 이상", "직급이 과장 이하"
  const ordinalTextMatch = original.match(
    /([가-힣A-Za-z0-9_]+)\s*(?:이|가)?\s*([가-힣A-Za-z0-9_]+)\s*(이상|이하|초과|미만)/u,
  );
  if (ordinalTextMatch) {
    const header = ordinalTextMatch[1].trim();
    const value = ordinalTextMatch[2].trim();
    const bound = ordinalTextMatch[3].trim();

    if (/(직급|등급|단계|레벨|level|grade)/i.test(header)) {
      const op =
        bound === "이상"
          ? ">="
          : bound === "이하"
            ? "<="
            : bound === "초과"
              ? ">"
              : "<";

      _pushUniqueFilterSpec(out, {
        role: "ordinal_filter",
        header_hint: header,
        operator: op,
        value,
        value_type: "ordinal_text",
      });
    }
  }

  // 집계 기준 조건: "연봉이 평균 이상", "점수가 평균 이하"
  const aggregateBoundMatch = original.match(
    /([가-힣A-Za-z0-9_()\/ -]+?)\s*(?:이|가)?\s*(평균|average|avg|mean)\s*(이상|이하|초과|미만)/i,
  );
  if (aggregateBoundMatch) {
    const header = _cleanHintPhrase(aggregateBoundMatch[1]);
    const agg = String(aggregateBoundMatch[2]).toLowerCase();
    const bound = aggregateBoundMatch[3];

    const op =
      bound === "이상"
        ? ">="
        : bound === "이하"
          ? "<="
          : bound === "초과"
            ? ">"
            : "<";

    _pushUniqueFilterSpec(out, {
      role: "aggregate_filter",
      header_hint: header,
      operator: op,
      value_type: "aggregate",
      aggregate: /avg|mean|average|평균/i.test(agg) ? "average" : agg,
    });
  }

  // 1) 숫자 조건
  if (headerHint) {
    let operator = null;
    if (/(이상|greater|over|>=|>)/i.test(original)) operator = ">=";
    else if (/(이하|under|<=|<|작은)/i.test(original)) operator = "<=";
    else if (/(초과)/i.test(original)) operator = ">";
    else if (/(미만)/i.test(original)) operator = "<";
    else if (/(같|=|equal)/i.test(original)) operator = "=";

    const numMatch = original.match(/([0-9][0-9,]*(?:\.[0-9]+)?)/);
    if (operator && numMatch) {
      _pushUniqueFilterSpec(out, {
        role: "metric_filter",
        header_hint: headerHint,
        operator,
        value: numMatch[1].replace(/,/g, ""),
        value_type: "number",
      });
    }
  }

  // 2) 날짜 조건
  const isoRange = original.match(
    /(\d{4}[./-]\d{1,2}[./-]\d{1,2})\s*(?:부터|~|-)\s*(\d{4}[./-]\d{1,2}[./-]\d{1,2})/,
  );
  if (isoRange) {
    _pushUniqueFilterSpec(out, {
      role: "date_filter",
      header_hint: null,
      operator: "between",
      min: isoRange[1].replace(/[./]/g, "-"),
      max: isoRange[2].replace(/[./]/g, "-"),
      value_type: "date",
    });
  } else {
    const isoAfter = original.match(
      /(\d{4}[./-]\d{1,2}[./-]\d{1,2})\s*(이후|후|부터)/,
    );
    if (isoAfter) {
      _pushUniqueFilterSpec(out, {
        role: "date_filter",
        header_hint: null,
        operator: ">=",
        value: isoAfter[1].replace(/[./]/g, "-"),
        value_type: "date",
      });
    }

    const isoBefore = original.match(
      /(\d{4}[./-]\d{1,2}[./-]\d{1,2})\s*(이전|전|까지)/,
    );
    if (isoBefore) {
      _pushUniqueFilterSpec(out, {
        role: "date_filter",
        header_hint: null,
        operator: "<=",
        value: isoBefore[1].replace(/[./]/g, "-"),
        value_type: "date",
      });
    }

    const yearRange = original.match(
      /(20\d{2})\s*년?\s*[~\-부터]\s*(20\d{2})\s*년?/,
    );
    if (yearRange) {
      _pushUniqueFilterSpec(out, {
        role: "date_filter",
        header_hint: null,
        operator: "between",
        min: `${yearRange[1]}-01-01`,
        max: `${yearRange[2]}-12-31`,
        value_type: "date",
      });
    } else {
      const yearAfter = original.match(/(20\d{2})\s*년?\s*(이후|후)/);
      if (yearAfter) {
        _pushUniqueFilterSpec(out, {
          role: "date_filter",
          header_hint: null,
          operator: ">=",
          value: `${yearAfter[1]}-01-01`,
          value_type: "date",
        });
      }

      const yearBefore = original.match(/(20\d{2})\s*년?\s*(이전|전)/);
      if (yearBefore) {
        _pushUniqueFilterSpec(out, {
          role: "date_filter",
          header_hint: null,
          operator: "<",
          value: `${yearBefore[1]}-01-01`,
          value_type: "date",
        });
      }
    }
  }

  // 3) 텍스트 매칭 조건
  if (quotedText) {
    const textHeaderHint = headerHint || null;

    if (textHeaderHint) {
      if (/포함/.test(original) || /contains|include/i.test(original)) {
        _pushUniqueFilterSpec(out, {
          role: "text_filter",
          header_hint: textHeaderHint,
          operator: "contains",
          value: quotedText,
          value_type: "text",
        });
      } else if (/시작/.test(original) || /starts?\s*with/i.test(original)) {
        _pushUniqueFilterSpec(out, {
          role: "text_filter",
          header_hint: textHeaderHint,
          operator: "starts_with",
          value: quotedText,
          value_type: "text",
        });
      } else if (/끝/.test(original) || /ends?\s*with|끝나는/i.test(original)) {
        _pushUniqueFilterSpec(out, {
          role: "text_filter",
          header_hint: textHeaderHint,
          operator: "ends_with",
          value: quotedText,
          value_type: "text",
        });
      }
    }
  }

  return _finalizeFilterSpecs(out);
}

function _filterSpecsToLegacyConditions(filterSpecs = []) {
  return (Array.isArray(filterSpecs) ? filterSpecs : [])
    .map((spec) => {
      if (!spec) return null;
      if (spec.operator === "between" && spec.min != null && spec.max != null) {
        return {
          target: spec.header_hint || spec.header || null,
          operator: "between",
          min: spec.min,
          max: spec.max,
          value_type: spec.value_type || null,
        };
      }
      return {
        target: spec.header_hint || spec.header || null,
        operator: spec.operator || "=",
        value: spec.value,
        value_type: spec.value_type || null,
      };
    })
    .filter(Boolean);
}

function resolveOp(op) {
  if (!op) return null;
  const k = String(op).toLowerCase().replace(/[ \-]/g, "");
  const base = OP_ALIASES[k] || k;
  return typeof formulaBuilder[base] === "function" ? base : null;
}

function getDebugReturnHeaders(ctx = {}) {
  return (ctx?.resolved?.returnColumns || [])
    .map((x) => {
      const h = String(x?.header || "").trim();
      if (!h) return null;
      if (/^-?\d+(?:\.\d+)?$/.test(h.replace(/,/g, ""))) return null;
      return h;
    })
    .filter(Boolean);
}

/* ---------------------------------------------
 * 파일 전처리 유틸
 * -------------------------------------------*/
const LOCAL_TEST_DIRS = [
  path.join(__dirname, "..", ".local_test_files"),
  path.join(__dirname, "..", ".local_uploads"),
];

function normalizeNameForLocalMatch(name) {
  return String(name || "")
    .trim()
    .normalize("NFC")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function walkFilesRecursive(rootDir) {
  const out = [];
  if (!fs.existsSync(rootDir)) return out;

  const stack = [rootDir];
  while (stack.length) {
    const current = stack.pop();
    let entries = [];
    try {
      entries = fs.readdirSync(current, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      const fullPath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        stack.push(fullPath);
      } else if (entry.isFile()) {
        out.push(fullPath);
      }
    }
  }

  return out;
}

function findLocalTestFilePath(fileName) {
  if (!fileName) return null;

  const wanted = normalizeNameForLocalMatch(fileName);

  for (const dir of LOCAL_TEST_DIRS) {
    if (!fs.existsSync(dir)) continue;

    // 1) exact path
    const directPath = path.join(dir, fileName);
    if (fs.existsSync(directPath)) {
      return directPath;
    }

    // 2) recursive basename match
    const allFiles = walkFilesRecursive(dir);

    const exact = allFiles.find((fullPath) => {
      const base = path.basename(fullPath);
      return normalizeNameForLocalMatch(base) === wanted;
    });
    if (exact) return exact;

    // 3) loose basename match
    const loose = allFiles.find((fullPath) => {
      const base = normalizeNameForLocalMatch(path.basename(fullPath));
      return base.includes(wanted) || wanted.includes(base);
    });
    if (loose) return loose;
  }

  return null;
}

async function loadAndPreprocessFromStorageIfPossible(user, fileName) {
  const logLP = shouldLogCache();
  if (logLP) console.log("[loadAndPreprocess] user?.id:", user?.id);
  if (logLP) console.log("[loadAndPreprocess] fileName:", fileName);

  if (!user || !fileName) {
    if (logLP) {
      console.log("[loadAndPreprocess] early return (no user/fileName)");
    }
    return { isFileAttached: false, preprocessed: null };
  }

  const isLocalDev = process.env.LOCAL_DEV === "1";
  if (isLocalDev && fileName) {
    const fallbackPath = findLocalTestFilePath(fileName);
    if (logLP) {
      console.log(
        "[loadAndPreprocess] local fallback resolved path:",
        fallbackPath,
      );
    }

    if (fallbackPath) {
      if (logLP) {
        console.log(
          "[loadAndPreprocess] using LOCAL_DEV fallback first:",
          fallbackPath,
        );
      }

      const buffer = fs.readFileSync(fallbackPath);
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
  }

  if (logLP)
    console.log("[loadAndPreprocess] user.uploadedFiles:", user?.uploadedFiles);
  let fileInfo = user.uploadedFiles?.find((f) => f.originalName === fileName);
  if (logLP) console.log("[loadAndPreprocess] fileInfo:", fileInfo);

  if (!fileInfo) {
    if (logLP)
      console.log("[loadAndPreprocess] early return (fileInfo not found)");
    return { isFileAttached: false, preprocessed: null };
  }

  const storageName = fileInfo.localName || fileInfo.gcsName;
  if (!storageName) {
    if (logLP) console.log("[loadAndPreprocess] early return (no storageName)");
    return { isFileAttached: false, preprocessed: null };
  }

  const buffer = await downloadToBuffer(storageName);
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
        "textjoin", "textsplit", "regexmatch", "regexreplace",
        "vstack", "tocol", "byrow".
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
  - If the user explicitly specifies A1-style ranges such as "A1:C3", prefer
    range-based operations instead of header-based hints.
  - Use "vstack" when the user asks to vertically combine multiple explicit ranges.
  - Use "tocol" when the user asks to flatten a range into a single column.
  - Use "byrow" when the user asks for row-wise calculations on an explicit range.
  - For "byrow", only use it when an explicit range is present. If no explicit
    range is given, omit range-specific fields rather than inventing a range.
  - The output MUST be valid JSON. No comments, no trailing commas, no extra text.
  - Prefer header-based hints (lookup_hint, return_hint, header_hint, group_by)
    instead of hard-coded cell references or ranges.
  - Do NOT invent sheet names, column letters, or A1-style ranges.
    Never output things like "Sheet1!B2:B100" or column letters like "A", "B", "C".

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

4) Vertical stack of explicit ranges
User: "A1:A3와 B1:B3를 세로로 합쳐줘"
Intent:
{
  "intent": {
    "operation": "vstack"
  }
}

5) Flatten range into one column
User: "A1:C3를 한 열로 펴줘"
Intent:
{
  "intent": {
    "operation": "tocol"
  }
}

6) Row-wise calculation on explicit range
User: "A1:C3의 각 행의 합계를 구해줘"
Intent:
{
  "intent": {
    "operation": "byrow"
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
      const out = `=ERROR("요청 정보가 부족합니다.")`;
      return sendFormulaResponse(res.status(400), {
        excelFormula: out,
        sheetsFormula: out,
        compatibility: detectFormulaCompatibility(out),
        debugMeta: null,
      });
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
      await loadAndPreprocessFromStorageIfPossible(req.user, fileName);
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
    intent.raw_message = message;
    intent = normalizeLookupIntent(intent);
    intent = applyStructuralOverrides(intent);
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
          const out = `=ERROR("필요한 열을 파일에서 찾을 수 없습니다.")`;
          const debugMeta = buildDebugMeta({
            rawReason: "MISSING_REQUIRED_COLUMN",
            cacheHit: _dbgCacheHit,
            intentOp: intent?.operation,
            intentCacheKey: _dbgIntentCacheKey,
            intentVersion: INTENT_VERSION,
            clusterVersion: CLUSTER_VERSION,
            resolverMode: getResolverMode(context),
            validator: null,
            timing: {
              preprocess: _ms(_tPreStart, _tPreEnd),
              intent: _ms(_tIntentStart, _tIntentEnd),
              build: _ms(_tBuildStart, _tBuildEnd),
              total: Date.now() - startedAt,
            },
            extra: {
              compatibilityLevel:
                detectFormulaCompatibility(out)?.level || null,
              fallbackAttempted: false,
              fallbackFunctions: [],
              resolvedBaseSheet: context?.resolved?.baseSheet || null,
              resolvedReturnHeaders: [],
              resolvedLookupHeader: null,
              resolvedGroupHeader: null,
            },
          });

          return sendFormulaResponse(res, {
            excelFormula: out,
            sheetsFormula: out,
            compatibility: detectFormulaCompatibility(out),
            debugMeta,
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
      direct?.canHandleWithoutFile?.(intent) &&
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
        const debugMeta = buildDebugMeta({
          rawReason,
          cacheHit: _dbgCacheHit,
          intentOp: intent?.operation,
          intentCacheKey: _dbgIntentCacheKey,
          intentVersion: INTENT_VERSION,
          clusterVersion: CLUSTER_VERSION,
          resolverMode: getResolverMode(context),
          validator: v,
          timing: {
            preprocess: _ms(_tPreStart, _tPreEnd),
            intent: _ms(_tIntentStart, _tIntentEnd),
            build: _ms(_tBuildStart, _tBuildEnd),
            total: Date.now() - startedAt,
          },
          extra: {
            compatibility: directCompatibility,
            compatibilityLevel: directCompatibility?.level || null,
            fallbackAttempted: false,
            fallbackFunctions: [],
            resolvedBaseSheet: context?.resolved?.baseSheet || null,
            resolvedReturnHeaders: getDebugReturnHeaders(context),
            resolvedLookupHeader:
              context?.resolved?.lookupColumn?.header || null,
            resolvedGroupHeader: context?.resolved?.groupColumn?.header || null,
          },
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
          debugMeta,
        });
        return sendFormulaResponse(res, {
          excelFormula: safeOut,
          sheetsFormula: safeOut,
          compatibility: directCompatibility,
          debugMeta,
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
    let safeFinal = v.ok
      ? finalFormula
      : `=ERROR("결과 검증에 실패했습니다. 입력을 더 구체적으로 작성해 주세요.")`;
    const baseFormula = safeFinal;

    let compatibility = detectFormulaCompatibility(safeFinal || "");
    let fallbackFunctions = [];

    if (v.ok && shouldAttemptCompatibilityFallback(compatibility)) {
      const fallback = tryGenerateFallbackFormula(safeFinal, compatibility);
      if (fallback?.formula && fallback.formula !== safeFinal) {
        const fallbackValidation = validateFormula(fallback.formula);
        if (fallbackValidation.ok) {
          const fallbackCompatibility = detectFormulaCompatibility(
            fallback.formula || "",
          );
          const improved =
            fallbackCompatibility.level === "common" ||
            (compatibility.level !== "common" &&
              fallbackCompatibility.blockers.length <
                compatibility.blockers.length);

          if (improved) {
            safeFinal = fallback.formula;
            compatibility = fallbackCompatibility;
            fallbackFunctions = fallback.appliedFunctions || [];
          }
        }
      }
    }

    _dbgCompatibility = compatibility;

    const debugMeta = buildDebugMeta({
      rawReason,
      cacheHit: _dbgCacheHit,
      intentOp: intent?.operation,
      intentCacheKey: _dbgIntentCacheKey,
      intentVersion: INTENT_VERSION,
      clusterVersion: CLUSTER_VERSION,
      resolverMode: getResolverMode(context),
      validator: v,
      timing: {
        preprocess: _ms(_tPreStart, _tPreEnd),
        intent: _ms(_tIntentStart, _tIntentEnd),
        build: _ms(_tBuildStart, _tBuildEnd),
        total: Date.now() - startedAt,
      },
      extra: {
        compatibility,
        compatibilityLevel: compatibility?.level || null,
        fallbackAttempted:
          v.ok && shouldAttemptCompatibilityFallback(compatibility),
        fallbackFunctions,
        resolvedBaseSheet: context?.resolved?.baseSheet || null,
        resolvedReturnHeaders: getDebugReturnHeaders(context),
        resolvedLookupHeader: context?.resolved?.lookupColumn?.header || null,
        resolvedGroupHeader: context?.resolved?.groupColumn?.header || null,
      },
    });

    await writeRequestLog({
      traceId,
      userId: req.user?.id,
      route: "/convert",
      engine: "formula",
      status: shouldCountConversion(safeFinal) ? "success" : "fail",
      reason: reasonNorm,
      isFallback: v.ok ? fallbackFunctions.length > 0 : true,
      prompt: message,
      latencyMs: Date.now() - startedAt,
      debugMeta,
    });
    const finalCompatibility = detectFormulaCompatibility(
      safeFinal || finalFormula || "",
    );
    const excelFormula = baseFormula;
    const sheetsFormula =
      fallbackFunctions.length > 0 ? safeFinal : baseFormula;

    return sendFormulaResponse(res, {
      excelFormula,
      sheetsFormula,
      compatibility: finalCompatibility,
      debugMeta,
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
  let intent = normalizeIntentSchema(baseIntent, nl);
  intent = normalizeLookupIntent(intent);
  intent = applyStructuralOverrides(intent);
  intent.raw_message = nl;

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
