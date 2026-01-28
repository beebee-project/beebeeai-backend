/**
 * 표준 reason 목록 (단일값)
 * - 나중에 더 늘려도 됨. 지금은 운영에 바로 필요한 것만.
 */
const REASONS = {
  OK: "OK",
  MISSING_INPUT: "MISSING_INPUT",
  LIMIT_EXCEEDED: "LIMIT_EXCEEDED",
  UNSUPPORTED_OPERATION: "UNSUPPORTED_OPERATION",
  VALIDATION_ERROR: "VALIDATION_ERROR",
  MODEL_OR_INTENT_ERROR: "MODEL_OR_INTENT_ERROR",
  FILE_ERROR: "FILE_ERROR",
  CACHE: "CACHE",
  INTERNAL_ERROR: "INTERNAL_ERROR",
  UNKNOWN: "UNKNOWN",
};

function normalizeText(x) {
  return String(x ?? "").trim();
}

function classifyFromResult({ result }) {
  const r = normalizeText(result);
  if (!r) return null;

  // convert: "=ERROR(" 로 내려오는 케이스들
  if (/^=ERROR\s*\(/i.test(r)) {
    // 구체 메시지 기반 분류(운영에 도움되는 것만)
    if (r.includes("지원하지 않는 작업")) return REASONS.UNSUPPORTED_OPERATION;
    if (r.includes("모호합니다")) return REASONS.VALIDATION_ERROR;
    if (r.includes("열을 파일에서 찾을 수 없습니다")) return REASONS.FILE_ERROR;
    return REASONS.VALIDATION_ERROR;
  }

  return null;
}

/**
 * 컨텍스트 기반 정규화 (컨트롤러/서비스에서 호출)
 *
 * @param {object} args
 * @param {string} args.reason - 기존 reason 코드(있으면)
 * @param {string} [args.route]
 * @param {string} [args.engine]
 * @param {string} [args.prompt]
 * @param {string} [args.result] - convert의 result/finalFormula 등
 * @param {any} [args.error] - catch(e)
 */
function classifyReason(args = {}) {
  const rawReason = normalizeText(args.reason).toUpperCase();
  const prompt = normalizeText(args.prompt);
  const result = normalizeText(args.result);

  // 1) 입력 부족
  if (
    !prompt &&
    (rawReason === "MISSING_PROMPT" || rawReason === "MISSING_MESSAGE")
  ) {
    return REASONS.MISSING_INPUT;
  }
  if (!prompt && !rawReason) {
    return REASONS.MISSING_INPUT;
  }

  // 2) 사용량 제한
  if (rawReason === "LIMIT_EXCEEDED") return REASONS.LIMIT_EXCEEDED;

  // 3) 명시적 미지원
  if (rawReason === "UNSUPPORTED_MACRO") return REASONS.UNSUPPORTED_OPERATION;

  // 4) convert에서 ERROR_FORMULA / UNKNOWN일 때 result 메시지로 더 세분
  if (rawReason === "ERROR_FORMULA" || rawReason === "UNKNOWN" || !rawReason) {
    const fromResult = classifyFromResult({ result });
    if (fromResult) return fromResult;
  }

  // 5) 예외/서버 에러
  if (rawReason === "EXCEPTION" || rawReason === "MACRO_FAILED") {
    // file 관련 힌트
    const msg = normalizeText(args.error?.message);
    if (/download|bucket|gcs|storage|xlsx|buffer/i.test(msg))
      return REASONS.FILE_ERROR;
    return REASONS.INTERNAL_ERROR;
  }

  // 6) 정상
  if (rawReason === "OK") return REASONS.OK;

  // 7) 나머지
  return REASONS.UNKNOWN;
}

module.exports = { REASONS, classifyReason };
