const { REASONS } = require("./reasonClassifier");

/**
 * reason별 "우선 수정 액션" 가이드(관리자 요약에 포함)
 * - 5-B 핵심: 무엇부터 고칠지 자동으로 보이게
 */
const RECOMMENDATIONS = {
  [REASONS.MISSING_INPUT]: {
    title: "입력 검증 강화",
    next: ["프론트/백에서 prompt/message 필수 검증", "에러 메시지 통일"],
  },
  [REASONS.LIMIT_EXCEEDED]: {
    title: "플랜/사용량 UX 개선",
    next: [
      "한도 초과 안내 문구/업그레이드 CTA",
      "관리자에서 제한 발생 비율 모니터링",
    ],
  },
  [REASONS.UNSUPPORTED_OPERATION]: {
    title: "미지원 intent/operation 처리",
    next: [
      "operation whitelist/alias 확장",
      "미지원 시 추가 질문 유도(6단계에서 강화)",
    ],
  },
  [REASONS.VALIDATION_ERROR]: {
    title: "열 매핑/조건 모호성 해결",
    next: [
      "best column ambiguous 시 follow-up 템플릿",
      "헤더 후보 Top2 노출(6-2 debugMeta)",
    ],
  },
  [REASONS.FILE_ERROR]: {
    title: "파일 전처리 안정화",
    next: [
      "GCS 다운로드/파싱 실패율 확인",
      "파일명 매칭/업로드Files 메타 점검",
    ],
  },
  [REASONS.INTERNAL_ERROR]: {
    title: "서버 예외 수집/재현",
    next: ["traceId 기반 상세 로그(6-2) 확장", "에러 타입별 재시도/폴백 정의"],
  },
  [REASONS.UNKNOWN]: {
    title: "분류 규칙 보강",
    next: ["unknown 샘플 Top N 보고 규칙 추가", "reason raw 값 보존"],
  },
  [REASONS.OK]: {
    title: "성공 케이스 유지",
    next: ["성공 샘플로 few-shot/피드백 연결(후속)"],
  },
};

function getRecommendation(reason) {
  return RECOMMENDATIONS[reason] || RECOMMENDATIONS[REASONS.UNKNOWN];
}

module.exports = { getRecommendation };
