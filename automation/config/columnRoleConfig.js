const COLUMN_ROLE_PATTERNS = Object.freeze({
  date: /date|일자|날짜|기간|월|연도|year|month/i,
  id: /id$|_id$|번호|코드|식별자/i,
  status: /상태|status|구분|여부|단계|분류/i,
  metric:
    /금액|매출|비용|수익|연봉|급여|점수|수량|가격|단가|평균|합계|율|rate|amount|price|salary|score|count/i,
});

const BOOLEAN_VALUES = Object.freeze([
  "true",
  "false",
  "yes",
  "no",
  "y",
  "n",
  "예",
  "아니오",
]);

const COLUMN_INFERENCE_THRESHOLDS = Object.freeze({
  sampleSize: 50,
  dateRatio: 0.7,
  numberRatio: 0.7,
  booleanRatio: 0.7,

  emptyRatioWarning: 0.45,
  headerConfidenceWarning: 0.6,
  typeConsistencyWarning: 0.5,

  confidenceWeights: Object.freeze({
    headerConfidence: 0.45,
    typeConsistency: 0.35,
    nonEmptyRatio: 0.2,
  }),
});

module.exports = {
  COLUMN_ROLE_PATTERNS,
  BOOLEAN_VALUES,
  COLUMN_INFERENCE_THRESHOLDS,
};
