const FORMULA_HEURISTICS = Object.freeze({
  simpleAggregateOperations: ["sum", "count", "average"],

  aggregateOperationKeywords: Object.freeze({
    average: ["average", "avg", "mean", "평균"],
    count: ["count", "건수", "개수", "인원", "대상 수", "전체 대상 수"],
    sum: ["sum", "total", "amount", "금액", "합계", "매출", "집행", "수량"],
  }),

  rankingSkipPattern: /top|bottom|ranking|rank|상위|하위|순위/i,

  criteriaHeaderFallbacks: ["기준", "그룹", "구분", "분류"],

  aggregateTargetHeaders: Object.freeze({
    count: ["count", "COUNT", "건수", "개수", "인원수", "대상 수", "행수"],
    sum: ["sum", "SUM", "합계", "총합", "금액합계", "수량합계"],
    average: ["average", "avg", "AVG", "평균"],
    value: ["값"],
  }),

  valueHeaderFallbacks: [
    "값",
    "합계",
    "평균",
    "금액",
    "집행금액",
    "순매출액",
    "매출수량",
    "인원수",
    "건수",
  ],

  ignoredNumericHeadersPattern: /행수|rowCount|작업|지표/i,

  numericColumnPattern:
    /number|numeric|amount|metric|value|금액|합계|평균|수량|매출|집행|연봉|건수|count|sum|average/i,

  headerMatch: Object.freeze({
    minScore: 60,
    numericBonus: 5,
  }),
});

module.exports = {
  FORMULA_HEURISTICS,
};
