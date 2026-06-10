const COLUMN_ROLE_PATTERNS = {
  date: /date|일자|날짜|기간|월|연도|year|month/i,
  id: /id$|_id$|번호|코드|식별자/i,
  status: /상태|status|구분|여부|단계|분류/i,
  metric:
    /금액|매출|비용|수익|연봉|급여|점수|수량|가격|단가|평균|합계|율|rate|amount|price|salary|score|count/i,
};

const BOOLEAN_VALUES = ["true", "false", "yes", "no", "y", "n", "예", "아니오"];

module.exports = {
  COLUMN_ROLE_PATTERNS,
  BOOLEAN_VALUES,
};
