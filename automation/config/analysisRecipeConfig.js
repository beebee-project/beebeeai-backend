const ANALYSIS_RECIPE_TYPES = Object.freeze({
  GROUP_SUMMARY: "group_summary",
  GROUP_SUM: "group_sum",
  GROUP_AVG: "group_avg",
  GROUP_COUNT: "group_count",
  CATEGORY_COUNT: "category_count",
  STATUS_COUNT: "status_count",
  TOP_BOTTOM: "top_bottom",
  COMPOSITION_RATIO: "composition_ratio",
  TIME_TREND: "time_trend",
  TIME_SUM: "time_sum",
  TIME_AVG: "time_avg",
  TIME_COUNT: "time_count",
  TIME_GROWTH: "time_growth",
  CUMULATIVE_SUM: "cumulative_sum",
  CROSS_COUNT: "cross_count",
  CROSS_SUM: "cross_sum",
  WIDE_TIME_TREND: "wide_time_trend",
});

const ANALYSIS_RECIPE_DEFS = Object.freeze([
  {
    recipeType: ANALYSIS_RECIPE_TYPES.GROUP_SUMMARY,
    titleTemplate: ({ dimension, metric }) =>
      `${dimension.header}별 ${metric.header} 요약`,
    descriptionTemplate: ({ dimension, metric }) =>
      `${dimension.header} 기준으로 ${metric.header}의 합계, 평균, 건수를 분석합니다.`,
    categoryId: "summary",
    requires: ["dimension", "metric"],
    operation: "summary",
    priority: 920,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.GROUP_SUM,
    titleTemplate: ({ dimension, metric }) =>
      `${dimension.header}별 ${metric.header} 합계`,
    descriptionTemplate: ({ dimension, metric }) =>
      `${dimension.header} 기준으로 ${metric.header} 합계를 집계합니다.`,
    categoryId: "summary",
    requires: ["dimension", "metric"],
    operation: "sum",
    priority: 880,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.GROUP_AVG,
    titleTemplate: ({ dimension, metric }) =>
      `${dimension.header}별 ${metric.header} 평균`,
    descriptionTemplate: ({ dimension, metric }) =>
      `${dimension.header} 기준으로 ${metric.header} 평균을 집계합니다.`,
    categoryId: "summary",
    requires: ["dimension", "metric"],
    operation: "average",
    priority: 840,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT,
    titleTemplate: ({ dimension }) => `${dimension.header}별 건수`,
    descriptionTemplate: ({ dimension }) =>
      `${dimension.header} 기준으로 데이터 건수를 집계합니다.`,
    categoryId: "summary",
    requires: ["dimension"],
    operation: "count",
    priority: 800,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TOP_BOTTOM,
    titleTemplate: ({ metric }) => `${metric.header} 상위/하위 항목`,
    descriptionTemplate: ({ metric, dimension }) =>
      `${metric.header} 기준으로 ${dimension?.header || "항목"}의 상위·하위 항목을 확인합니다.`,
    categoryId: "ranking",
    requires: ["metric"],
    operation: "rank",
    priority: 760,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.COMPOSITION_RATIO,
    titleTemplate: ({ dimension, metric }) =>
      metric
        ? `${dimension.header}별 ${metric.header} 구성비`
        : `${dimension.header}별 구성비`,
    descriptionTemplate: ({ dimension, metric }) =>
      metric
        ? `${dimension.header} 기준 ${metric.header} 합계의 구성비를 계산합니다.`
        : `${dimension.header} 기준 건수 구성비를 계산합니다.`,
    categoryId: "composition",
    requires: ["dimension"],
    operation: "ratio",
    priority: 720,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TIME_TREND,
    titleTemplate: ({ date, metric }) =>
      `${date.header} 기준 ${metric.header} 추이`,
    descriptionTemplate: ({ date, metric }) =>
      `${date.header}를 기준으로 ${metric.header}의 기간별 흐름을 분석합니다.`,
    categoryId: "trend",
    requires: ["date", "metric"],
    operation: "timeSummary",
    priority: 950,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TIME_SUM,
    titleTemplate: ({ date, metric }) =>
      `${date.header}별 ${metric.header} 합계 추이`,
    descriptionTemplate: ({ date, metric }) =>
      `${date.header} 기준 ${metric.header} 합계 추이를 생성합니다.`,
    categoryId: "trend",
    requires: ["date", "metric"],
    operation: "sum",
    priority: 910,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TIME_AVG,
    titleTemplate: ({ date, metric }) =>
      `${date.header}별 ${metric.header} 평균 추이`,
    descriptionTemplate: ({ date, metric }) =>
      `${date.header} 기준 ${metric.header} 평균 추이를 생성합니다.`,
    categoryId: "trend",
    requires: ["date", "metric"],
    operation: "average",
    priority: 870,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TIME_COUNT,
    titleTemplate: ({ date }) => `${date.header}별 건수 추이`,
    descriptionTemplate: ({ date }) =>
      `${date.header} 기준 데이터 건수 추이를 생성합니다.`,
    categoryId: "trend",
    requires: ["date"],
    operation: "count",
    priority: 820,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TIME_GROWTH,
    titleTemplate: ({ date, metric }) =>
      `${date.header}별 ${metric.header} 증감률`,
    descriptionTemplate: ({ date, metric }) =>
      `${date.header} 기준 ${metric.header}의 전기 대비 증감률을 계산합니다.`,
    categoryId: "trend",
    requires: ["date", "metric"],
    operation: "growthRate",
    priority: 790,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.CUMULATIVE_SUM,
    titleTemplate: ({ date, metric }) =>
      `${date.header}별 ${metric.header} 누적합계`,
    descriptionTemplate: ({ date, metric }) =>
      `${date.header} 기준 ${metric.header} 누적합계를 계산합니다.`,
    categoryId: "trend",
    requires: ["date", "metric"],
    operation: "cumulativeSum",
    priority: 750,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.CROSS_COUNT,
    titleTemplate: ({ dimension, dimension2 }) =>
      `${dimension.header} × ${dimension2.header} 교차 건수`,
    descriptionTemplate: ({ dimension, dimension2 }) =>
      `${dimension.header}와 ${dimension2.header}를 교차하여 건수를 집계합니다.`,
    categoryId: "cross_summary",
    requires: ["dimension", "dimension2"],
    operation: "crossCount",
    priority: 700,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.CROSS_SUM,
    titleTemplate: ({ dimension, dimension2, metric }) =>
      `${dimension.header} × ${dimension2.header} ${metric.header} 합계`,
    descriptionTemplate: ({ dimension, dimension2, metric }) =>
      `${dimension.header}와 ${dimension2.header}를 교차하여 ${metric.header} 합계를 집계합니다.`,
    categoryId: "cross_summary",
    requires: ["dimension", "dimension2", "metric"],
    operation: "crossSum",
    priority: 680,
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.WIDE_TIME_TREND,
    titleTemplate: ({ date, metric }) =>
      `${date.header} 기준 ${metric.header} 가로형 추이`,
    descriptionTemplate: ({ date, metric }) =>
      `가로형 기간 컬럼을 정규화한 뒤 ${date.header} 기준 ${metric.header} 추이를 분석합니다.`,
    categoryId: "trend",
    requires: ["date", "metric"],
    operation: "wideTimeTrend",
    priority: 930,
    virtualOnly: true,
  },
]);

// Backward compatibility: older builder imports ANALYSIS_RECIPE_RULES.
const ANALYSIS_RECIPE_RULES = [
  {
    recipeType: ANALYSIS_RECIPE_TYPES.GROUP_SUMMARY,
    required: ["primaryDimension", "primaryMetric"],
    build: ({ table, primaryDimension, primaryMetric, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_group_summary`,
        title: `${primaryDimension.header}별 ${primaryMetric.header} 요약`,
        description: `${primaryDimension.header} 기준으로 ${primaryMetric.header}의 합계, 평균, 건수를 분석합니다.`,
        recipeType: ANALYSIS_RECIPE_TYPES.GROUP_SUMMARY,
        table,
        metric: primaryMetric,
        dimension: primaryDimension,
      }),
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TIME_TREND,
    required: ["primaryDate", "primaryMetric"],
    build: ({ table, primaryDate, primaryMetric, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_time_trend`,
        title: `${primaryDate.header} 기준 ${primaryMetric.header} 추이`,
        description: `${primaryDate.header}를 기준으로 ${primaryMetric.header}의 기간별 흐름을 분석합니다.`,
        recipeType: ANALYSIS_RECIPE_TYPES.TIME_TREND,
        table,
        metric: primaryMetric,
        date: primaryDate,
      }),
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT,
    required: ["primaryDimension"],
    build: ({ table, primaryDimension, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_category_count`,
        title: `${primaryDimension.header}별 건수`,
        description: `${primaryDimension.header} 기준으로 데이터 건수를 집계합니다.`,
        recipeType: ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT,
        table,
        dimension: primaryDimension,
      }),
  },
  {
    recipeType: ANALYSIS_RECIPE_TYPES.TOP_BOTTOM,
    required: ["primaryMetric"],
    build: ({ table, primaryDimension, primaryMetric, makeCandidate }) =>
      makeCandidate({
        id: `${table.tableId}_top_bottom`,
        title: `${primaryMetric.header} 상위/하위 항목`,
        description: `${primaryMetric.header} 기준으로 높은 항목과 낮은 항목을 확인합니다.`,
        recipeType: ANALYSIS_RECIPE_TYPES.TOP_BOTTOM,
        table,
        metric: primaryMetric,
        dimension: primaryDimension,
      }),
  },
];

const ANALYSIS_RECIPE_OPTIONS = Object.freeze({
  minTableConfidence: 0.45,
  maxCandidatesPerTable: 24,
  maxMetricsPerTable: 3,
  maxDimensionsPerTable: 4,
  maxDatesPerTable: 2,
  maxDimensionPairsPerTable: 3,
  categoryTypesForDimensionFallback: ["category", "text", "string"],
  excludedRolesForDimensionFallback: ["metric", "date", "id"],
  metricHeaderHints: [
    "금액",
    "매출",
    "수량",
    "집행",
    "예산",
    "연봉",
    "급여",
    "비용",
    "합계",
    "평균",
    "율",
    "비율",
    "건수",
    "인원",
    "명",
    "value",
    "amount",
    "sales",
    "revenue",
    "count",
  ],
  dateHeaderHints: [
    "일자",
    "날짜",
    "기간",
    "연도",
    "년도",
    "월",
    "연월",
    "분기",
    "date",
    "year",
    "month",
    "period",
  ],
  idHeaderHints: ["id", "번호", "코드", "식별", "고유", "순번"],
  statusHeaderHints: [
    "상태",
    "구분",
    "등급",
    "분류",
    "status",
    "grade",
    "type",
  ],
  semanticColumnPriorityVersion: "general_analysis_column_priority_v1",
  semanticColumnPriorityGroups: {
    dimension: [
      {
        score: 110,
        hints: [
          "부서",
          "지점",
          "지사",
          "영업점",
          "점포",
          "매장",
          "팀",
          "본부",
          "센터",
          "조직",
          "사업부",
          "department",
          "branch",
          "store",
          "team",
          "division",
          "center",
          "organization",
        ],
      },
      {
        score: 80,
        hints: [
          "kpi",
          "지표",
          "품목",
          "상품",
          "제품",
          "프로젝트",
          "과제",
          "사용유형",
          "비목",
          "카테고리",
          "지역",
          "권역",
          "시설",
          "창고",
          "item",
          "product",
          "project",
          "category",
          "region",
          "facility",
          "warehouse",
        ],
      },
      {
        score: 30,
        hints: [
          "거래처",
          "가맹점",
          "사용자",
          "카드명",
          "카드구분",
          "merchant",
          "vendor",
          "customer",
          "user",
          "card",
        ],
      },
    ],
    metric: [
      {
        score: 110,
        hints: ["실적", "실제", "성과", "actual", "performance"],
      },
      {
        score: 100,
        hints: [
          "매출",
          "판매금액",
          "사용금액",
          "금액",
          "집행액",
          "사용량",
          "현재재고",
          "판매수량",
          "영업이익",
          "revenue",
          "sales",
          "amount",
          "usage",
          "currentstock",
          "profit",
        ],
      },
      {
        score: 80,
        hints: [
          "달성률",
          "달성율",
          "이익률",
          "증감률",
          "증가율",
          "감소율",
          "비율",
          "점유율",
          "rate",
          "ratio",
          "margin",
          "percent",
        ],
      },
      {
        score: 50,
        hints: [
          "목표",
          "계획",
          "예산",
          "기준값",
          "target",
          "plan",
          "budget",
          "baseline",
        ],
      },
      {
        score: 20,
        hints: ["전월재고", "기초재고", "단가", "unitprice", "openingstock"],
      },
    ],
    date: [
      {
        score: 110,
        hints: ["연월", "기준월", "yearmonth"],
      },
      {
        score: 100,
        hints: [
          "월",
          "분기",
          "기간",
          "연도",
          "년도",
          "month",
          "quarter",
          "period",
          "year",
        ],
      },
      {
        score: 60,
        hints: ["승인일", "사용일", "일자", "날짜", "date", "day"],
      },
    ],
  },
});

function hasRequiredRecipeContext(rule = {}, context = {}) {
  return (rule.required || []).every((key) => Boolean(context[key]));
}

module.exports = {
  ANALYSIS_RECIPE_TYPES,
  ANALYSIS_RECIPE_DEFS,
  ANALYSIS_RECIPE_RULES,
  ANALYSIS_RECIPE_OPTIONS,
  hasRequiredRecipeContext,
};
