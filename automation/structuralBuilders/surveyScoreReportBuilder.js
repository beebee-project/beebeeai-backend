const {
  findColumnHeader,
  getColumnHeader,
  getColumns,
  makeTemplateCandidate,
  makeTemplateSection,
  executeTemplateSections,
  getRows,
  getRowValue,
  toNumber,
} = require("../businessTemplates/commonTemplateHelpers");

const SURVEY_SCORE_REPORT_VERSION = "survey_score_report_builder_v1";

const DEFAULT_SCORE_HINTS = [
  "점수",
  "평점",
  "평가점수",
  "만족도",
  "만족점수",
  "응답점수",
  "척도",
  "별점",
  "추천점수",
  "재추천",
  "nps",
  "score",
  "rating",
  "satisfaction",
  "evaluation",
];

const DEFAULT_QUESTION_HINTS = [
  "문항",
  "질문",
  "항목",
  "평가항목",
  "설문항목",
  "질문내용",
  "문항명",
  "question",
  "item",
  "criteria",
];

const DEFAULT_RESPONDENT_HINTS = [
  "응답자",
  "참여자",
  "수강생",
  "교육생",
  "고객",
  "직원",
  "성명",
  "이름",
  "respondent",
  "participant",
  "customer",
  "employee",
  "name",
];

const DEFAULT_DEPARTMENT_HINTS = [
  "부서",
  "소속",
  "조직",
  "팀",
  "기관",
  "학과",
  "department",
  "team",
  "organization",
];

const DEFAULT_CATEGORY_HINTS = [
  "구분",
  "분류",
  "유형",
  "카테고리",
  "과정",
  "교육과정",
  "행사명",
  "프로그램",
  "서비스",
  "강사",
  "지점",
  "채널",
  "category",
  "type",
  "course",
  "program",
  "service",
  "instructor",
];

const DEFAULT_DATE_HINTS = [
  "일자",
  "날짜",
  "월",
  "연월",
  "기준월",
  "응답일",
  "설문일",
  "평가일",
  "교육일",
  "행사일",
  "date",
  "month",
  "period",
];

const DEFAULT_COMMENT_HINTS = [
  "의견",
  "건의",
  "개선사항",
  "비고",
  "서술",
  "주관식",
  "comment",
  "feedback",
  "opinion",
];

const NON_SCORE_HINTS = [
  "id",
  "번호",
  "순번",
  "코드",
  "연도",
  "년도",
  "월",
  "일자",
  "날짜",
  "금액",
  "비용",
  "단가",
  "수량",
  "건수",
  "매출",
  "집행",
  "예산",
  "amount",
  "cost",
  "price",
  "quantity",
  "count",
];

function normalizeText(value = "") {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()\[\]{}]/g, "")
    .trim();
}

function includesAny(text = "", hints = []) {
  const normalized = normalizeText(text);
  return hints.some((hint) => {
    const target = normalizeText(hint);
    return target && normalized.includes(target);
  });
}

function safeRate(numerator = 0, denominator = 0) {
  const n = Number(numerator || 0);
  const d = Number(denominator || 0);
  if (!Number.isFinite(n) || !Number.isFinite(d) || d === 0) return null;
  return n / d;
}

function makePercent(value) {
  return value == null ? null : value * 100;
}

function average(values = []) {
  const nums = values.filter((value) => Number.isFinite(Number(value)));
  if (!nums.length) return null;
  return nums.reduce((sum, value) => sum + Number(value), 0) / nums.length;
}

function minNumber(values = []) {
  const nums = values.filter((value) => Number.isFinite(Number(value)));
  return nums.length ? Math.min(...nums) : null;
}

function maxNumber(values = []) {
  const nums = values.filter((value) => Number.isFinite(Number(value)));
  return nums.length ? Math.max(...nums) : null;
}

function detectScaleMax(values = []) {
  const max = maxNumber(values);
  if (max == null) return 5;
  if (max <= 5) return 5;
  if (max <= 10) return 10;
  if (max <= 100) return 100;
  return max;
}

function positiveThreshold(scaleMax = 5) {
  if (scaleMax <= 5) return 4;
  if (scaleMax <= 10) return 8;
  if (scaleMax <= 100) return 80;
  return scaleMax * 0.8;
}

function scoreValuesForHeader(table = {}, header = "") {
  return getRows(table)
    .map((row) => toNumber(getRowValue(row, header)))
    .filter((value) => value != null && Number.isFinite(Number(value)));
}

function isLikelyScoreHeader({ table = {}, header = "", scoreHints = [] } = {}) {
  if (!header) return false;
  const normalizedHeader = normalizeText(header);
  const hasScoreHint = includesAny(header, [...scoreHints, ...DEFAULT_SCORE_HINTS]);
  const hasNonScoreHint = includesAny(header, NON_SCORE_HINTS);

  if (hasNonScoreHint && !hasScoreHint) return false;

  const values = scoreValuesForHeader(table, header);
  if (values.length < 2) return false;

  const min = minNumber(values);
  const max = maxNumber(values);
  if (min == null || max == null) return false;

  const looksLikeScale = min >= 0 && max <= 100;
  if (hasScoreHint && looksLikeScale) return true;

  // 숫자형 컬럼이지만 명시 힌트가 없는 경우는 너무 넓게 잡지 않는다.
  return (
    !hasNonScoreHint &&
    looksLikeScale &&
    values.length >= 3 &&
    /score|rating|satisfaction|평가|만족|점수|평점|추천/.test(normalizedHeader)
  );
}

function findScoreHeaders(table = {}, hints = []) {
  const columns = getColumns(table);
  const headers = columns
    .map(getColumnHeader)
    .filter(Boolean)
    .filter((header) => isLikelyScoreHeader({ table, header, scoreHints: hints }));

  return [...new Set(headers)].slice(0, 12);
}

function findSurveyScoreHeaders(table = {}, config = {}) {
  const hints = config.hints || {};
  const scoreHeaders = findScoreHeaders(table, hints.score || []);

  const questionHeader = findColumnHeader(table, [
    ...(hints.question || []),
    ...DEFAULT_QUESTION_HINTS,
  ]);

  const dateHeader = findColumnHeader(table, [
    ...(hints.date || []),
    ...DEFAULT_DATE_HINTS,
  ]);

  const respondentHeader = findColumnHeader(table, [
    ...(hints.respondent || []),
    ...DEFAULT_RESPONDENT_HINTS,
  ]);

  const departmentHeader = findColumnHeader(table, [
    ...(hints.department || []),
    ...DEFAULT_DEPARTMENT_HINTS,
  ]);

  const categoryHeader = findColumnHeader(table, [
    ...(hints.category || []),
    ...DEFAULT_CATEGORY_HINTS,
  ]);

  const commentHeader = findColumnHeader(table, [
    ...(hints.comment || []),
    ...DEFAULT_COMMENT_HINTS,
  ]);

  const npsHeader =
    scoreHeaders.find((header) => includesAny(header, ["nps", "추천", "재추천"])) ||
    "";

  return {
    scoreHeaders,
    primaryScoreHeader: scoreHeaders[0] || "",
    questionHeader,
    dateHeader,
    respondentHeader,
    departmentHeader,
    categoryHeader,
    commentHeader,
    npsHeader,
  };
}

function makeCustomSurveySection({
  sectionId,
  sectionType,
  title,
  table,
  rows,
  columns = {},
  chartHint = {},
  narrativeHint = {},
  meta = {},
}) {
  return makeTemplateSection({
    sectionId,
    sectionType,
    title,
    candidate: {
      recipeType: "custom_metric",
      title,
      tableId: table.tableId,
      columns,
      meta: {
        ...meta,
        surveyScoreReportVersion: SURVEY_SCORE_REPORT_VERSION,
      },
    },
    result: {
      ok: true,
      recipeType: "custom_metric",
      resultType: sectionType,
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
      meta: {
        ...meta,
        surveyScoreReportVersion: SURVEY_SCORE_REPORT_VERSION,
      },
    },
    chartHint,
    narrativeHint,
  });
}

function summarizeScore(table = {}, scoreHeader = "") {
  const values = scoreValuesForHeader(table, scoreHeader);
  const scaleMax = detectScaleMax(values);
  const threshold = positiveThreshold(scaleMax);
  const positiveCount = values.filter((value) => Number(value) >= threshold).length;

  return {
    scoreHeader,
    responseCount: values.length,
    averageScore: average(values),
    minScore: minNumber(values),
    maxScore: maxNumber(values),
    scaleMax,
    positiveCount,
    positiveRate: safeRate(positiveCount, values.length),
    positiveRatePercent: makePercent(safeRate(positiveCount, values.length)),
  };
}

function buildOverallScoreSection({ table, headers, config = {} }) {
  const scoreHeaders = headers.scoreHeaders || [];
  if (!table?.tableId || !scoreHeaders.length) return null;

  const rows = scoreHeaders
    .map((scoreHeader) => summarizeScore(table, scoreHeader))
    .filter((item) => item.responseCount > 0)
    .map((item) => ({
      지표: item.scoreHeader,
      응답수: item.responseCount,
      평균점수: item.averageScore,
      최저점: item.minScore,
      최고점: item.maxScore,
      척도상한: item.scaleMax,
      긍정응답수: item.positiveCount,
      긍정률: item.positiveRate,
      긍정률Percent: item.positiveRatePercent,
    }));

  if (!rows.length) return null;

  return makeCustomSurveySection({
    sectionId: config.sectionIds?.overview || "survey_score_overview",
    sectionType: "survey_score_overview",
    title: config.titles?.overview || "설문 점수 요약",
    table,
    rows,
    columns: {
      score: scoreHeaders,
      count: "응답수",
      average: "평균점수",
      positiveRate: "긍정률Percent",
    },
    chartHint: {
      preferredType: "metric_card",
      valueField: "평균점수",
      ratioField: "긍정률Percent",
    },
    narrativeHint: {
      focus: "survey_score_overview",
      scoreHeaders,
    },
  });
}

function buildQuestionAverageSection({ table, headers, config = {} }) {
  const { questionHeader, primaryScoreHeader, scoreHeaders } = headers || {};
  if (!table?.tableId || !scoreHeaders?.length) return null;

  let rows = [];

  if (questionHeader && primaryScoreHeader) {
    const map = new Map();
    getRows(table).forEach((row) => {
      const question = String(getRowValue(row, questionHeader) ?? "").trim() || "미입력";
      const score = toNumber(getRowValue(row, primaryScoreHeader));
      if (score == null) return;
      if (!map.has(question)) map.set(question, []);
      map.get(question).push(score);
    });

    rows = Array.from(map.entries()).map(([question, values]) => {
      const scaleMax = detectScaleMax(values);
      const threshold = positiveThreshold(scaleMax);
      const positiveCount = values.filter((value) => value >= threshold).length;
      return {
        [questionHeader]: question,
        응답수: values.length,
        평균점수: average(values),
        최저점: minNumber(values),
        최고점: maxNumber(values),
        긍정응답수: positiveCount,
        긍정률Percent: makePercent(safeRate(positiveCount, values.length)),
      };
    });
  } else {
    rows = scoreHeaders.map((scoreHeader) => {
      const summary = summarizeScore(table, scoreHeader);
      return {
        문항: scoreHeader,
        응답수: summary.responseCount,
        평균점수: summary.averageScore,
        최저점: summary.minScore,
        최고점: summary.maxScore,
        긍정응답수: summary.positiveCount,
        긍정률Percent: summary.positiveRatePercent,
      };
    });
  }

  rows = rows
    .filter((row) => Number(row.응답수 || 0) > 0)
    .sort((a, b) => Number(b.평균점수 || 0) - Number(a.평균점수 || 0));

  if (!rows.length) return null;

  const dimensionField = questionHeader || "문항";

  return makeCustomSurveySection({
    sectionId: config.sectionIds?.questionAverage || "survey_question_average",
    sectionType: "survey_question_average",
    title: config.titles?.questionAverage || "문항별 평균 점수",
    table,
    rows,
    columns: {
      question: dimensionField,
      score: primaryScoreHeader || scoreHeaders,
      count: "응답수",
      average: "평균점수",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionField,
      valueField: "평균점수",
    },
    narrativeHint: {
      focus: "question_average",
      question: dimensionField,
    },
  });
}

function buildScoreDistributionSection({ table, headers, config = {} }) {
  const scoreHeader = headers.primaryScoreHeader;
  if (!table?.tableId || !scoreHeader) return null;

  const values = scoreValuesForHeader(table, scoreHeader);
  if (!values.length) return null;

  const total = values.length;
  const map = new Map();

  values.forEach((value) => {
    const key = String(value);
    map.set(key, (map.get(key) || 0) + 1);
  });

  const rows = Array.from(map.entries())
    .map(([score, count]) => ({
      점수: score,
      응답수: count,
      비율: safeRate(count, total),
      비율Percent: makePercent(safeRate(count, total)),
    }))
    .sort((a, b) => Number(a.점수) - Number(b.점수));

  return makeCustomSurveySection({
    sectionId: config.sectionIds?.scoreDistribution || "survey_score_distribution",
    sectionType: "survey_score_distribution",
    title: config.titles?.scoreDistribution || `${scoreHeader} 응답 분포`,
    table,
    rows,
    columns: {
      score: scoreHeader,
      count: "응답수",
      ratio: "비율Percent",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: "점수",
      valueField: "응답수",
    },
    narrativeHint: {
      focus: "score_distribution",
      score: scoreHeader,
    },
  });
}

function buildDimensionScoreSection({
  table,
  headers,
  dimensionHeader = "",
  title = "",
  sectionId = "",
}) {
  const scoreHeader = headers.primaryScoreHeader;
  if (!table?.tableId || !scoreHeader || !dimensionHeader) return null;

  const map = new Map();
  getRows(table).forEach((row) => {
    const dimension = String(getRowValue(row, dimensionHeader) ?? "").trim() || "미입력";
    const score = toNumber(getRowValue(row, scoreHeader));
    if (score == null) return;
    if (!map.has(dimension)) map.set(dimension, []);
    map.get(dimension).push(score);
  });

  const rows = Array.from(map.entries())
    .map(([dimension, values]) => {
      const scaleMax = detectScaleMax(values);
      const threshold = positiveThreshold(scaleMax);
      const positiveCount = values.filter((value) => value >= threshold).length;
      return {
        [dimensionHeader]: dimension,
        응답수: values.length,
        평균점수: average(values),
        최저점: minNumber(values),
        최고점: maxNumber(values),
        긍정응답수: positiveCount,
        긍정률Percent: makePercent(safeRate(positiveCount, values.length)),
      };
    })
    .filter((row) => Number(row.응답수 || 0) > 0)
    .sort((a, b) => Number(b.평균점수 || 0) - Number(a.평균점수 || 0));

  if (!rows.length) return null;

  return makeCustomSurveySection({
    sectionId: sectionId || `survey_score_by_${dimensionHeader}`,
    sectionType: "survey_score_by_dimension",
    title: title || `${dimensionHeader}별 평균 점수`,
    table,
    rows,
    columns: {
      dimension: dimensionHeader,
      score: scoreHeader,
      count: "응답수",
      average: "평균점수",
      positiveRate: "긍정률Percent",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "평균점수",
    },
    narrativeHint: {
      focus: "survey_score_by_dimension",
      dimension: dimensionHeader,
      score: scoreHeader,
    },
  });
}

function buildNpsSection({ table, headers, config = {} }) {
  const npsHeader = headers.npsHeader;
  if (!table?.tableId || !npsHeader) return null;

  const values = scoreValuesForHeader(table, npsHeader).filter(
    (value) => value >= 0 && value <= 10,
  );
  if (!values.length) return null;

  const promoters = values.filter((value) => value >= 9).length;
  const passives = values.filter((value) => value >= 7 && value <= 8).length;
  const detractors = values.filter((value) => value <= 6).length;
  const total = values.length;
  const nps = makePercent(safeRate(promoters - detractors, total));

  const rows = [
    {
      지표: "전체 응답수",
      값: total,
      비율Percent: 100,
    },
    {
      지표: "추천 고객(Promoter)",
      값: promoters,
      비율Percent: makePercent(safeRate(promoters, total)),
    },
    {
      지표: "중립 고객(Passive)",
      값: passives,
      비율Percent: makePercent(safeRate(passives, total)),
    },
    {
      지표: "비추천 고객(Detractor)",
      값: detractors,
      비율Percent: makePercent(safeRate(detractors, total)),
    },
    {
      지표: "NPS",
      값: nps,
      비율Percent: nps,
    },
  ];

  return makeCustomSurveySection({
    sectionId: config.sectionIds?.nps || "survey_nps_summary",
    sectionType: "survey_nps_summary",
    title: config.titles?.nps || `${npsHeader} NPS 요약`,
    table,
    rows,
    columns: {
      score: npsHeader,
      value: "값",
      ratio: "비율Percent",
    },
    chartHint: {
      preferredType: "metric_card",
      valueField: "값",
      ratioField: "비율Percent",
    },
    narrativeHint: {
      focus: "nps_summary",
      score: npsHeader,
    },
  });
}

function buildSurveyScoreCandidates({ table, headers, config = {} }) {
  if (!table?.tableId) return [];

  const {
    primaryScoreHeader,
    scoreHeaders,
    questionHeader,
    dateHeader,
    departmentHeader,
    categoryHeader,
    respondentHeader,
  } = headers || {};

  if (!primaryScoreHeader) return [];

  const candidates = [];
  const tableId = table.tableId;

  const dimensions = [departmentHeader, categoryHeader, questionHeader]
    .filter(Boolean)
    .filter((value, index, arr) => arr.indexOf(value) === index);

  for (const dimensionHeader of dimensions) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: `avg_score_by_${dimensionHeader}`,
        sectionType: "survey_dimension_average",
        recipeType: "group_avg",
        title: `${dimensionHeader}별 ${primaryScoreHeader} 평균`,
        tableId,
        columns: {
          dimension: dimensionHeader,
          metric: primaryScoreHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: dimensionHeader,
          valueField: primaryScoreHeader,
        },
        narrativeHint: {
          focus: "group_avg",
          dimension: dimensionHeader,
          metric: primaryScoreHeader,
        },
      }),
    );

    candidates.push(
      makeTemplateCandidate({
        sectionId: `count_by_${dimensionHeader}`,
        sectionType: "survey_dimension_count",
        recipeType: "category_count",
        title: `${dimensionHeader}별 응답 수`,
        tableId,
        columns: {
          dimension: dimensionHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: dimensionHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "category_count",
          dimension: dimensionHeader,
        },
      }),
    );
  }

  if (dateHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: "survey_time_average",
        sectionType: "survey_time_average",
        recipeType: "time_avg",
        title: `${dateHeader}별 ${primaryScoreHeader} 평균 추이`,
        tableId,
        columns: {
          date: dateHeader,
          metric: primaryScoreHeader,
        },
        chartHint: {
          preferredType: "line",
          categoryField: dateHeader,
          valueField: primaryScoreHeader,
        },
        narrativeHint: {
          focus: "time_avg",
          date: dateHeader,
          metric: primaryScoreHeader,
        },
      }),
    );
  }

  if (categoryHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: "survey_category_composition",
        sectionType: "survey_category_composition",
        recipeType: "composition_ratio",
        title: `${categoryHeader} 응답 구성비`,
        tableId,
        columns: {
          dimension: categoryHeader,
        },
        chartHint: {
          preferredType: "donut",
          categoryField: categoryHeader,
          valueField: "value",
        },
        narrativeHint: {
          focus: "composition_ratio",
          dimension: categoryHeader,
        },
      }),
    );
  }

  if (scoreHeaders?.length && (categoryHeader || respondentHeader || questionHeader)) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: "survey_score_top_bottom",
        sectionType: "survey_score_top_bottom",
        recipeType: "top_bottom",
        title: `${primaryScoreHeader} 상위·하위 항목`,
        tableId,
        columns: {
          dimension: categoryHeader || respondentHeader || questionHeader,
          metric: primaryScoreHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: categoryHeader || respondentHeader || questionHeader,
          valueField: primaryScoreHeader,
        },
        narrativeHint: {
          focus: "top_bottom",
          metric: primaryScoreHeader,
        },
      }),
    );
  }

  return candidates.filter((candidate) => {
    if (!candidate.columns) return true;
    return Object.values(candidate.columns).every(Boolean);
  });
}

function buildSurveyScoreReportSections({
  normalizedQueryTables = [],
  table,
  templateCandidate = {},
  config = {},
}) {
  if (!table?.tableId) return [];

  const headers = findSurveyScoreHeaders(table, config);

  if (!headers.primaryScoreHeader) {
    const fallbackCandidates = Array.isArray(templateCandidate.candidates)
      ? templateCandidate.candidates
      : [];

    if (!fallbackCandidates.length) return [];

    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const customSections = [
    buildOverallScoreSection({ table, headers, config }),
    buildQuestionAverageSection({ table, headers, config }),
    buildScoreDistributionSection({ table, headers, config }),
    buildDimensionScoreSection({
      table,
      headers,
      dimensionHeader: headers.departmentHeader,
      sectionId: "survey_score_by_department",
      title: headers.departmentHeader
        ? `${headers.departmentHeader}별 평균 점수`
        : "부서별 평균 점수",
    }),
    buildDimensionScoreSection({
      table,
      headers,
      dimensionHeader: headers.categoryHeader,
      sectionId: "survey_score_by_category",
      title: headers.categoryHeader
        ? `${headers.categoryHeader}별 평균 점수`
        : "유형별 평균 점수",
    }),
    buildNpsSection({ table, headers, config }),
  ].filter(Boolean);

  const candidates = buildSurveyScoreCandidates({ table, headers, config });
  const recipeSections = executeTemplateSections({
    normalizedQueryTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });

  return [...customSections, ...recipeSections];
}

module.exports = {
  SURVEY_SCORE_REPORT_VERSION,
  findSurveyScoreHeaders,
  buildSurveyScoreCandidates,
  buildSurveyScoreReportSections,
};
