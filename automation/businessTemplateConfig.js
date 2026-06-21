const { normalizeOutputTypes } = require("./businessTemplateContract");

const BUSINESS_TEMPLATE_DEFS = [
  {
    templateId: "sales_report",
    title: "매출 분석 보고서",
    description:
      "기간별 매출, 수량, 상위 항목, 평균 판매금액을 보고서 형태로 생성합니다.",
    requiredAnyRecipeTypes: ["time_trend", "group_summary", "top_bottom"],
    requiredAnyHeaderHints: [
      "매출",
      "순매출액",
      "판매",
      "수량",
      "카드매출",
      "revenue",
      "sales",
    ],
    optionalHeaderHints: [
      "연도",
      "월",
      "연월",
      "제품",
      "상품",
      "지역",
      "업종",
      "거래처",
    ],
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    priority: 100,
  },
  {
    templateId: "research_budget_report",
    title: "연구비 집행 현황",
    description:
      "연구비, 집행액, 항목별·기관별·연도별 현황을 보고서 형태로 생성합니다.",
    requiredAnyRecipeTypes: ["group_summary", "time_trend", "top_bottom"],
    requiredAnyHeaderHints: [
      "연구비",
      "집행",
      "정부출연금",
      "과제",
      "항목명",
      "기관분류",
      "전문기관",
    ],
    optionalHeaderHints: [
      "예산",
      "현금",
      "현물",
      "사업명",
      "연구기관",
      "연구책임자",
      "진행년도",
      "예산년도",
    ],
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    priority: 90,
  },
  {
    templateId: "hr_monthly_report",
    title: "월간 인사 보고서",
    description:
      "명단, 상태, 부서별·직급별 현황, 입사 추이, 연봉 요약을 보고서 형태로 생성합니다.",
    requiredAnyRecipeTypes: ["category_count", "group_summary", "top_bottom"],
    requiredAnyHeaderHints: [
      "부서",
      "소속",
      "직급",
      "직위",
      "재직",
      "상태",
      "성명",
      "이름",
      "입사",
      "연봉",
      "급여",
    ],
    optionalHeaderHints: [
      "직원",
      "사원",
      "인사",
      "평가",
      "조직",
      "팀",
      "근무",
      "퇴사",
    ],
    outputTypes: ["summarySheet", "analysisReport", "ppt"],
    priority: 80,
  },
];

function collectHeadersFromCandidates(analysisCandidates = []) {
  const headers = new Set();

  for (const c of analysisCandidates || []) {
    [
      c.groupHeader,
      c.metricHeader,
      c.dateHeader,
      c.statusHeader,
      c.title,
      c.description,
    ].forEach((v) => {
      if (v) headers.add(String(v));
    });

    if (c.columns && typeof c.columns === "object") {
      Object.values(c.columns).forEach((v) => {
        if (v) headers.add(String(v));
      });
    }

    (c.dimensions || []).forEach((v) => headers.add(String(v)));
    (c.metrics || []).forEach((v) => headers.add(String(v)));
    (c.dates || []).forEach((v) => headers.add(String(v)));
    (c.statuses || []).forEach((v) => headers.add(String(v)));
  }

  return Array.from(headers);
}

function normalizeText(value = "") {
  return String(value)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function hasHeaderHint(headers = [], hint = "") {
  const h = normalizeText(hint);
  if (!h) return false;

  return headers.some((header) => {
    const target = normalizeText(header);
    return target.includes(h) || h.includes(target);
  });
}

function scoreHeaderHints(headers = [], hints = []) {
  return (hints || []).filter((hint) => hasHeaderHint(headers, hint)).length;
}

function getRecipeType(candidate = {}) {
  return candidate.recipeType || candidate.type || candidate.recipeId || "";
}

function findCandidatesByTypes(analysisCandidates = [], types = []) {
  return types
    .map((type) => analysisCandidates.find((c) => getRecipeType(c) === type))
    .filter(Boolean);
}

function buildBusinessTemplateCandidate(def, analysisCandidates = []) {
  const headers = collectHeadersFromCandidates(analysisCandidates);

  const requiredRecipeTypes = def.requiredRecipeTypes || [];
  const requiredAnyRecipeTypes = def.requiredAnyRecipeTypes || [];

  const matchedRequired = findCandidatesByTypes(
    analysisCandidates,
    requiredRecipeTypes,
  );

  if (matchedRequired.length < requiredRecipeTypes.length) {
    return null;
  }

  const matchedAnyRequired = findCandidatesByTypes(
    analysisCandidates,
    requiredAnyRecipeTypes,
  );

  if (requiredAnyRecipeTypes.length && !matchedAnyRequired.length) {
    return null;
  }

  const requiredHeaderHints = def.requiredHeaderHints || [];
  const missingRequiredHeaders = requiredHeaderHints.filter(
    (hint) => !hasHeaderHint(headers, hint),
  );

  if (missingRequiredHeaders.length) {
    return null;
  }

  const requiredAnyHeaderHints = def.requiredAnyHeaderHints || [];
  const matchedAnyHeaderCount = scoreHeaderHints(
    headers,
    requiredAnyHeaderHints,
  );

  if (requiredAnyHeaderHints.length && !matchedAnyHeaderCount) {
    return null;
  }

  const optionalRecipeTypes = def.optionalRecipeTypes || [];
  const matchedOptional = findCandidatesByTypes(
    analysisCandidates,
    optionalRecipeTypes,
  );

  const matchedCandidates = [
    ...matchedRequired,
    ...matchedAnyRequired,
    ...matchedOptional,
  ].filter((candidate, index, arr) => arr.indexOf(candidate) === index);

  const optionalHeaderScore = scoreHeaderHints(
    headers,
    def.optionalHeaderHints || [],
  );

  const recipeDenominator =
    requiredRecipeTypes.length +
    Math.min(1, requiredAnyRecipeTypes.length) +
    optionalRecipeTypes.length * 0.5;

  const recipeNumerator =
    matchedRequired.length +
    Math.min(1, matchedAnyRequired.length) +
    matchedOptional.length * 0.5;

  const recipeScore = recipeDenominator
    ? recipeNumerator / recipeDenominator
    : 0.5;

  const headerDenominator =
    requiredHeaderHints.length +
    Math.min(1, requiredAnyHeaderHints.length) +
    (def.optionalHeaderHints || []).length * 0.5;

  const headerNumerator =
    requiredHeaderHints.length +
    Math.min(1, matchedAnyHeaderCount) +
    optionalHeaderScore * 0.5;

  const headerScore = headerDenominator
    ? headerNumerator / headerDenominator
    : 0.5;

  const matchedHeaderHints = [
    ...(def.requiredHeaderHints || []),
    ...(def.requiredAnyHeaderHints || []),
    ...(def.optionalHeaderHints || []),
  ];

  return {
    type: "businessTemplate",
    templateId: def.templateId,
    title: def.title,
    description: def.description,
    outputTypes: normalizeOutputTypes(def.outputTypes),
    priority: def.priority,
    confidence: Math.min(1, recipeScore * 0.55 + headerScore * 0.45),
    matchedHeaders: headers.filter((h) =>
      matchedHeaderHints.some((hint) => hasHeaderHint([h], hint)),
    ),
    matchedCount: matchedCandidates.length,
    candidates: matchedCandidates,
    primaryCandidate: matchedCandidates[0] || null,
  };
}

function buildBusinessTemplateCandidates(analysisCandidates = []) {
  if (!Array.isArray(analysisCandidates)) return [];

  return BUSINESS_TEMPLATE_DEFS.map((def) =>
    buildBusinessTemplateCandidate(def, analysisCandidates),
  )
    .filter(Boolean)
    .sort((a, b) => b.priority - a.priority || b.confidence - a.confidence);
}

module.exports = {
  BUSINESS_TEMPLATE_DEFS,
  buildBusinessTemplateCandidates,
};
