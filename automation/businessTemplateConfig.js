const BUSINESS_TEMPLATE_DEFS = [
  {
    templateId: "hr_monthly_report",
    title: "월간 인사 보고서",
    description:
      "인원 현황, 부서별 집계, 추이, 상위/하위 항목을 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["category_count", "group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 100,
  },
  {
    templateId: "research_budget_report",
    title: "연구비 집행 현황",
    description:
      "예산·집행액·항목별 집계와 집행 추이를 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 90,
  },
  {
    templateId: "sales_report",
    title: "매출 분석 보고서",
    description:
      "매출 합계, 기간별 추이, 상위 항목을 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 80,
  },
  {
    templateId: "hr_monthly_report",
    title: "월간 인사 보고서",
    description:
      "인원 현황, 부서별 집계, 추이, 상위/하위 항목을 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["category_count"],
    optionalRecipeTypes: ["group_summary", "time_trend", "top_bottom"],
    requiredHeaderHints: ["부서"],
    optionalHeaderHints: ["직급", "입사일", "직원", "이름", "평가", "연봉"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 100,
  },
  {
    templateId: "research_budget_report",
    title: "연구비 집행 현황",
    description:
      "예산·집행액·항목별 집계와 집행 추이를 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    requiredHeaderHints: ["집행"],
    optionalHeaderHints: ["예산", "연구비", "비목", "항목", "잔액", "과제"],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
    priority: 90,
  },
  {
    templateId: "sales_report",
    title: "매출 분석 보고서",
    description:
      "매출 합계, 기간별 추이, 상위 항목을 보고서 형태로 생성합니다.",
    requiredRecipeTypes: ["group_summary"],
    optionalRecipeTypes: ["time_trend", "top_bottom"],
    requiredHeaderHints: ["매출"],
    optionalHeaderHints: [
      "고객",
      "제품",
      "상품",
      "거래처",
      "월",
      "일자",
      "수량",
    ],
    outputTypes: ["summarySheet", "ppt", "reportJson"],
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
    ].forEach((v) => {
      if (v) headers.add(String(v));
    });

    (c.dimensions || []).forEach((v) => headers.add(String(v)));
    (c.metrics || []).forEach((v) => headers.add(String(v)));
    (c.dates || []).forEach((v) => headers.add(String(v)));
    (c.statuses || []).forEach((v) => headers.add(String(v)));
  }

  return Array.from(headers);
}

function normalizeText(value = "") {
  return String(value).toLowerCase().replace(/\s+/g, "").trim();
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

function buildBusinessTemplateCandidate(def, analysisCandidates = []) {
  const matchedRequired = def.requiredRecipeTypes
    .map((type) => analysisCandidates.find((c) => getRecipeType(c) === type))
    .filter(Boolean);

  if (matchedRequired.length < def.requiredRecipeTypes.length) {
    return null;
  }

  const headers = collectHeadersFromCandidates(analysisCandidates);

  const requiredHeaderHints = def.requiredHeaderHints || [];
  const missingRequiredHeaders = requiredHeaderHints.filter(
    (hint) => !hasHeaderHint(headers, hint),
  );

  if (missingRequiredHeaders.length) {
    return null;
  }

  const matchedOptional = def.optionalRecipeTypes
    .map((type) => analysisCandidates.find((c) => getRecipeType(c) === type))
    .filter(Boolean);

  const matchedCandidates = [...matchedRequired, ...matchedOptional];

  const optionalHeaderScore = scoreHeaderHints(
    headers,
    def.optionalHeaderHints || [],
  );

  const recipeScore =
    (matchedRequired.length + matchedOptional.length * 0.5) /
    (def.requiredRecipeTypes.length + def.optionalRecipeTypes.length * 0.5);

  const headerScore =
    ((def.requiredHeaderHints || []).length + optionalHeaderScore * 0.5) /
    Math.max(
      1,
      (def.requiredHeaderHints || []).length +
        (def.optionalHeaderHints || []).length * 0.5,
    );

  return {
    templateId: def.templateId,
    title: def.title,
    description: def.description,
    outputTypes: def.outputTypes,
    priority: def.priority,
    confidence: Math.min(1, recipeScore * 0.6 + headerScore * 0.4),
    matchedHeaders: headers.filter((h) =>
      [
        ...(def.requiredHeaderHints || []),
        ...(def.optionalHeaderHints || []),
      ].some((hint) => hasHeaderHint([h], hint)),
    ),
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
