const { normalizeOutputTypes } = require("./businessTemplateContract");
const {
  IMPLEMENTATION_LEVELS,
  normalizeDomain,
  normalizeImplementationLevel,
  getTemplateDomainDef,
  getImplementationLevelDef,
} = require("./config/templateDomainConfig");
const {
  BUSINESS_TEMPLATE_DEFS,
} = require("./businessTemplates/templateDefinitions");

const GENERIC_TEMPLATE_HEADER_HINTS = new Set([
  "매출",
  "판매",
  "금액",
  "실적",
  "합계",
  "평균",
  "건수",
  "수량",
  "월",
  "연월",
  "연도",
  "기간",
  "상태",
  "구분",
  "유형",
  "비중",
  "구성비",
  "total",
  "sum",
  "amount",
  "count",
  "date",
  "month",
  "year",
  "status",
  "type",
]);

function addHeaderToken(headers, value) {
  if (value == null || value === "") return;
  const text = String(value).trim();
  if (text) headers.add(text);
}

function collectHeadersFromCandidates(analysisCandidates = [], context = {}) {
  const headers = new Set();

  [
    context.fileName,
    context.originalName,
    context.sourceFileName,
    context.sheetName,
    context.sourceSheetName,
  ].forEach((v) => addHeaderToken(headers, v));

  for (const c of analysisCandidates || []) {
    [
      c.groupHeader,
      c.metricHeader,
      c.dateHeader,
      c.statusHeader,
      c.dimension2Header,
      c.title,
      c.description,
      c.recommendationReason,
      c.sheetName,
      c.sourceSheetName,
      c.tableName,
      c.tableTitle,
      c.fileName,
      c.originalName,
      c.sourceFileName,
      c.metadata?.fileName,
      c.metadata?.originalName,
      c.source?.sheetName,
      c.source?.sourceSheetName,
      c.source?.tableName,
      c.source?.fileName,
      c.source?.originalName,
    ].forEach((v) => addHeaderToken(headers, v));

    if (c.columns && typeof c.columns === "object") {
      Object.values(c.columns).forEach((v) => addHeaderToken(headers, v));
    }

    (c.dimensions || []).forEach((v) => addHeaderToken(headers, v));
    (c.metrics || []).forEach((v) => addHeaderToken(headers, v));
    (c.dates || []).forEach((v) => addHeaderToken(headers, v));
    (c.statuses || []).forEach((v) => addHeaderToken(headers, v));
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

function isSpecificHeaderHint(hint = "") {
  const normalized = normalizeText(hint);
  if (!normalized || normalized.length < 2) return false;
  return !GENERIC_TEMPLATE_HEADER_HINTS.has(normalized);
}

function hasHeaderHint(headers = [], hint = "") {
  const h = normalizeText(hint);
  if (!h) return false;

  return headers.some((header) => {
    const target = normalizeText(header);
    if (!target) return false;
    if (target === h) return true;
    if (target.includes(h)) return true;
    // Reverse matching is useful for headers like "매출" vs hint "매출액",
    // but one-character headers such as "월" must not match unrelated hints like "월급".
    return target.length >= 2 && h.includes(target);
  });
}

function scoreHeaderHints(headers = [], hints = []) {
  return (hints || []).filter((hint) => hasHeaderHint(headers, hint)).length;
}

function unique(values = []) {
  const seen = new Set();
  const result = [];

  for (const value of values || []) {
    const text = String(value ?? "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    result.push(text);
  }

  return result;
}

function getTemplateMatchPolicy(templateMeta = {}) {
  const level = templateMeta.implementationLevel;

  if (level === IMPLEMENTATION_LEVELS.CUSTOM) {
    return {
      minimumFitScore: 0,
      recommendationFitScore: 55,
      minRecommendedMatchedCandidates: 1,
      minRecommendedMatchedHeaderHints: 2,
      minRecommendedSpecificRequiredAnyHints: 0,
      weakPriorityPenalty: 0,
    };
  }

  if (level === IMPLEMENTATION_LEVELS.DEFINITION_ONLY) {
    return {
      minimumFitScore: 42,
      recommendationFitScore: 63,
      minRecommendedMatchedCandidates: 1,
      minRecommendedMatchedHeaderHints: 2,
      minRecommendedSpecificRequiredAnyHints: 1,
      weakPriorityPenalty: 12,
    };
  }

  return {
    minimumFitScore: 50,
    recommendationFitScore: 72,
    minRecommendedMatchedCandidates: 2,
    minRecommendedMatchedHeaderHints: 2,
    minRecommendedSpecificRequiredAnyHints: 1,
    weakPriorityPenalty: 18,
  };
}

function collectMatchedHeaderHints(headers = [], hints = []) {
  return unique((hints || []).filter((hint) => hasHeaderHint(headers, hint)));
}

function collectMatchedHeaderValues(headers = [], hints = []) {
  return headers.filter((header) =>
    (hints || []).some((hint) => hasHeaderHint([header], hint)),
  );
}

function collectCandidateSourceTableIds(candidates = []) {
  return unique(
    (candidates || []).flatMap((candidate) => [
      candidate?.sourceTableId,
      candidate?.tableId,
      ...(Array.isArray(candidate?.sourceTableIds)
        ? candidate.sourceTableIds
        : []),
    ]),
  );
}

function collectCandidateSourceSheetNames(candidates = []) {
  return unique(
    (candidates || []).flatMap((candidate) => [
      candidate?.sourceSheetName,
      ...(Array.isArray(candidate?.sourceSheetNames)
        ? candidate.sourceSheetNames
        : []),
    ]),
  );
}

function getRecipeType(candidate = {}) {
  return candidate.recipeType || candidate.type || candidate.recipeId || "";
}

function templateDefinitionMeta(def = {}) {
  const domain = normalizeDomain(def.domain);
  const domainDef = getTemplateDomainDef(domain);
  const implementationLevel = normalizeImplementationLevel(
    def.implementationLevel,
    domain,
  );
  const implementationDef = getImplementationLevelDef(
    implementationLevel,
    domain,
  );

  return {
    domain,
    domainLabel: def.domainLabel || domainDef.label,
    domainGroup: def.domainGroup || domainDef.group,
    implementationLevel,
    implementationLevelLabel:
      def.implementationLevelLabel || implementationDef.label,
    preferredRecipeTypes: Array.isArray(def.preferredRecipeTypes)
      ? def.preferredRecipeTypes
      : [],
    templateTags: Array.isArray(def.templateTags) ? def.templateTags : [],
    templateDomainVersion: def.templateDomainVersion || null,
  };
}

function findCandidatesByTypes(analysisCandidates = [], types = []) {
  return types
    .map((type) => analysisCandidates.find((c) => getRecipeType(c) === type))
    .filter(Boolean);
}

function buildBusinessTemplateCandidate(
  def,
  analysisCandidates = [],
  context = {},
) {
  const headers = collectHeadersFromCandidates(analysisCandidates, context);
  const templateMeta = templateDefinitionMeta(def);
  const matchPolicy = getTemplateMatchPolicy(templateMeta);

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

  const allHeaderHints = [
    ...(def.requiredHeaderHints || []),
    ...(def.requiredAnyHeaderHints || []),
    ...(def.optionalHeaderHints || []),
  ];
  const matchedHeaderHints = collectMatchedHeaderHints(headers, allHeaderHints);
  const matchedRequiredAnyHeaderHints = collectMatchedHeaderHints(
    headers,
    def.requiredAnyHeaderHints || [],
  );
  const matchedSpecificRequiredAnyHeaderHints =
    matchedRequiredAnyHeaderHints.filter(isSpecificHeaderHint);
  const matchedHeaderValues = collectMatchedHeaderValues(
    headers,
    matchedHeaderHints,
  );
  const fitScore = Math.round(
    Math.max(0, Math.min(1, recipeScore * 0.45 + headerScore * 0.55)) * 100,
  );
  const hasEnoughSpecificRequiredAnyHints =
    !matchPolicy.minRecommendedSpecificRequiredAnyHints ||
    matchedSpecificRequiredAnyHeaderHints.length >=
      matchPolicy.minRecommendedSpecificRequiredAnyHints;
  const recommendedEligible =
    fitScore >= matchPolicy.recommendationFitScore &&
    matchedCandidates.length >= matchPolicy.minRecommendedMatchedCandidates &&
    matchedHeaderHints.length >= matchPolicy.minRecommendedMatchedHeaderHints &&
    hasEnoughSpecificRequiredAnyHints;
  const weakPriorityPenalty = recommendedEligible
    ? 0
    : matchPolicy.weakPriorityPenalty;
  const sourceTableIds = collectCandidateSourceTableIds(matchedCandidates);
  const sourceSheetNames = collectCandidateSourceSheetNames(matchedCandidates);

  if (fitScore < matchPolicy.minimumFitScore) {
    return null;
  }

  return {
    type: "businessTemplate",
    candidateType: "businessTemplate",
    templateId: def.templateId,
    title: def.title,
    description: def.description,
    domain: templateMeta.domain,
    domainLabel: templateMeta.domainLabel,
    domainGroup: templateMeta.domainGroup,
    implementationLevel: templateMeta.implementationLevel,
    implementationLevelLabel: templateMeta.implementationLevelLabel,
    preferredRecipeTypes: templateMeta.preferredRecipeTypes,
    templateTags: templateMeta.templateTags,
    templateDomainVersion: templateMeta.templateDomainVersion,
    outputTypes: normalizeOutputTypes(def.outputTypes),
    sourceTableIds,
    sourceSheetNames,
    sourceTableId: sourceTableIds[0] || "",
    sourceSheetName: sourceSheetNames[0] || "",
    sourceScope: sourceTableIds.length > 1 ? "multiTable" : "singleTable",
    priority: Math.max(0, Number(def.priority || 0) - weakPriorityPenalty),
    rawPriority: def.priority,
    confidence: Math.min(1, recipeScore * 0.55 + headerScore * 0.45),
    recommendedEligible,
    recommendationReason: recommendedEligible
      ? "업로드 데이터의 헤더와 분석 후보가 템플릿 조건에 충분히 매칭됩니다."
      : "템플릿 후보로는 표시하되, 추천 상위 노출에는 추가 데이터 적합도 확인이 필요합니다.",
    matchedHeaders: unique(matchedHeaderValues),
    matchedHeaderHints,
    matchedCount: matchedCandidates.length,
    templateMatch: {
      recipeScore: Math.round(recipeScore * 1000) / 1000,
      headerScore: Math.round(headerScore * 1000) / 1000,
      fitScore,
      matchedCandidateCount: matchedCandidates.length,
      matchedRequiredRecipeCount: matchedRequired.length,
      matchedAnyRecipeCount: matchedAnyRequired.length,
      matchedOptionalRecipeCount: matchedOptional.length,
      matchedHeaderHintCount: matchedHeaderHints.length,
      matchedRequiredAnyHeaderHints,
      matchedSpecificRequiredAnyHeaderHints,
      matchedSpecificRequiredAnyHeaderHintCount:
        matchedSpecificRequiredAnyHeaderHints.length,
      matchedHeaderValueCount: unique(matchedHeaderValues).length,
      matchedAnyHeaderCount,
      optionalHeaderScore,
      hasEnoughSpecificRequiredAnyHints,
      recommendationFitScore: matchPolicy.recommendationFitScore,
      minimumFitScore: matchPolicy.minimumFitScore,
    },
    candidates: matchedCandidates,
    primaryCandidate: matchedCandidates[0] || null,
  };
}

function buildBusinessTemplateCandidates(
  analysisCandidates = [],
  context = {},
) {
  if (!Array.isArray(analysisCandidates)) return [];

  return BUSINESS_TEMPLATE_DEFS.map((def) =>
    buildBusinessTemplateCandidate(def, analysisCandidates, context),
  )
    .filter(Boolean)
    .sort((a, b) => b.priority - a.priority || b.confidence - a.confidence);
}

module.exports = {
  BUSINESS_TEMPLATE_DEFS,
  templateDefinitionMeta,
  buildBusinessTemplateCandidates,
};
