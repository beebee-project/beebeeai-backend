const {
  ALLOWED_OUTPUT_TYPES,
  normalizeOutputTypes,
} = require("./businessTemplateContract");

const CANDIDATE_CONTRACT_VERSION = "candidate_contract_v2";
const CANDIDATE_TYPE = Object.freeze({
  ANALYSIS_RECIPE: "analysisRecipe",
  DASHBOARD: "dashboard",
  BUSINESS_TEMPLATE: "businessTemplate",
  AUTOMATION_CATEGORY: "automationCategory",
});

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const result = [];

  for (const value of values) {
    const text = String(value ?? "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    result.push(text);
  }

  return result;
}

function cleanText(value = "", max = 300) {
  const text = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "";
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function clampNumber(value, fallback = 0.5, min = 0, max = 1) {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

function stableCandidateId(candidate = {}, fallbackPrefix = "candidate") {
  return (
    cleanText(
      candidate.candidateId ||
        candidate.id ||
        candidate.dashboardId ||
        candidate.categoryId ||
        candidate.templateId ||
        candidate.recipeId ||
        [
          candidate.candidateType,
          candidate.type,
          candidate.recipeType,
          candidate.tableId,
          candidate.sourceTableId,
          candidate.title,
        ]
          .filter(Boolean)
          .join(":"),
      180,
    ) || `${fallbackPrefix}_${Math.random().toString(36).slice(2, 8)}`
  );
}

function inferCandidateType(candidate = {}, fallback = "") {
  if (candidate.candidateType) return String(candidate.candidateType);
  if (candidate.templateId || candidate.type === "businessTemplate") {
    return CANDIDATE_TYPE.BUSINESS_TEMPLATE;
  }
  if (candidate.dashboardId || candidate.type === "dashboard") {
    return CANDIDATE_TYPE.DASHBOARD;
  }
  if (candidate.categoryId && Array.isArray(candidate.candidates)) {
    return CANDIDATE_TYPE.AUTOMATION_CATEGORY;
  }
  if (candidate.recipeType || candidate.tableId) {
    return CANDIDATE_TYPE.ANALYSIS_RECIPE;
  }
  return fallback || "candidate";
}

function collectNestedCandidates(candidate = {}) {
  return [
    ...asArray(candidate.candidates),
    ...asArray(candidate.recipes),
    ...asArray(candidate.sections)
      .map((section) => section?.candidate)
      .filter(Boolean),
  ];
}

function collectRecipeIds(candidate = {}) {
  const nested = collectNestedCandidates(candidate);
  return unique([
    ...asArray(candidate.recipeIds),
    candidate.recipeId,
    candidate.recipeType,
    candidate.type && candidate.type !== "businessTemplate"
      ? candidate.type
      : "",
    ...nested.flatMap((item) => [
      item?.candidateId,
      item?.id,
      item?.recipeId,
      item?.recipeType,
    ]),
  ]);
}

function collectSourceTableIds(candidate = {}) {
  const nested = collectNestedCandidates(candidate);
  return unique([
    ...asArray(candidate.sourceTableIds),
    candidate.sourceTableId,
    candidate.tableId,
    ...nested.flatMap((item) => [item?.sourceTableId, item?.tableId]),
  ]);
}

function collectSourceSheetNames(candidate = {}) {
  const nested = collectNestedCandidates(candidate);
  return unique([
    ...asArray(candidate.sourceSheetNames),
    candidate.sourceSheetName,
    ...nested.flatMap((item) => [item?.sourceSheetName]),
  ]);
}

function inferSourceScope(candidate = {}, sourceTableIds = []) {
  const explicit = String(candidate.sourceScope || "").trim();
  if (explicit) return explicit;
  if (sourceTableIds.length > 1) return "multiTable";
  return "singleTable";
}

function normalizeReasonCodes(candidate = {}) {
  return unique([
    ...asArray(candidate.reasonCodes),
    ...asArray(candidate.reasons),
    candidate.recommendationReason ? "HAS_RECOMMENDATION_REASON" : "",
    candidate.aiAssisted ? "AI_ASSISTED" : "",
  ]).slice(0, 20);
}

function normalizeCandidateContractV2(candidate = {}, options = {}) {
  if (!candidate || typeof candidate !== "object") return candidate;

  const candidateType = inferCandidateType(candidate, options.candidateType);
  const sourceTableIds = collectSourceTableIds(candidate);
  const sourceSheetNames = collectSourceSheetNames(candidate);
  const recipeIds = collectRecipeIds(candidate);
  const outputTypes = normalizeOutputTypes(
    candidate.outputTypes ||
      options.outputTypes ||
      (candidateType === CANDIDATE_TYPE.ANALYSIS_RECIPE
        ? ["summarySheet"]
        : ALLOWED_OUTPUT_TYPES),
  );
  const candidateId = stableCandidateId(candidate, candidateType);
  const priority = Number.isFinite(Number(candidate.priority))
    ? Number(candidate.priority)
    : Number.isFinite(Number(options.priority))
      ? Number(options.priority)
      : 0;

  return {
    ...candidate,
    candidateContractVersion:
      candidate.candidateContractVersion || CANDIDATE_CONTRACT_VERSION,
    candidateId,
    id: candidate.id || candidateId,
    candidateType,
    title: cleanText(candidate.title || candidate.name || candidateId, 160),
    description: cleanText(
      candidate.description || candidate.summary || "",
      600,
    ),
    sourceScope: inferSourceScope(candidate, sourceTableIds),
    sourceTableIds,
    sourceSheetNames,
    sourceTableId: candidate.sourceTableId || sourceTableIds[0] || "",
    sourceSheetName: candidate.sourceSheetName || sourceSheetNames[0] || "",
    recipeIds,
    outputTypes,
    confidence: clampNumber(
      candidate.confidence,
      options.confidence ?? 0.5,
      0,
      1,
    ),
    priority,
    reasonCodes: normalizeReasonCodes(candidate),
  };
}

function normalizeAnalysisRecipeCandidate(candidate = {}, index = 0) {
  return normalizeCandidateContractV2(candidate, {
    candidateType: CANDIDATE_TYPE.ANALYSIS_RECIPE,
    outputTypes: ["summarySheet"],
    priority: Number.isFinite(Number(candidate.priority))
      ? Number(candidate.priority)
      : 100 - index,
  });
}

function normalizeBusinessTemplateCandidate(candidate = {}, index = 0) {
  const nested = asArray(candidate.candidates).map((item, itemIndex) =>
    normalizeAnalysisRecipeCandidate(item, itemIndex),
  );

  return normalizeCandidateContractV2(
    {
      ...candidate,
      candidates: nested,
      primaryCandidate: nested[0] || candidate.primaryCandidate || null,
    },
    {
      candidateType: CANDIDATE_TYPE.BUSINESS_TEMPLATE,
      outputTypes: candidate.outputTypes || ALLOWED_OUTPUT_TYPES,
      priority: Number.isFinite(Number(candidate.priority))
        ? Number(candidate.priority)
        : 1000 - index,
    },
  );
}

function buildDashboardCandidateFromCategory(category = {}, index = 0) {
  const nested = asArray(category.candidates).map((item, itemIndex) =>
    normalizeAnalysisRecipeCandidate(item, itemIndex),
  );
  if (!nested.length) return null;

  const dashboardId =
    category.dashboardId || category.categoryId || `dashboard_${index + 1}`;

  return normalizeCandidateContractV2(
    {
      ...category,
      dashboardId,
      candidateId: category.candidateId || dashboardId,
      type: "dashboard",
      candidateType: CANDIDATE_TYPE.DASHBOARD,
      title: category.title || `자동화 대시보드 ${index + 1}`,
      description:
        category.description ||
        "관련 분석 후보를 묶어 자동화 시트로 구성합니다.",
      candidates: nested,
      recipeIds: nested.flatMap(
        (item) => item.recipeIds || item.recipeType || item.id,
      ),
      sourceTableIds: nested.flatMap(
        (item) => item.sourceTableIds || item.sourceTableId,
      ),
      sourceSheetNames: nested.flatMap(
        (item) => item.sourceSheetNames || item.sourceSheetName,
      ),
      reasonCodes: ["CATEGORY_DASHBOARD"],
      outputTypes: ["summarySheet", "analysisReport", "ppt"],
    },
    {
      candidateType: CANDIDATE_TYPE.DASHBOARD,
      priority: Number(category.priority || 500 - index),
      outputTypes: ["summarySheet", "analysisReport", "ppt"],
    },
  );
}

function normalizeCategoryCandidate(category = {}, index = 0) {
  const nested = asArray(category.candidates).map((item, itemIndex) =>
    normalizeAnalysisRecipeCandidate(item, itemIndex),
  );

  return normalizeCandidateContractV2(
    {
      ...category,
      candidates: nested,
      recipeIds: nested.flatMap(
        (item) => item.recipeIds || item.recipeType || item.id,
      ),
      sourceTableIds: nested.flatMap(
        (item) => item.sourceTableIds || item.sourceTableId,
      ),
      sourceSheetNames: nested.flatMap(
        (item) => item.sourceSheetNames || item.sourceSheetName,
      ),
      reasonCodes: ["AUTOMATION_CATEGORY"],
      outputTypes: ["summarySheet"],
    },
    {
      candidateType: CANDIDATE_TYPE.AUTOMATION_CATEGORY,
      priority: Number(category.priority || 400 - index),
      outputTypes: ["summarySheet"],
    },
  );
}

function normalizeCandidateBundleContractV2(bundle = {}, options = {}) {
  const analysisRecipeCandidates = asArray(bundle.analysisRecipeCandidates).map(
    normalizeAnalysisRecipeCandidate,
  );
  const categoryCandidates = asArray(bundle.categoryCandidates).map(
    normalizeCategoryCandidate,
  );
  const dashboardCandidates = (
    asArray(bundle.dashboardCandidates).length
      ? asArray(bundle.dashboardCandidates).map((candidate, index) =>
          normalizeCandidateContractV2(candidate, {
            candidateType: CANDIDATE_TYPE.DASHBOARD,
            priority: Number(candidate.priority || 600 - index),
            outputTypes: ["summarySheet", "analysisReport", "ppt"],
          }),
        )
      : categoryCandidates.map(buildDashboardCandidateFromCategory)
  ).filter(Boolean);
  const businessTemplateCandidates = asArray(
    bundle.businessTemplateCandidates,
  ).map(normalizeBusinessTemplateCandidate);

  return {
    ...bundle,
    analysisRecipeCandidates,
    categoryCandidates,
    dashboardCandidates,
    businessTemplateCandidates,
    candidateContract: {
      version: CANDIDATE_CONTRACT_VERSION,
      generatedAt: new Date().toISOString(),
      counts: {
        analysisRecipeCandidates: analysisRecipeCandidates.length,
        categoryCandidates: categoryCandidates.length,
        dashboardCandidates: dashboardCandidates.length,
        businessTemplateCandidates: businessTemplateCandidates.length,
      },
      fields: [
        "candidateId",
        "candidateType",
        "title",
        "description",
        "sourceScope",
        "sourceTableIds",
        "sourceSheetNames",
        "recipeIds",
        "confidence",
        "priority",
        "reasonCodes",
        "outputTypes",
      ],
      source:
        options.source ||
        bundle.candidateGeneration?.source ||
        "candidate-bundle",
      fileName: options.fileName || "",
    },
    candidateGeneration: {
      ...(bundle.candidateGeneration || {}),
      candidateContract: {
        version: CANDIDATE_CONTRACT_VERSION,
        applied: true,
      },
    },
  };
}

module.exports = {
  CANDIDATE_CONTRACT_VERSION,
  CANDIDATE_TYPE,
  stableCandidateId,
  normalizeCandidateContractV2,
  normalizeAnalysisRecipeCandidate,
  normalizeBusinessTemplateCandidate,
  normalizeCategoryCandidate,
  buildDashboardCandidateFromCategory,
  normalizeCandidateBundleContractV2,
};
