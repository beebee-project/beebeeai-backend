const { BUSINESS_TEMPLATE_DEFS } = require("../businessTemplateConfig");

const ALLOWED_OUTPUT_TYPES = ["summarySheet", "analysisReport", "ppt"];

function stableCandidateId(candidate = {}) {
  return String(
    candidate.candidateId ||
      candidate.id ||
      candidate.recipeId ||
      [candidate.recipeType, candidate.tableId, candidate.title]
        .filter(Boolean)
        .join(":"),
  );
}

function normalizeOutputTypes(outputTypes = []) {
  const mapped = (Array.isArray(outputTypes) ? outputTypes : [])
    .map((type) => (type === "reportJson" ? "analysisReport" : type))
    .filter((type) => ALLOWED_OUTPUT_TYPES.includes(type));

  const deduped = [...new Set(mapped)];
  return deduped.length ? deduped : [...ALLOWED_OUTPUT_TYPES];
}

function clampNumber(value, fallback = 0, min = 0, max = 1) {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

function cleanText(value = "", max = 300) {
  const s = String(value || "").replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

function sanitizeAnalysisCandidate(candidate = {}, tableIds = new Set()) {
  const recipeType = String(
    candidate.recipeType || candidate.type || candidate.recipeId || "",
  ).trim();
  const tableId = String(candidate.tableId || "").trim();

  if (!recipeType || !tableId) return null;
  if (tableIds.size && !tableIds.has(tableId)) return null;

  return {
    ...candidate,
    id: stableCandidateId(candidate),
    recipeType,
    tableId,
    title: cleanText(candidate.title || `${recipeType} 후보`, 120),
    description: cleanText(candidate.description || "", 500),
    recommendationReason: cleanText(candidate.recommendationReason || "", 500),
    confidence: clampNumber(candidate.confidence, 0.5, 0, 1),
    aiAssisted: Boolean(candidate.aiAssisted),
  };
}

function sanitizeBusinessTemplateCandidate(
  candidate = {},
  analysisIdSet = new Set(),
) {
  const def = BUSINESS_TEMPLATE_DEFS.find(
    (x) => x.templateId === candidate.templateId,
  );
  if (!def) return null;

  const nested = Array.isArray(candidate.candidates) ? candidate.candidates : [];
  const safeNested = nested.filter((c) => analysisIdSet.has(stableCandidateId(c)));

  return {
    type: "businessTemplate",
    templateId: def.templateId,
    title: cleanText(candidate.title || def.title, 120),
    description: cleanText(candidate.description || def.description, 500),
    outputTypes: normalizeOutputTypes(candidate.outputTypes || def.outputTypes),
    priority: Number.isFinite(Number(candidate.priority))
      ? Number(candidate.priority)
      : Number(def.priority || 0),
    confidence: clampNumber(candidate.confidence, 0.5, 0, 1),
    matchedHeaders: Array.isArray(candidate.matchedHeaders)
      ? candidate.matchedHeaders.map((x) => cleanText(x, 80)).filter(Boolean)
      : [],
    matchedCount: Number(candidate.matchedCount || safeNested.length || 0),
    candidates: safeNested,
    primaryCandidate: safeNested[0] || null,
    recommendationReason: cleanText(candidate.recommendationReason || "", 500),
    aiAssisted: Boolean(candidate.aiAssisted),
  };
}

function validateCandidateBundle(bundle = {}, normalizedQueryTables = []) {
  const tableIds = new Set(
    (Array.isArray(normalizedQueryTables) ? normalizedQueryTables : [])
      .map((table) => String(table?.tableId || ""))
      .filter(Boolean),
  );

  const originalAnalysis = Array.isArray(bundle.analysisRecipeCandidates)
    ? bundle.analysisRecipeCandidates
    : [];

  const analysisRecipeCandidates = originalAnalysis
    .map((candidate) => sanitizeAnalysisCandidate(candidate, tableIds))
    .filter(Boolean);

  const analysisIdSet = new Set(analysisRecipeCandidates.map(stableCandidateId));

  const originalBusiness = Array.isArray(bundle.businessTemplateCandidates)
    ? bundle.businessTemplateCandidates
    : [];

  const businessTemplateCandidates = originalBusiness
    .map((candidate) =>
      sanitizeBusinessTemplateCandidate(candidate, analysisIdSet),
    )
    .filter(Boolean)
    .sort((a, b) => b.priority - a.priority || b.confidence - a.confidence);

  const categoryCandidates = (Array.isArray(bundle.categoryCandidates)
    ? bundle.categoryCandidates
    : []
  )
    .map((category) => ({
      ...category,
      internalOnly: category.internalOnly !== false,
      candidates: (Array.isArray(category.candidates)
        ? category.candidates
        : []
      ).filter((candidate) => analysisIdSet.has(stableCandidateId(candidate))),
    }))
    .filter((category) => category.candidates.length);

  const validation = {
    used: true,
    removed: {
      analysisRecipeCandidates:
        originalAnalysis.length - analysisRecipeCandidates.length,
      businessTemplateCandidates:
        originalBusiness.length - businessTemplateCandidates.length,
    },
    allowedOutputTypes: ALLOWED_OUTPUT_TYPES,
  };

  return {
    ...bundle,
    analysisRecipeCandidates,
    categoryCandidates,
    businessTemplateCandidates,
    candidateGeneration: {
      ...(bundle.candidateGeneration || {}),
      validation,
    },
  };
}

module.exports = {
  ALLOWED_OUTPUT_TYPES,
  stableCandidateId,
  normalizeOutputTypes,
  validateCandidateBundle,
};
