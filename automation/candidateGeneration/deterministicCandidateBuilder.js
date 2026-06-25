const {
  buildAnalysisRecipeCandidates,
} = require("../analysisRecipeCandidateBuilder");
const {
  buildBusinessTemplateCandidates,
} = require("../businessTemplateConfig");
const {
  AUTOMATION_CATEGORY_DEFS,
} = require("../config/candidateCategoryConfig");

function getRecipeType(candidate = {}) {
  return candidate.recipeType || candidate.type || candidate.recipeId || "";
}

function isRecipeType(candidate = {}, types = []) {
  return types.includes(getRecipeType(candidate));
}

function buildAutomationCategoryCandidates(analysisRecipeCandidates = []) {
  const list = Array.isArray(analysisRecipeCandidates)
    ? analysisRecipeCandidates
    : [];

  return AUTOMATION_CATEGORY_DEFS.map((def) => {
    const candidates = list.filter((candidate) =>
      isRecipeType(candidate, def.recipeTypes || []),
    );

    if (!candidates.length) return null;

    return {
      categoryId: def.categoryId,
      title: def.title,
      description: def.description,
      internalOnly: def.internalOnly,
      candidates,
    };
  }).filter(Boolean);
}

function buildDeterministicCandidateBundle({
  normalizedQueryTables = [],
  source = "deterministic",
} = {}) {
  const safeTables = Array.isArray(normalizedQueryTables)
    ? normalizedQueryTables
    : [];

  const analysisRecipeCandidates = buildAnalysisRecipeCandidates(safeTables);
  const categoryCandidates = buildAutomationCategoryCandidates(
    analysisRecipeCandidates,
  );
  const businessTemplateCandidates = buildBusinessTemplateCandidates(
    analysisRecipeCandidates,
  );

  return {
    analysisRecipeCandidates,
    categoryCandidates,
    businessTemplateCandidates,
    candidateGeneration: {
      version: "candidate_generation_v1",
      source,
      deterministic: {
        used: true,
        counts: {
          normalizedTables: safeTables.length,
          analysisRecipeCandidates: analysisRecipeCandidates.length,
          categoryCandidates: categoryCandidates.length,
          businessTemplateCandidates: businessTemplateCandidates.length,
        },
      },
      aiReranker: {
        enabled: false,
        used: false,
        skippedReason: "NOT_REQUESTED",
      },
      validation: {
        used: false,
      },
      generatedAt: new Date().toISOString(),
    },
  };
}

module.exports = {
  buildDeterministicCandidateBundle,
  buildAutomationCategoryCandidates,
};
