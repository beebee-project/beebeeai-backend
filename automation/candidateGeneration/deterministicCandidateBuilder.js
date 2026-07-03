const {
  buildAnalysisRecipeCandidates,
} = require("../analysisRecipeCandidateBuilder");
const {
  buildBusinessTemplateCandidates,
} = require("../businessTemplateConfig");
const {
  AUTOMATION_CATEGORY_DEFS,
} = require("../config/candidateCategoryConfig");
const {
  buildSourceTablePolicy,
  summarizeSourceTablePolicy,
  enrichCandidateListWithSourceTablePolicy,
} = require("../sourceTablePolicy");
const {
  normalizeCandidateBundleContractV2,
} = require("../candidateContractV2");
const { scoreCandidateBundle } = require("../candidateScorer");

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

  const sourceTablePolicy = buildSourceTablePolicy({
    tables: safeTables,
    normalizedQueryTables: safeTables,
  });
  const analysisRecipeCandidates = enrichCandidateListWithSourceTablePolicy(
    buildAnalysisRecipeCandidates(safeTables),
    sourceTablePolicy,
  );
  const categoryCandidates = buildAutomationCategoryCandidates(
    analysisRecipeCandidates,
  );
  const businessTemplateCandidates = enrichCandidateListWithSourceTablePolicy(
    buildBusinessTemplateCandidates(analysisRecipeCandidates),
    sourceTablePolicy,
  );

  const bundle = normalizeCandidateBundleContractV2(
    {
      analysisRecipeCandidates,
      categoryCandidates,
      businessTemplateCandidates,
      candidateGeneration: {
        version: "candidate_generation_v2",
        source,
        deterministic: {
          used: true,
          counts: {
            normalizedTables: safeTables.length,
            sourceTables: sourceTablePolicy.counts?.sourceTableCount || 0,
            virtualTables: sourceTablePolicy.counts?.virtualTableCount || 0,
            analysisRecipeCandidates: analysisRecipeCandidates.length,
            categoryCandidates: categoryCandidates.length,
            businessTemplateCandidates: businessTemplateCandidates.length,
          },
        },
        sourceTablePolicy: summarizeSourceTablePolicy(sourceTablePolicy),
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
    },
    { source },
  );

  return scoreCandidateBundle(bundle, { sourceTablePolicy });
}

module.exports = {
  buildDeterministicCandidateBundle,
  buildAutomationCategoryCandidates,
};
