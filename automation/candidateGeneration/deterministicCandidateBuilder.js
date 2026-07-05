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
const {
  MULTI_SOURCE_CANDIDATE_VERSION,
  MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION,
  buildMultiSourceCandidates,
} = require("../multiSourceCandidateBuilder");
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

function buildMultiSourceCandidatesWithDiagnostics(args = {}) {
  const candidates = buildMultiSourceCandidates(args);
  return {
    candidates: Array.isArray(candidates) ? candidates : [],
    diagnostics: candidates?.diagnostics || null,
  };
}

function mergeUniqueByCandidateId(...groups) {
  const seen = new Set();
  const result = [];
  for (const group of groups) {
    for (const candidate of Array.isArray(group) ? group : []) {
      const id =
        candidate?.candidateId ||
        candidate?.id ||
        JSON.stringify(candidate).slice(0, 120);
      if (!id || seen.has(id)) continue;
      seen.add(id);
      result.push(candidate);
    }
  }
  return result;
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
  const initialMultiSource = buildMultiSourceCandidatesWithDiagnostics({
    normalizedQueryTables: safeTables,
    sourceTablePolicy,
    analysisRecipeCandidates,
  });

  const contractInputMultiSourceCandidates = initialMultiSource.candidates;

  const bundle = normalizeCandidateBundleContractV2(
    {
      analysisRecipeCandidates,
      categoryCandidates,
      businessTemplateCandidates,
      multiSourceCandidates: contractInputMultiSourceCandidates,
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
            multiSourceCandidates: contractInputMultiSourceCandidates.length,
          },
        },
        multiSourceCandidates: {
          version: MULTI_SOURCE_CANDIDATE_VERSION,
          payloadFallbackHotfixVersion:
            MULTI_SOURCE_PAYLOAD_FALLBACK_HOTFIX_VERSION || "",
          applied: true,
          count: contractInputMultiSourceCandidates.length,
          diagnostics: initialMultiSource.diagnostics,
          buildStage: "before_contract",
        },
        sourceTablePolicy: summarizeSourceTablePolicy(sourceTablePolicy),
        sourceTablePolicyRaw: sourceTablePolicy,
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

  let finalBundle = bundle;

  if (
    !Array.isArray(finalBundle.multiSourceCandidates) ||
    !finalBundle.multiSourceCandidates.length
  ) {
    const afterContractMultiSource = buildMultiSourceCandidatesWithDiagnostics({
      normalizedQueryTables: safeTables,
      sourceTablePolicy,
      analysisRecipeCandidates: [
        ...(finalBundle.analysisRecipeCandidates || []),
        ...(finalBundle.businessTemplateCandidates || []),
        ...(finalBundle.dashboardCandidates || []),
        ...(finalBundle.categoryCandidates || []),
      ],
    });

    finalBundle = {
      ...finalBundle,
      multiSourceCandidates: mergeUniqueByCandidateId(
        finalBundle.multiSourceCandidates || [],
        afterContractMultiSource.candidates,
      ),
      candidateGeneration: {
        ...(finalBundle.candidateGeneration || {}),
        multiSourceCandidates: {
          ...(finalBundle.candidateGeneration?.multiSourceCandidates || {}),
          afterContractFallbackVersion:
            "multi_source_after_contract_fallback_v1",
          afterContractFallbackApplied: true,
          afterContractFallbackCount:
            afterContractMultiSource.candidates.length,
          count: mergeUniqueByCandidateId(
            finalBundle.multiSourceCandidates || [],
            afterContractMultiSource.candidates,
          ).length,
          diagnostics: {
            beforeContract:
              finalBundle.candidateGeneration?.multiSourceCandidates
                ?.diagnostics || null,
            afterContract: afterContractMultiSource.diagnostics,
          },
        },
      },
    };
  }

  return scoreCandidateBundle(finalBundle, {
    sourceTablePolicy,
    normalizedQueryTables: safeTables,
  });
}

module.exports = {
  buildDeterministicCandidateBundle,
  buildAutomationCategoryCandidates,
};
