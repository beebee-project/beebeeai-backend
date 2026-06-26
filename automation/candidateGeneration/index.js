const { stableCandidateId } = require("./candidateValidator");

function isEnabled() {
  return (
    String(process.env.USE_AI_CANDIDATE_RERANKER || "") === "true" ||
    process.env.USE_AI_CANDIDATE_RERANKER === "1"
  );
}

function getModelName() {
  return process.env.CANDIDATE_RERANKER_MODEL || "gpt-4o-mini";
}

function getOpenAIClient() {
  if (!process.env.OPENAI_API_KEY) return null;
  try {
    const OpenAI = require("openai");
    return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  } catch (error) {
    console.warn("[aiCandidateReranker] OpenAI init skipped:", error.message);
    return null;
  }
}

function safeHeaders(normalizedQueryTables = []) {
  return (Array.isArray(normalizedQueryTables) ? normalizedQueryTables : [])
    .slice(0, 5)
    .map((table) => ({
      tableId: table.tableId,
      sheetName: table.sheetName,
      rowCount: table.rowCount,
      columns: (table.columns || []).slice(0, 30).map((column) => ({
        header: column.header,
        type: column.type,
        role: column.role,
        uniqueCount: column.uniqueCount,
      })),
    }));
}

function compactAnalysis(candidate = {}) {
  return {
    id: stableCandidateId(candidate),
    recipeType: candidate.recipeType,
    tableId: candidate.tableId,
    title: candidate.title,
    description: candidate.description,
    columns: candidate.columns || {},
  };
}

function compactBusiness(candidate = {}) {
  return {
    templateId: candidate.templateId,
    title: candidate.title,
    description: candidate.description,
    confidence: candidate.confidence,
    outputTypes: candidate.outputTypes,
    matchedHeaders: candidate.matchedHeaders || [],
    candidateIds: (candidate.candidates || []).map(stableCandidateId),
  };
}

function parseJson(text = "") {
  const raw = String(text || "").trim();
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (_) {
    const matched = raw.match(/\{[\s\S]*\}/);
    if (!matched) return null;
    try {
      return JSON.parse(matched[0]);
    } catch (_error) {
      return null;
    }
  }
}

function clampBoost(value) {
  const n = Number(value || 0);
  if (!Number.isFinite(n)) return 0;
  return Math.max(-5, Math.min(5, n));
}

function applyAiSuggestions(bundle = {}, ai = {}) {
  const analysisSuggestions = new Map(
    (Array.isArray(ai.analysisCandidates) ? ai.analysisCandidates : [])
      .filter((item) => item?.id)
      .map((item) => [String(item.id), item]),
  );

  const businessSuggestions = new Map(
    (Array.isArray(ai.businessTemplates) ? ai.businessTemplates : [])
      .filter((item) => item?.templateId)
      .map((item) => [String(item.templateId), item]),
  );

  const analysisRecipeCandidates = (bundle.analysisRecipeCandidates || [])
    .map((candidate) => {
      const id = stableCandidateId(candidate);
      const suggestion = analysisSuggestions.get(id);
      if (!suggestion) return candidate;

      return {
        ...candidate,
        description: suggestion.description || candidate.description,
        recommendationReason:
          suggestion.recommendationReason || candidate.recommendationReason,
        _aiPriorityBoost: clampBoost(suggestion.priorityBoost),
        aiAssisted: true,
      };
    })
    .sort(
      (a, b) =>
        (b._aiPriorityBoost || 0) - (a._aiPriorityBoost || 0) ||
        (b.confidence || 0) - (a.confidence || 0),
    );

  const businessTemplateCandidates = (bundle.businessTemplateCandidates || [])
    .map((candidate) => {
      const suggestion = businessSuggestions.get(candidate.templateId);
      if (!suggestion) return candidate;

      return {
        ...candidate,
        description: suggestion.description || candidate.description,
        recommendationReason:
          suggestion.recommendationReason || candidate.recommendationReason,
        _aiPriorityBoost: clampBoost(suggestion.priorityBoost),
        aiAssisted: true,
      };
    })
    .sort(
      (a, b) =>
        b.priority +
          (b._aiPriorityBoost || 0) -
          (a.priority + (a._aiPriorityBoost || 0)) ||
        (b.confidence || 0) - (a.confidence || 0),
    );

  return {
    ...bundle,
    analysisRecipeCandidates,
    businessTemplateCandidates,
  };
}

async function rerankCandidateBundle({
  normalizedQueryTables = [],
  bundle = {},
  fileName = "",
} = {}) {
  const enabled = isEnabled();

  if (!enabled) {
    return {
      bundle,
      meta: {
        enabled: false,
        used: false,
        skippedReason: "USE_AI_CANDIDATE_RERANKER_DISABLED",
      },
    };
  }

  const client = getOpenAIClient();
  if (!client) {
    return {
      bundle,
      meta: {
        enabled: true,
        used: false,
        skippedReason: "OPENAI_CLIENT_UNAVAILABLE",
      },
    };
  }

  const payload = {
    fileName,
    tables: safeHeaders(normalizedQueryTables),
    analysisCandidates: (bundle.analysisRecipeCandidates || [])
      .slice(0, 20)
      .map(compactAnalysis),
    businessTemplates: (bundle.businessTemplateCandidates || [])
      .slice(0, 10)
      .map(compactBusiness),
  };

  try {
    const model = getModelName();
    const completion = await client.chat.completions.create({
      model,
      temperature: 0.2,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content:
            "You are a candidate reranker for spreadsheet business templates. You may only improve descriptions, recommendation reasons, and priorityBoost for existing candidates. Do not add candidates, remove candidates, calculate data, create sections, create formulas, or generate code. Return JSON only.",
        },
        {
          role: "user",
          content: JSON.stringify({
            task: "Rerank and explain existing candidates. Keep ids/templateIds unchanged. priorityBoost must be an integer from -5 to 5.",
            schema: {
              businessTemplates: [
                {
                  templateId: "existing templateId only",
                  description: "short Korean description",
                  recommendationReason: "short Korean reason",
                  priorityBoost: 0,
                },
              ],
              analysisCandidates: [
                {
                  id: "existing id only",
                  description: "short Korean description",
                  recommendationReason: "short Korean reason",
                  priorityBoost: 0,
                },
              ],
            },
            payload,
          }),
        },
      ],
    });

    const content = completion.choices?.[0]?.message?.content || "";
    const parsed = parseJson(content);
    if (!parsed) {
      return {
        bundle,
        meta: {
          enabled: true,
          used: false,
          model,
          skippedReason: "AI_RESPONSE_PARSE_FAILED",
        },
      };
    }

    return {
      bundle: applyAiSuggestions(bundle, parsed),
      meta: {
        enabled: true,
        used: true,
        model,
        applied: {
          businessTemplates: Array.isArray(parsed.businessTemplates)
            ? parsed.businessTemplates.length
            : 0,
          analysisCandidates: Array.isArray(parsed.analysisCandidates)
            ? parsed.analysisCandidates.length
            : 0,
        },
      },
    };
  } catch (error) {
    console.warn(
      "[aiCandidateReranker] fallback deterministic:",
      error.message,
    );
    return {
      bundle,
      meta: {
        enabled: true,
        used: false,
        skippedReason: "AI_RERANKER_FAILED",
        error: error.code || error.name || "ERROR",
      },
    };
  }
}

function resolveExportedFunction(mod, names = []) {
  if (typeof mod === "function") return mod;

  for (const name of names) {
    if (typeof mod?.[name] === "function") return mod[name];
  }

  if (typeof mod?.default === "function") return mod.default;
  return null;
}

function buildMinimalCandidateBundle({
  normalizedQueryTables = [],
  source = "candidate-generation-minimal-fallback",
} = {}) {
  const {
    buildAnalysisRecipeCandidates,
  } = require("../analysisRecipeCandidateBuilder");
  const {
    buildBusinessTemplateCandidates,
  } = require("../businessTemplateConfig");

  const analysisRecipeCandidates = buildAnalysisRecipeCandidates(
    normalizedQueryTables,
  );

  const businessTemplateCandidates = buildBusinessTemplateCandidates(
    analysisRecipeCandidates,
  );

  return {
    analysisRecipeCandidates,
    categoryCandidates: [],
    businessTemplateCandidates,
    candidateGeneration: {
      version: "candidate_generation_v1",
      source,
      deterministic: {
        used: true,
        fallback: true,
      },
      aiReranker: {
        enabled: false,
        used: false,
        skippedReason: "MINIMAL_FALLBACK",
      },
      validation: {
        used: false,
        skippedReason: "VALIDATOR_NOT_AVAILABLE_IN_MINIMAL_FALLBACK",
      },
      generatedAt: new Date().toISOString(),
    },
  };
}

async function generateCandidateBundle({
  normalizedQueryTables = [],
  fileName = "",
  source = "candidate-generation",
} = {}) {
  let bundle = null;

  try {
    const deterministicModule = require("./deterministicCandidateBuilder");
    const buildDeterministicCandidateBundle = resolveExportedFunction(
      deterministicModule,
      [
        "buildDeterministicCandidateBundle",
        "generateDeterministicCandidateBundle",
        "buildCandidateBundle",
      ],
    );

    if (buildDeterministicCandidateBundle) {
      bundle = await buildDeterministicCandidateBundle({
        normalizedQueryTables,
        fileName,
        source,
      });
    }
  } catch (error) {
    console.warn(
      "[candidateGeneration] deterministic builder failed:",
      error?.message || error,
    );
  }

  if (!bundle) {
    bundle = buildMinimalCandidateBundle({
      normalizedQueryTables,
      source,
    });
  }

  try {
    const matcherModule = require("./aiTemplateMatcher");
    const matchBusinessTemplatesWithAi = resolveExportedFunction(
      matcherModule,
      ["matchBusinessTemplatesWithAi", "aiTemplateMatcher", "match"],
    );

    if (matchBusinessTemplatesWithAi) {
      const matched = await matchBusinessTemplatesWithAi({
        normalizedQueryTables,
        bundle,
        fileName,
      });

      if (matched?.bundle) {
        bundle = {
          ...matched.bundle,
          candidateGeneration: {
            ...(matched.bundle.candidateGeneration || {}),
            aiTemplateMatcher: matched.meta ||
              matched.bundle.candidateGeneration?.aiTemplateMatcher || {
                enabled: true,
                used: true,
              },
          },
        };
      }
    }
  } catch (error) {
    const enabled = ["true", "1"].includes(
      String(process.env.USE_AI_TEMPLATE_MATCHER || "").toLowerCase(),
    );

    bundle = {
      ...bundle,
      candidateGeneration: {
        ...(bundle.candidateGeneration || {}),
        aiTemplateMatcher: {
          enabled,
          used: false,
          error: error?.message || String(error),
        },
      },
    };
  }

  try {
    const rerankerModule = require("./aiCandidateReranker");
    const rerankCandidateBundle = resolveExportedFunction(rerankerModule, [
      "rerankCandidateBundle",
      "aiCandidateReranker",
      "rerank",
    ]);

    if (rerankCandidateBundle) {
      const reranked = await rerankCandidateBundle({
        normalizedQueryTables,
        bundle,
        fileName,
      });

      if (reranked?.bundle) {
        bundle = {
          ...reranked.bundle,
          candidateGeneration: {
            ...(reranked.bundle.candidateGeneration || {}),
            aiReranker: reranked.meta ||
              reranked.bundle.candidateGeneration?.aiReranker || {
                enabled: true,
                used: true,
              },
          },
        };
      }
    }
  } catch (error) {
    bundle = {
      ...bundle,
      candidateGeneration: {
        ...(bundle.candidateGeneration || {}),
        aiReranker: {
          enabled: process.env.USE_AI_CANDIDATE_RERANKER === "true",
          used: false,
          error: error?.message || String(error),
        },
      },
    };
  }

  try {
    const validatorModule = require("./candidateValidator");
    const validateCandidateBundle = resolveExportedFunction(validatorModule, [
      "validateCandidateBundle",
      "validate",
    ]);

    if (validateCandidateBundle) {
      bundle = await validateCandidateBundle(bundle, normalizedQueryTables);
    }
  } catch (error) {
    bundle = {
      ...bundle,
      candidateGeneration: {
        ...(bundle.candidateGeneration || {}),
        validation: {
          used: false,
          error: error?.message || String(error),
        },
      },
    };
  }

  return {
    ...bundle,
    analysisRecipeCandidates: Array.isArray(bundle.analysisRecipeCandidates)
      ? bundle.analysisRecipeCandidates
      : [],
    categoryCandidates: Array.isArray(bundle.categoryCandidates)
      ? bundle.categoryCandidates
      : [],
    businessTemplateCandidates: Array.isArray(bundle.businessTemplateCandidates)
      ? bundle.businessTemplateCandidates
      : [],
    candidateGeneration: {
      ...(bundle.candidateGeneration || {}),
      version: bundle.candidateGeneration?.version || "candidate_generation_v1",
      fileName,
      source,
      generatedAt:
        bundle.candidateGeneration?.generatedAt || new Date().toISOString(),
    },
  };
}

module.exports = generateCandidateBundle;
module.exports.generateCandidateBundle = generateCandidateBundle;
module.exports.default = generateCandidateBundle;
