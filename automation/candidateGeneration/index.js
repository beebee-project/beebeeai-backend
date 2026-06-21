const { stableCandidateId } = require("./candidateValidator");

function isEnabled() {
  return String(process.env.USE_AI_CANDIDATE_RERANKER || "") === "true" ||
    process.env.USE_AI_CANDIDATE_RERANKER === "1";
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
        (b.priority + (b._aiPriorityBoost || 0)) -
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
            task:
              "Rerank and explain existing candidates. Keep ids/templateIds unchanged. priorityBoost must be an integer from -5 to 5.",
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
    console.warn("[aiCandidateReranker] fallback deterministic:", error.message);
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

module.exports = {
  rerankCandidateBundle,
};
