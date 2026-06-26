const { BUSINESS_TEMPLATE_DEFS } = require("../businessTemplateConfig");
const { stableCandidateId } = require("./candidateValidator");

const MAX_CONFIDENCE = 0.85;

function isEnabled() {
  const value = String(process.env.USE_AI_TEMPLATE_MATCHER || "").toLowerCase();
  return value === "true" || value === "1";
}

function getModelName() {
  return (
    process.env.CANDIDATE_MATCHER_MODEL ||
    process.env.CANDIDATE_RERANKER_MODEL ||
    "gpt-4o-mini"
  );
}

function getOpenAIClient() {
  if (!process.env.OPENAI_API_KEY) return null;

  try {
    const OpenAI = require("openai");
    return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  } catch (error) {
    console.warn("[aiTemplateMatcher] OpenAI init skipped:", error.message);
    return null;
  }
}

function cleanText(value = "", max = 300) {
  const s = String(value || "")
    .replace(/\s+/g, " ")
    .trim();
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

function clampConfidence(value, fallback = 0.6) {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(0, Math.min(MAX_CONFIDENCE, n));
}

function normalizeText(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function getRecipeType(candidate = {}) {
  return candidate.recipeType || candidate.type || candidate.recipeId || "";
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

function uniqueStrings(values = []) {
  const seen = new Set();
  const out = [];

  for (const value of values) {
    const text = cleanText(value, 100);
    if (!text || seen.has(text)) continue;
    seen.add(text);
    out.push(text);
  }

  return out;
}

function uniqueCandidates(candidates = []) {
  const seen = new Set();
  const out = [];

  for (const candidate of candidates) {
    const id = stableCandidateId(candidate);
    if (!id || seen.has(id)) continue;
    seen.add(id);
    out.push(candidate);
  }

  return out;
}

function safeHeaders(normalizedQueryTables = []) {
  return (Array.isArray(normalizedQueryTables) ? normalizedQueryTables : [])
    .slice(0, 8)
    .map((table) => ({
      tableId: table.tableId,
      sheetName: table.sheetName,
      tableName: table.tableName,
      rowCount: table.rowCount,
      columns: (table.columns || []).slice(0, 50).map((column) => ({
        header: column.header,
        type: column.type,
        role: column.role,
        uniqueCount: column.uniqueCount,
        uniqueRatio: column.uniqueRatio,
      })),
    }));
}

function compactTemplateDef(def = {}) {
  return {
    templateId: def.templateId,
    title: def.title,
    description: def.description,
    requiredRecipeTypes: def.requiredRecipeTypes || [],
    requiredAnyRecipeTypes: def.requiredAnyRecipeTypes || [],
    optionalRecipeTypes: def.optionalRecipeTypes || [],
    requiredHeaderHints: def.requiredHeaderHints || [],
    requiredAnyHeaderHints: def.requiredAnyHeaderHints || [],
    optionalHeaderHints: def.optionalHeaderHints || [],
    outputTypes: def.outputTypes || [],
  };
}

function compactAnalysisCandidate(candidate = {}) {
  return {
    id: stableCandidateId(candidate),
    recipeType: getRecipeType(candidate),
    tableId: candidate.tableId,
    title: candidate.title,
    description: candidate.description,
    columns: candidate.columns || {},
    dimensions: candidate.dimensions || [],
    metrics: candidate.metrics || [],
    dates: candidate.dates || [],
    statuses: candidate.statuses || [],
  };
}

function candidateMatchesTemplate(candidate = {}, def = {}) {
  const recipeType = getRecipeType(candidate);
  const allowedTypes = [
    ...(def.requiredRecipeTypes || []),
    ...(def.requiredAnyRecipeTypes || []),
    ...(def.optionalRecipeTypes || []),
  ];

  if (allowedTypes.includes(recipeType)) return true;

  const candidateText = normalizeText(
    [
      candidate.title,
      candidate.description,
      candidate.groupHeader,
      candidate.metricHeader,
      candidate.dateHeader,
      candidate.statusHeader,
      ...(Object.values(candidate.columns || {}) || []),
      ...(candidate.dimensions || []),
      ...(candidate.metrics || []),
      ...(candidate.dates || []),
      ...(candidate.statuses || []),
    ]
      .filter(Boolean)
      .join(" "),
  );

  const hints = [
    ...(def.requiredHeaderHints || []),
    ...(def.requiredAnyHeaderHints || []),
    ...(def.optionalHeaderHints || []),
  ].map(normalizeText);

  return hints.some(
    (hint) =>
      hint &&
      candidateText &&
      (candidateText.includes(hint) || hint.includes(candidateText)),
  );
}

function fallbackCandidatesForTemplate(
  def = {},
  analysisCandidates = [],
  limit = 4,
) {
  return (Array.isArray(analysisCandidates) ? analysisCandidates : [])
    .filter((candidate) => candidateMatchesTemplate(candidate, def))
    .slice(0, limit);
}

function normalizeAiItems(ai = {}) {
  if (Array.isArray(ai)) return ai;
  return Array.isArray(ai.businessTemplates) ? ai.businessTemplates : [];
}

function applyAiTemplateMatches(bundle = {}, ai = {}) {
  const analysisRecipeCandidates = Array.isArray(
    bundle.analysisRecipeCandidates,
  )
    ? bundle.analysisRecipeCandidates
    : [];
  const businessTemplateCandidates = Array.isArray(
    bundle.businessTemplateCandidates,
  )
    ? bundle.businessTemplateCandidates
    : [];

  const analysisById = new Map(
    analysisRecipeCandidates.map((candidate) => [
      stableCandidateId(candidate),
      candidate,
    ]),
  );
  const defsById = new Map(
    BUSINESS_TEMPLATE_DEFS.map((def) => [def.templateId, def]),
  );
  const businessByTemplateId = new Map(
    businessTemplateCandidates.map((candidate) => [
      String(candidate.templateId),
      candidate,
    ]),
  );

  const ignored = [];
  let added = 0;
  let updated = 0;

  for (const item of normalizeAiItems(ai)) {
    const templateId = String(item?.templateId || "").trim();
    const def = defsById.get(templateId);

    if (!def) {
      if (templateId)
        ignored.push({ templateId, reason: "UNKNOWN_TEMPLATE_ID" });
      continue;
    }

    const rawIds = Array.isArray(item.candidateIds)
      ? item.candidateIds.map((id) => String(id || "").trim()).filter(Boolean)
      : [];

    const linkedCandidates = rawIds
      .map((id) => analysisById.get(id))
      .filter(Boolean);
    const fallbackCandidates = rawIds.length
      ? []
      : fallbackCandidatesForTemplate(def, analysisRecipeCandidates);
    const safeCandidates = uniqueCandidates([
      ...linkedCandidates,
      ...fallbackCandidates,
    ]);

    if (!safeCandidates.length) {
      ignored.push({ templateId, reason: "NO_VALID_ANALYSIS_CANDIDATES" });
      continue;
    }

    const existing = businessByTemplateId.get(templateId);
    const matchedHeaders = uniqueStrings([
      ...(existing?.matchedHeaders || []),
      ...(Array.isArray(item.matchedHeaders) ? item.matchedHeaders : []),
    ]).slice(0, 12);

    const aiCandidate = {
      type: "businessTemplate",
      templateId: def.templateId,
      title: def.title,
      description: cleanText(
        item.description || existing?.description || def.description,
        500,
      ),
      outputTypes: def.outputTypes,
      priority: Number(def.priority || 0) - (existing ? 0 : 0.25),
      confidence: clampConfidence(item.confidence, existing?.confidence || 0.6),
      matchedHeaders,
      matchedCount: safeCandidates.length,
      candidates: safeCandidates,
      primaryCandidate: safeCandidates[0] || null,
      recommendationReason: cleanText(
        item.recommendationReason ||
          existing?.recommendationReason ||
          "AI가 헤더 유사도를 기준으로 업무 템플릿 후보를 보조 매칭했습니다.",
        500,
      ),
      aiAssisted: true,
    };

    if (existing) {
      const mergedCandidates = uniqueCandidates([
        ...(existing.candidates || []),
        ...safeCandidates,
      ]);

      businessByTemplateId.set(templateId, {
        ...existing,
        description: aiCandidate.description || existing.description,
        confidence: Math.max(
          Number(existing.confidence || 0),
          aiCandidate.confidence,
        ),
        matchedHeaders,
        matchedCount: Math.max(
          Number(existing.matchedCount || 0),
          mergedCandidates.length,
        ),
        candidates: mergedCandidates,
        primaryCandidate:
          existing.primaryCandidate || mergedCandidates[0] || null,
        recommendationReason:
          aiCandidate.recommendationReason || existing.recommendationReason,
        aiAssisted: true,
      });
      updated += 1;
    } else {
      businessByTemplateId.set(templateId, aiCandidate);
      added += 1;
    }
  }

  const businessTemplates = Array.from(businessByTemplateId.values()).sort(
    (a, b) =>
      Number(b.priority || 0) - Number(a.priority || 0) ||
      Number(b.confidence || 0) - Number(a.confidence || 0),
  );

  return {
    bundle: {
      ...bundle,
      businessTemplateCandidates: businessTemplates,
      candidateGeneration: {
        ...(bundle.candidateGeneration || {}),
        aiTemplateMatcher: {
          enabled: true,
          used: added > 0 || updated > 0,
          added,
          updated,
          ignored,
        },
      },
    },
    meta: {
      enabled: true,
      used: added > 0 || updated > 0,
      added,
      updated,
      ignored,
    },
  };
}

async function matchBusinessTemplatesWithAi({
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
        skippedReason: "USE_AI_TEMPLATE_MATCHER_DISABLED",
      },
    };
  }

  if (
    !Array.isArray(bundle.analysisRecipeCandidates) ||
    !bundle.analysisRecipeCandidates.length
  ) {
    return {
      bundle,
      meta: {
        enabled: true,
        used: false,
        skippedReason: "NO_ANALYSIS_CANDIDATES",
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
    templateDefinitions: BUSINESS_TEMPLATE_DEFS.map(compactTemplateDef),
    analysisCandidates: bundle.analysisRecipeCandidates
      .slice(0, 40)
      .map(compactAnalysisCandidate),
  };

  try {
    const model = getModelName();
    const completion = await client.chat.completions.create({
      model,
      temperature: 0.1,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content:
            "You are an AI-assisted matcher for spreadsheet business template candidates. You may only recommend existing templateIds and existing analysis candidateIds from the provided payload. Do not calculate data, create rows, create sections, create formulas, or generate code. Return Korean explanations and JSON only.",
        },
        {
          role: "user",
          content: JSON.stringify({
            task: "Match uploaded spreadsheet headers to existing business template definitions. Return at most 3 businessTemplates. Use only templateIds and candidateIds from payload. Prefer conservative matches. confidence must be 0 to 0.85.",
            schema: {
              businessTemplates: [
                {
                  templateId: "existing templateId only",
                  candidateIds: ["existing analysis candidate id only"],
                  matchedHeaders: ["source headers or close aliases"],
                  confidence: 0.7,
                  recommendationReason: "short Korean reason",
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

    const applied = applyAiTemplateMatches(bundle, parsed);

    return {
      bundle: applied.bundle,
      meta: {
        ...applied.meta,
        model,
        responseTemplateCount: normalizeAiItems(parsed).length,
      },
    };
  } catch (error) {
    console.warn("[aiTemplateMatcher] deterministic fallback:", error.message);
    return {
      bundle,
      meta: {
        enabled: true,
        used: false,
        skippedReason: "AI_TEMPLATE_MATCHER_FAILED",
        error: error.code || error.name || "ERROR",
      },
    };
  }
}

module.exports = matchBusinessTemplatesWithAi;
module.exports.matchBusinessTemplatesWithAi = matchBusinessTemplatesWithAi;
module.exports.aiTemplateMatcher = matchBusinessTemplatesWithAi;
module.exports.applyAiTemplateMatches = applyAiTemplateMatches;
module.exports.default = matchBusinessTemplatesWithAi;
