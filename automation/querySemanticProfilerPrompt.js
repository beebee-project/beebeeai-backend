const taxonomy = require("./querySemanticProfilerTaxonomy.json");

const QUERY_SEMANTIC_PROFILER_PROMPT_VERSION =
  "query_semantic_profiler_prompt_v1";

function buildSemanticProfilerSystemPrompt() {
  return [
    "You are the semantic profiler for BeeBee AI tabular automation.",
    "Return only the JSON object required by the supplied strict schema.",
    "Classify the business domain and dataset intent, assign a semantic interpretation to every included column, summarize every included table, and identify only materially supported relationships between tables.",
    "Use only the supplied taxonomy codes.",
    "Use only tableId and columnId values that exist in the input. Never invent identifiers.",
    "Provide exactly one tableSemantics item for every included table and exactly one columnSemantics item for every included column.",
    "Use KEEP when the deterministic role and metric family are already correct, REFINE for a more specific compatible interpretation, REPLACE only when the deterministic interpretation is materially wrong, and UNKNOWN when evidence is insufficient.",
    "Do not infer any sensitive personal attribute or identify a person.",
    "Do not create candidateId, recipeId, templates, formulas, rankings, or generation support decisions.",
    "Do not claim a join relationship without structural evidence. An explicit sourceTableId is strong evidence for SOURCE_DERIVATION or CROSS_TAB_TRANSFORM.",
    "When evidence is insufficient, use unknown/NONE/UNKNOWN values, lower confidence, add an ambiguity, and set requiresHumanReview to true when the uncertainty is material.",
    "Descriptions must be concise structural explanations, not hidden reasoning.",
  ].join("\n");
}

function buildSemanticProfilerUserPrompt(input) {
  return JSON.stringify(
    {
      task: "Produce one integrated semantic profile covering business domain, every included column, every included table, and supported table relationships.",
      promptVersion: QUERY_SEMANTIC_PROFILER_PROMPT_VERSION,
      taxonomy,
      privacyNotice: {
        rawRowsIncluded: false,
        sampleValuesIncluded: false,
        originalFileIncluded: false,
        fileNameIncluded: false,
      },
      dataset: input,
    },
    null,
    2,
  );
}

module.exports = {
  QUERY_SEMANTIC_PROFILER_PROMPT_VERSION,
  buildSemanticProfilerSystemPrompt,
  buildSemanticProfilerUserPrompt,
};
