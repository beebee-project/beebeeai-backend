const taxonomy = require("./queryBusinessDomainTaxonomy.json");

const QUERY_BUSINESS_DOMAIN_PROMPT_VERSION = "query_business_domain_prompt_v1";

function buildBusinessDomainSystemPrompt() {
  return [
    "You classify the business domain and dataset intent of tabular business data.",
    "Return only the JSON object required by the supplied strict schema.",
    "Use only the allowed domain, intent, evidence-signal, and ambiguity codes.",
    "Base the decision only on the supplied table IDs, column IDs, headers, types, deterministic semantic roles, metric families, and table structure.",
    "Never create a tableId or columnId that is not present in the input.",
    "Do not invent candidateId, recipeId, templates, formulas, or generation support.",
    "Do not infer sensitive personal attributes or identify any individual.",
    "When evidence is insufficient, choose UNKNOWN, lower confidence, explain the ambiguity, and set requiresHumanReview to true.",
    "Secondary domains must be materially supported, must differ from the primary domain, and should normally contain no more than two items.",
    "Evidence descriptions must be concise and refer to structural signals, not hidden reasoning.",
  ].join("\n");
}

function buildBusinessDomainUserPrompt(input) {
  return JSON.stringify(
    {
      task: "Classify the primary business domain and dataset intent.",
      promptVersion: QUERY_BUSINESS_DOMAIN_PROMPT_VERSION,
      taxonomy,
      privacyNotice: {
        rawRowsIncluded: false,
        sampleValuesIncluded: false,
        originalFileIncluded: false,
      },
      dataset: input,
    },
    null,
    2,
  );
}

module.exports = {
  QUERY_BUSINESS_DOMAIN_PROMPT_VERSION,
  buildBusinessDomainSystemPrompt,
  buildBusinessDomainUserPrompt,
};
