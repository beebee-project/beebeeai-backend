const {
  QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION,
} = require("./queryCandidatePlanner");

const QUERY_CANDIDATE_PLANNER_PROMPT_VERSION =
  "query_candidate_planner_prompt_v1";

function buildCandidatePlannerSystemPrompt() {
  return [
    "You are the conditional candidate planner for BeeBee AI tabular automation.",
    "Return only the JSON object required by the supplied strict schema.",
    "You are called only when deterministic candidate coverage is insufficient.",
    "Propose at most three missing, useful, executable summary-sheet candidates.",
    "Use only allowed operation names from the input.",
    "Use only tableId and columnId values present in the input; never invent identifiers.",
    "Every proposal must use exactly one physical table and match the required operand kinds for the operation.",
    "Do not duplicate an existing ranked or unresolved candidate intent.",
    "Prefer a small set of materially distinct candidates over superficial variations.",
    "Do not include formulas, raw values, sample values, personal information, file names, or hidden reasoning.",
    "A proposal is not READY. It will re-enter deterministic Resolver, Family, Feasibility, and Ranker stages.",
    `Set version to ${QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION}.`,
    "Use NO_ADDITION when no safe non-duplicate proposal is supported by the supplied evidence.",
  ].join("\n");
}

function buildCandidatePlannerUserPrompt(input) {
  return JSON.stringify(
    {
      task: "Fill only material deterministic candidate coverage gaps.",
      promptVersion: QUERY_CANDIDATE_PLANNER_PROMPT_VERSION,
      privacyNotice: {
        rawRowsIncluded: false,
        sampleValuesIncluded: false,
        originalFileIncluded: false,
        fileNameIncluded: false,
      },
      plannerInput: input,
    },
    null,
    2,
  );
}

module.exports = {
  QUERY_CANDIDATE_PLANNER_PROMPT_VERSION,
  buildCandidatePlannerSystemPrompt,
  buildCandidatePlannerUserPrompt,
};
