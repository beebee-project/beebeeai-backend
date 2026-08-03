const {
  DEFAULT_MODEL,
  DEFAULT_REASONING_EFFORT,
  MODEL_OUTPUT_SCHEMA,
} = require("./queryCandidatePlanner");
const {
  QUERY_CANDIDATE_PLANNER_PROMPT_VERSION,
  buildCandidatePlannerSystemPrompt,
  buildCandidatePlannerUserPrompt,
} = require("./queryCandidatePlannerPrompt");

function extractRefusal(response = {}) {
  for (const item of Array.isArray(response.output) ? response.output : []) {
    for (const part of Array.isArray(item?.content) ? item.content : []) {
      if (part?.type === "refusal" && part.refusal) return String(part.refusal);
    }
  }
  return "";
}

function parseOutputText(response = {}) {
  const refusal = extractRefusal(response);
  if (refusal) {
    const error = new Error(
      `OpenAI Candidate Planner가 거부되었습니다: ${refusal}`,
    );
    error.code = "OPENAI_CANDIDATE_PLANNER_REFUSAL";
    throw error;
  }
  if (
    typeof response.output_text !== "string" ||
    !response.output_text.trim()
  ) {
    const error = new Error("OpenAI Responses API의 output_text가 없습니다.");
    error.code = "OPENAI_CANDIDATE_PLANNER_OUTPUT_TEXT_MISSING";
    throw error;
  }
  try {
    return JSON.parse(response.output_text);
  } catch (cause) {
    const error = new Error(
      "OpenAI Candidate Planner JSON 파싱에 실패했습니다.",
    );
    error.code = "OPENAI_CANDIDATE_PLANNER_JSON_PARSE_FAILED";
    error.cause = cause;
    throw error;
  }
}

function buildCandidatePlannerOpenAIRequest({
  input,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  promptCacheKey = `beebee-candidate-planner-${QUERY_CANDIDATE_PLANNER_PROMPT_VERSION}`,
  safetyIdentifier,
  maxOutputTokens = 2200,
} = {}) {
  if (!input || typeof input !== "object") {
    throw new TypeError("Candidate Planner input 객체가 필요합니다.");
  }
  const request = {
    model,
    store: false,
    reasoning: { effort: reasoningEffort },
    input: [
      {
        role: "system",
        content: [
          { type: "input_text", text: buildCandidatePlannerSystemPrompt() },
        ],
      },
      {
        role: "user",
        content: [
          { type: "input_text", text: buildCandidatePlannerUserPrompt(input) },
        ],
      },
    ],
    text: {
      verbosity: "low",
      format: {
        type: "json_schema",
        name: "query_candidate_planner_model_output_v1",
        description: "BeeBee AI 조건부 Candidate Planner 신규 후보 제안",
        strict: true,
        schema: MODEL_OUTPUT_SCHEMA,
      },
    },
    max_output_tokens: maxOutputTokens,
    prompt_cache_key: promptCacheKey,
  };
  if (safetyIdentifier) request.safety_identifier = String(safetyIdentifier);
  return request;
}

function createOpenAICandidatePlannerProvider({
  client,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  promptCacheKey,
  safetyIdentifier,
  maxOutputTokens = 2200,
} = {}) {
  if (!client?.responses || typeof client.responses.create !== "function") {
    throw new TypeError("client.responses.create 함수가 필요합니다.");
  }
  return {
    async plan({
      input,
      model: requestedModel,
      reasoningEffort: requestedReasoning,
    } = {}) {
      const effectiveModel = requestedModel || model;
      const effectiveReasoning = requestedReasoning || reasoningEffort;
      const request = buildCandidatePlannerOpenAIRequest({
        input,
        model: effectiveModel,
        reasoningEffort: effectiveReasoning,
        promptCacheKey,
        safetyIdentifier,
        maxOutputTokens,
      });
      const response = await client.responses.create(request);
      return {
        provider: "OPENAI_RESPONSES",
        model: response.model || effectiveModel,
        reasoningEffort: effectiveReasoning,
        responseId: response.id || "",
        output: parseOutputText(response),
        usage: response.usage || {},
        rawResponseStatus: response.status || "",
      };
    },
  };
}

function createOpenAIClientFromEnvironment() {
  let OpenAI;
  try {
    OpenAI = require("openai");
  } catch (cause) {
    const error = new Error("openai 패키지가 필요합니다.");
    error.code = "OPENAI_PACKAGE_MISSING";
    error.cause = cause;
    throw error;
  }
  if (!process.env.OPENAI_API_KEY) {
    const error = new Error("OPENAI_API_KEY 환경변수가 필요합니다.");
    error.code = "OPENAI_API_KEY_MISSING";
    throw error;
  }
  const OpenAIClient = OpenAI.default || OpenAI;
  return new OpenAIClient({ apiKey: process.env.OPENAI_API_KEY });
}

module.exports = {
  extractRefusal,
  parseOutputText,
  buildCandidatePlannerOpenAIRequest,
  createOpenAICandidatePlannerProvider,
  createOpenAIClientFromEnvironment,
};
