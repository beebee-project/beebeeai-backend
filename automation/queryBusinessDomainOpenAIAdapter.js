const {
  DEFAULT_MODEL,
  DEFAULT_REASONING_EFFORT,
  MODEL_OUTPUT_SCHEMA,
} = require("./queryBusinessDomainProfiler");
const {
  QUERY_BUSINESS_DOMAIN_PROMPT_VERSION,
  buildBusinessDomainSystemPrompt,
  buildBusinessDomainUserPrompt,
} = require("./queryBusinessDomainPrompt");

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
      `OpenAI 업무 영역 분류가 거부되었습니다: ${refusal}`,
    );
    error.code = "OPENAI_BUSINESS_DOMAIN_REFUSAL";
    throw error;
  }
  const text = response.output_text;
  if (typeof text !== "string" || !text.trim()) {
    const error = new Error("OpenAI Responses API의 output_text가 없습니다.");
    error.code = "OPENAI_BUSINESS_DOMAIN_OUTPUT_TEXT_MISSING";
    throw error;
  }
  try {
    return JSON.parse(text);
  } catch (cause) {
    const error = new Error("OpenAI 업무 영역 JSON 파싱에 실패했습니다.");
    error.code = "OPENAI_BUSINESS_DOMAIN_JSON_PARSE_FAILED";
    error.cause = cause;
    throw error;
  }
}

function buildBusinessDomainOpenAIRequest({
  input,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  promptCacheKey = `beebee-business-domain-${QUERY_BUSINESS_DOMAIN_PROMPT_VERSION}`,
  safetyIdentifier,
  maxOutputTokens = 2400,
} = {}) {
  if (!input || typeof input !== "object") {
    throw new TypeError("업무 영역 input 객체가 필요합니다.");
  }
  const request = {
    model,
    store: false,
    reasoning: {
      effort: reasoningEffort,
    },
    input: [
      {
        role: "system",
        content: [
          {
            type: "input_text",
            text: buildBusinessDomainSystemPrompt(),
          },
        ],
      },
      {
        role: "user",
        content: [
          {
            type: "input_text",
            text: buildBusinessDomainUserPrompt(input),
          },
        ],
      },
    ],
    text: {
      verbosity: "low",
      format: {
        type: "json_schema",
        name: "query_business_domain_model_output_v1",
        description: "BeeBee AI 업무 영역과 데이터셋 목적 분류 결과",
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

function createOpenAIBusinessDomainProvider({
  client,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  promptCacheKey,
  safetyIdentifier,
  maxOutputTokens = 2400,
} = {}) {
  if (!client?.responses || typeof client.responses.create !== "function") {
    throw new TypeError("client.responses.create 함수가 필요합니다.");
  }
  return {
    async profile({
      input,
      model: requestedModel,
      reasoningEffort: requestedReasoningEffort,
    } = {}) {
      const effectiveModel = requestedModel || model;
      const effectiveReasoning = requestedReasoningEffort || reasoningEffort;
      const request = buildBusinessDomainOpenAIRequest({
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
    const error = new Error(
      "openai 패키지가 필요합니다. 백엔드 의존성에 openai를 설치해 주세요.",
    );
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
  buildBusinessDomainOpenAIRequest,
  createOpenAIBusinessDomainProvider,
  createOpenAIClientFromEnvironment,
};
