const { parseMacroIntent } = require("./intentParser");
const { buildVbaScript } = require("./vbaBuilder");
const { buildAppsScript } = require("./appScriptBuilder");
const OpenAI = require("openai");
let client = null;
if (process.env.OPENAI_API_KEY) {
  client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

/**
 * 매크로 코드 생성 엔트리
 * @param {{ prompt: string, target?: "vba" | "appsScript" | "officeScript" }} param0
 */
exports.generate = async ({ prompt, target }) => {
  // 기본값은 VBA
  let macroTarget = "vba";
  if (target === "appsScript") {
    macroTarget = "appsScript";
  } else if (target === "officeScript" || target === "vba" || !target) {
    macroTarget = "vba";
  }

  // 1) 규칙 기반 파서 먼저 시도
  let intent = parseMacroIntent(prompt);

  // 2) 규칙 기반이 unknown이면 GPT fallback
  if (intent.type === "unknown") {
    intent = await llmMacroParser(prompt);
  }

  // 3) intent → 코드 생성 (VBA / Apps 분기)
  let code;
  if (macroTarget === "appsScript") {
    code = buildAppsScript(intent);
  } else {
    code = buildVbaScript(intent);
  }

  return { intent, code, target: macroTarget };
};

async function llmMacroParser(prompt) {
  if (!client) {
    return { type: "unknown" };
  }
  const systemPrompt = `
당신은 스프레드시트 매크로(VBA / Google Apps Script) 생성을 위한 Intent Parser 입니다.
사용자의 자연어 명령을 다음 Intent JSON 형태로 변환하세요.

형식:
{
  "type": "<formatRange | insertRow | createSheet | duplicateSheet | sortRange | filterRange | ...>",
  "target": {
    "range": "A1:B10" 또는 "B:B"
  },
  "style": {
    "bold": true/false,
    "fillColor": "#FFFF00"
  },
  "position": "above" 또는 "below",
  "name": "<시트 이름>"
}

지원되지 않는 명령은:
{
  "type": "unknown"
}
`;

  const userPrompt = `사용자 입력: "${prompt}"`;

  const response = await client.chat.completions.create({
    model: "gpt-4o-mini",
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
  });

  try {
    return JSON.parse(response.choices[0].message.content);
  } catch {
    return { type: "unknown" };
  }
}
