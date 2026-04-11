const { parseMacroIntent } = require("./intentParser");
const { buildVbaScript } = require("./vbaBuilder");
const { buildAppsScript } = require("./appScriptBuilder");
const OpenAI = require("openai");
let client = null;
if (process.env.OPENAI_API_KEY) {
  client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

const ALLOWED_TYPES = new Set([
  "formatRange",
  "setValue",
  "copyRange",
  "clearRange",
  "moveRange",
  "removeDuplicates",
  "sortRange",
  "filterRange",
  "insertRow",
  "deleteRow",
  "insertColumn",
  "deleteColumn",
  "createSheet",
  "duplicateSheet",
  "renameSheet",
  "deleteSheet",
  "activateSheet",
  "unknown",
]);

function isValidRangeRef(value) {
  if (!value || typeof value !== "string") return false;
  const v = value.trim();
  if (v === "__USED_RANGE__") return true;
  return (
    /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i.test(v) ||
    /^[A-Z]+[0-9]+$/i.test(v) ||
    /^[A-Z]+:[A-Z]+$/i.test(v) ||
    /^[0-9]+:[0-9]+$/i.test(v)
  );
}

function sanitizeColumn(column) {
  if (!column || typeof column !== "object") return undefined;

  const out = { letter: null, index: null };

  if (
    typeof column.letter === "string" &&
    /^[A-Z]+$/i.test(column.letter.trim())
  ) {
    out.letter = column.letter.trim().toUpperCase();
  } else if (Number.isInteger(column.index) && column.index > 0) {
    out.index = column.index;
  } else if (
    typeof column.index === "string" &&
    /^[0-9]+$/.test(column.index) &&
    Number(column.index) > 0
  ) {
    out.index = Number(column.index);
  }

  if (!out.letter && !out.index) return undefined;
  return out;
}

function sanitizeStyle(style) {
  if (!style || typeof style !== "object") return undefined;
  const out = {};

  if (style.bold === true) out.bold = true;
  if (style.italic === true) out.italic = true;
  if (style.underline === true) out.underline = true;
  if (style.border) out.border = "thin";

  if (
    typeof style.fillColor === "string" &&
    /^#[0-9A-Fa-f]{6}$/.test(style.fillColor.trim())
  ) {
    out.fillColor = style.fillColor.trim().toUpperCase();
  }

  if (
    typeof style.fontColor === "string" &&
    /^#[0-9A-Fa-f]{6}$/.test(style.fontColor.trim())
  ) {
    out.fontColor = style.fontColor.trim().toUpperCase();
  }

  const alignMap = {
    left: "Left",
    center: "Center",
    right: "Right",
  };
  if (typeof style.horizontalAlign === "string") {
    const normalized = alignMap[style.horizontalAlign.trim().toLowerCase()];
    if (normalized) out.horizontalAlign = normalized;
  }

  return Object.keys(out).length ? out : undefined;
}

function sanitizeIntent(raw) {
  if (!raw || typeof raw !== "object") {
    return { type: "unknown" };
  }

  const type = typeof raw.type === "string" ? raw.type.trim() : "unknown";
  if (!ALLOWED_TYPES.has(type)) {
    return { type: "unknown" };
  }

  const intent = { type };

  if (typeof raw.text === "string" && raw.text.trim()) {
    intent.text = raw.text.trim();
  }

  if (raw.target && typeof raw.target === "object") {
    const range = raw.target.range;
    if (isValidRangeRef(range)) {
      intent.target = { range: range.trim() };
    }
  }

  const column = sanitizeColumn(raw.column);
  if (column) intent.column = column;

  const style = sanitizeStyle(raw.style);
  if (style) intent.style = style;

  if (
    typeof raw.direction === "string" &&
    ["ascending", "descending"].includes(raw.direction.trim().toLowerCase())
  ) {
    intent.direction = raw.direction.trim().toLowerCase();
  }

  if (typeof raw.criteria === "string") {
    intent.criteria = raw.criteria.trim();
  }

  if (typeof raw.value === "string" || typeof raw.value === "number") {
    intent.value = raw.value;
  }

  const positiveInt = (v) =>
    Number.isInteger(v) && v > 0
      ? v
      : typeof v === "string" && /^[0-9]+$/.test(v) && Number(v) > 0
        ? Number(v)
        : null;

  const rowIndex = positiveInt(raw.rowIndex);
  if (rowIndex) intent.rowIndex = rowIndex;

  if (typeof raw.position === "string") {
    const p = raw.position.trim().toLowerCase();
    if (["left", "right", "above", "below"].includes(p)) {
      intent.position = p;
    }
  }

  ["name", "fromName", "toName", "from", "to"].forEach((key) => {
    if (typeof raw[key] === "string" && raw[key].trim()) {
      if (
        (key === "from" || key === "to") &&
        !isValidRangeRef(raw[key].trim())
      ) {
        return;
      }
      intent[key] = raw[key].trim();
    }
  });

  if (typeof raw.hasHeader === "boolean") {
    intent.hasHeader = raw.hasHeader;
  }

  return intent.type ? intent : { type: "unknown" };
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
    intent = sanitizeIntent(await llmMacroParser(prompt));
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

반드시 JSON 객체만 출력하세요. 설명 문장, 코드블록, 주석은 금지합니다.
type 은 아래 중 하나만 허용됩니다:
formatRange, setValue, copyRange, clearRange, moveRange, removeDuplicates,
sortRange, filterRange, insertRow, deleteRow, insertColumn, deleteColumn,
createSheet, duplicateSheet, renameSheet, deleteSheet, activateSheet, unknown

형식:
{
  "type": "<formatRange | insertRow | createSheet | duplicateSheet | sortRange | filterRange | ...>",
  "target": {
    "range": "A1:B10" 또는 "B:B" 또는 "__USED_RANGE__"
  },
  "column": {
    "letter": "A"
  },
  "style": {
    "bold": true/false,
    "fillColor": "#FFFF00"
  },
  "direction": "ascending" 또는 "descending",
  "position": "above" 또는 "below" 또는 "left" 또는 "right",
  "name": "<시트 이름>",
  "fromName": "<기존 시트 이름>",
  "toName": "<새 시트 이름>",
  "criteria": "<필터 값>",
  "value": "<입력 값>",
  "rowIndex": 3,
  "hasHeader": true
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
    return sanitizeIntent(JSON.parse(response.choices[0].message.content));
  } catch {
    return { type: "unknown" };
  }
}
