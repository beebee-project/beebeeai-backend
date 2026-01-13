const macroService = require("../macro/index");
const { assertCanUse, bumpUsage } = require("../services/usageService");

function isUnsupportedMacro(result) {
  // 1) intent 기반
  if (result?.intent?.type === "unknown") return true;
  // 2) fallback 코드 텍스트 기반(지금 너 fallback에 있는 문구)
  const code = String(result?.code || "");
  if (code.includes("지원하지 않는 작업입니다")) return true;
  return false;
}

exports.generateMacro = async (req, res) => {
  try {
    const { prompt, target } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: "prompt is required" });
    }

    // ✅ FREE면 월 10회(변환 카운트)에 매크로도 포함
    if (req.user?.id) {
      await assertCanUse(req.user.id, "formulaConversions", 1);
    }

    // ← target도 함께 전달
    const result = await macroService.generate({ prompt, target });

    // ✅ 미지원/실패는 성공으로 치지 않음
    if (isUnsupportedMacro(result)) {
      return res.status(422).json({
        code: "UNSUPPORTED_MACRO",
        message:
          "현재 요청은 매크로 생성에서 지원하지 않습니다. 좀 더 구체적으로 입력해 주세요.",
      });
    }

    // ✅ 성공 시 카운트 증가
    if (req.user?.id) {
      await bumpUsage(req.user.id, "formulaConversions", 1);
    }

    res.json(result);
  } catch (e) {
    console.error("[generateMacro] error:", e);
    const status = e?.status || 500;
    res.status(status).json({
      code: e?.code || "MACRO_FAILED",
      message:
        e?.code === "LIMIT_EXCEEDED"
          ? "사용량 한도를 초과했습니다."
          : "매크로 생성에 실패했습니다.",
      meta: e?.meta,
    });
  }
};
