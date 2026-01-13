const macroService = require("../macro/index");
const { assertCanUse, bumpUsage } = require("../services/usageService");

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

    // ✅ 성공 시 카운트 증가
    if (req.user?.id) {
      await bumpUsage(req.user.id, "formulaConversions", 1);
    }

    res.json(result);
  } catch (e) {
    console.error("[generateMacro] error:", e);
    // usage 제한 에러는 그대로 노출(프론트가 메시지 처리하기 좋게)
    const status = e?.status || 500;
    res.status(status).json({
      error:
        e?.code === "LIMIT_EXCEEDED"
          ? "USAGE_LIMIT"
          : "Macro generation failed",
      code: e?.code,
      meta: e?.meta,
    });
  }
};
