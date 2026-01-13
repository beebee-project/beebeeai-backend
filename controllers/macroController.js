const macroService = require("../macro/index");

exports.generateMacro = async (req, res) => {
  try {
    const { prompt, target } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: "prompt is required" });
    }

    // ← target도 함께 전달
    const result = await macroService.generate({ prompt, target });

    res.json(result);
  } catch (e) {
    console.error("[generateMacro] error:", e);
    res.status(500).json({ error: "Macro generation failed" });
  }
};
