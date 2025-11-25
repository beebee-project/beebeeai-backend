const macroService = require("../services/macro/index");

exports.generateMacro = async (req, res) => {
  try {
    const { prompt } = req.body;
    if (!prompt) return res.status(400).json({ error: "prompt is required" });

    const result = await macroService.generate({ prompt });
    res.json(result);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Macro generation failed" });
  }
};
