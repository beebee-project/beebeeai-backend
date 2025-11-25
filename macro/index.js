const { parseMacroIntent } = require("./intentParser");
const { buildOfficeScript } = require("./officeScriptBuilder");

exports.generate = async ({ prompt }) => {
  const intent = parseMacroIntent(prompt);
  const code = buildOfficeScript(intent);

  return { intent, code };
};
