const {
  getBusinessTemplateExecutor,
  findBusinessTemplateDefinition,
} = require("./businessTemplates/templateRegistry");
const {
  normalizeBusinessTemplateResult,
  validateBusinessTemplateResultContract,
} = require("./businessTemplateContract");

function executeBusinessTemplate({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const templateId = templateCandidate.templateId;

  if (!templateId) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_ID_REQUIRED",
      message: "templateId가 필요합니다.",
    };
  }

  const definition = findBusinessTemplateDefinition(templateId);
  const executeTemplate = getBusinessTemplateExecutor(templateId);

  const sections = executeTemplate({
    normalizedQueryTables,
    templateCandidate,
    definition,
  });

  if (!sections.length) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_EXECUTION_EMPTY",
      message: "실행 가능한 템플릿 섹션이 없습니다.",
    };
  }

  const normalized = normalizeBusinessTemplateResult(
    {
      ok: true,
      resultType: "businessTemplate",
      templateId,
      title: templateCandidate.title || definition?.title || templateId,
      description:
        templateCandidate.description || definition?.description || "",
      outputTypes: templateCandidate.outputTypes || definition?.outputTypes,
      sections,
    },
    templateCandidate,
  );

  const contract = validateBusinessTemplateResultContract(normalized);

  return {
    ...normalized,
    contract,
  };
}

module.exports = {
  executeBusinessTemplate,
};
