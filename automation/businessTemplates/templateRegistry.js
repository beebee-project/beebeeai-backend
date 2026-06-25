const {
  BUSINESS_TEMPLATE_DEFS,
  getBusinessTemplateDefinitions,
  findBusinessTemplateDefinition,
} = require("./templateDefinitions");
const {
  BUSINESS_TEMPLATE_EXECUTORS,
  getBusinessTemplateExecutor,
  hasBusinessTemplateExecutor,
} = require("./templateExecutorRegistry");

function getBusinessTemplateRegistryItem(templateId = "") {
  const definition = findBusinessTemplateDefinition(templateId);
  const executor = getBusinessTemplateExecutor(templateId);

  return {
    templateId,
    definition,
    executor,
    hasCustomExecutor: hasBusinessTemplateExecutor(templateId),
  };
}

module.exports = {
  BUSINESS_TEMPLATE_DEFS,
  BUSINESS_TEMPLATE_EXECUTORS,
  getBusinessTemplateDefinitions,
  findBusinessTemplateDefinition,
  getBusinessTemplateExecutor,
  hasBusinessTemplateExecutor,
  getBusinessTemplateRegistryItem,
};
