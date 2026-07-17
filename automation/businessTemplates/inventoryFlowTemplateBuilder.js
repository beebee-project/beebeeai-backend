const {
  findTableForTemplate,
  executeTemplateSections,
} = require("./commonTemplateHelpers");
const {
  buildInventoryFlowReportSections,
} = require("../structuralBuilders/inventoryFlowReportBuilder");

const INVENTORY_FLOW_TEMPLATE_BUILDER_VERSION =
  "inventory_flow_template_builder_v2";

function buildInventoryFlowConfig({
  definition = {},
  templateCandidate = {},
} = {}) {
  return {
    templateId: templateCandidate.templateId || definition.templateId || "",
    title:
      templateCandidate.title || definition.title || "재고·입출고 흐름 분석",
    description: templateCandidate.description || definition.description || "",
    hints: {
      flowType: definition.flowTypeHeaderHints || [],
      inboundQuantity: definition.inboundQuantityHeaderHints || [],
      outboundQuantity: definition.outboundQuantityHeaderHints || [],
      stockQuantity: definition.stockQuantityHeaderHints || [],
      quantity: definition.quantityHeaderHints || [],
      amount: definition.amountHeaderHints || definition.valueHeaderHints || [],
      unitPrice: definition.unitPriceHeaderHints || [],
      value: definition.valueHeaderHints || [],
      category: definition.categoryHeaderHints || [],
      location: definition.locationHeaderHints || [],
      date: definition.dateHeaderHints || [],
      status: definition.statusHeaderHints || [],
    },
    sectionIds: definition.inventoryFlowSectionIds || {},
    titles: definition.inventoryFlowTitles || {},
    labels: definition.inventoryFlowLabels || {},
    version: INVENTORY_FLOW_TEMPLATE_BUILDER_VERSION,
  };
}

function executeInventoryFlowReport({
  normalizedQueryTables = [],
  templateCandidate = {},
  definition = {},
}) {
  const table = findTableForTemplate(normalizedQueryTables, templateCandidate);

  if (!table?.tableId) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const sections = buildInventoryFlowReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: buildInventoryFlowConfig({ definition, templateCandidate }),
  });

  if (sections.length) return sections;

  return executeTemplateSections({
    normalizedQueryTables,
    templateCandidate,
  });
}

module.exports = {
  INVENTORY_FLOW_TEMPLATE_BUILDER_VERSION,
  buildInventoryFlowConfig,
  executeInventoryFlowReport,
};
