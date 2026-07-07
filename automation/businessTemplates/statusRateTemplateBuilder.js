const {
  findTableForTemplate,
  executeTemplateSections,
} = require("./commonTemplateHelpers");
const {
  buildStatusRateReportSections,
} = require("../structuralBuilders/statusRateReportBuilder");

const STATUS_RATE_TEMPLATE_BUILDER_VERSION = "status_rate_template_builder_v1";

function buildStatusRateConfig({ definition = {}, templateCandidate = {} } = {}) {
  return {
    templateId: templateCandidate.templateId || definition.templateId || "",
    title: templateCandidate.title || definition.title || "상태 처리율 보고서",
    description: templateCandidate.description || definition.description || "",
    hints: {
      status: definition.statusHeaderHints || [],
      date: definition.dateHeaderHints || [],
      metric: definition.metricHeaderHints || [],
      department: definition.departmentHeaderHints || [],
      owner: definition.ownerHeaderHints || [],
      category: definition.categoryHeaderHints || [],
    },
    sectionIds: definition.statusRateSectionIds || {},
    titles: definition.statusRateTitles || {},
    labels: definition.statusRateLabels || {},
    version: STATUS_RATE_TEMPLATE_BUILDER_VERSION,
  };
}

function executeStatusRateReport({
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

  const sections = buildStatusRateReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: buildStatusRateConfig({ definition, templateCandidate }),
  });

  if (sections.length) return sections;

  return executeTemplateSections({
    normalizedQueryTables,
    templateCandidate,
  });
}

module.exports = {
  STATUS_RATE_TEMPLATE_BUILDER_VERSION,
  buildStatusRateConfig,
  executeStatusRateReport,
};
