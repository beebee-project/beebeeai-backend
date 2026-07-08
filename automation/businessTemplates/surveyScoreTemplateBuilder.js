const {
  findTableForTemplate,
  executeTemplateSections,
} = require("./commonTemplateHelpers");
const {
  buildSurveyScoreReportSections,
} = require("../structuralBuilders/surveyScoreReportBuilder");

const SURVEY_SCORE_TEMPLATE_BUILDER_VERSION = "survey_score_template_builder_v1";

function buildSurveyScoreConfig({ definition = {}, templateCandidate = {} } = {}) {
  return {
    templateId: templateCandidate.templateId || definition.templateId || "",
    title: templateCandidate.title || definition.title || "설문·점수 분석 보고서",
    description: templateCandidate.description || definition.description || "",
    hints: {
      score: definition.scoreHeaderHints || [],
      question: definition.questionHeaderHints || [],
      respondent: definition.respondentHeaderHints || [],
      department: definition.departmentHeaderHints || [],
      category: definition.categoryHeaderHints || [],
      date: definition.dateHeaderHints || [],
      comment: definition.commentHeaderHints || [],
    },
    sectionIds: definition.surveyScoreSectionIds || {},
    titles: definition.surveyScoreTitles || {},
    labels: definition.surveyScoreLabels || {},
    version: SURVEY_SCORE_TEMPLATE_BUILDER_VERSION,
  };
}

function executeSurveyScoreReport({
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

  const sections = buildSurveyScoreReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: buildSurveyScoreConfig({ definition, templateCandidate }),
  });

  if (sections.length) return sections;

  return executeTemplateSections({
    normalizedQueryTables,
    templateCandidate,
  });
}

module.exports = {
  SURVEY_SCORE_TEMPLATE_BUILDER_VERSION,
  buildSurveyScoreConfig,
  executeSurveyScoreReport,
};
