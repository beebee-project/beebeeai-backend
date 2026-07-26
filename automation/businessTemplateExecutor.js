const { executeAnalysisRecipeCandidate } = require("./analysisRecipeExecutor");
const {
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  augmentBusinessTemplateResult,
} = require("./semanticOutputPlanner");

const BUSINESS_SEMANTIC_AUGMENTATION_VERSION =
  "business_semantic_augmentation_v1";

function positiveInteger(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 ? Math.floor(parsed) : fallback;
}

function semanticAugmentationEnabled(templateCandidate = {}) {
  if (templateCandidate.semanticPlannerAugment === false) return false;
  return process.env.SEMANTIC_PLANNER_BUSINESS_AUGMENT !== "0";
}

function executeTemplateSections({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const candidates = Array.isArray(templateCandidate.candidates)
    ? templateCandidate.candidates
    : [];

  return candidates
    .map((candidate, index) => {
      const result = executeAnalysisRecipeCandidate({
        normalizedQueryTables,
        candidate,
      });

      if (!result?.ok) return null;

      return {
        sectionId:
          candidate.recipeType ||
          candidate.type ||
          candidate.recipeId ||
          `section_${index + 1}`,
        title:
          candidate.title ||
          candidate.name ||
          candidate.label ||
          `섹션 ${index + 1}`,
        candidate,
        result,
      };
    })
    .filter(Boolean);
}

function executeHrMonthlyReport(args) {
  return executeTemplateSections(args);
}

function augmentResult({ result, normalizedQueryTables, templateCandidate }) {
  if (!semanticAugmentationEnabled(templateCandidate)) return result;

  try {
    return augmentBusinessTemplateResult({
      executionResult: result,
      tables: normalizedQueryTables,
      templateCandidate,
      options: {
        maxDimensionsPerSeries: positiveInteger(
          process.env.SEMANTIC_PLANNER_BUSINESS_MAX_DIMENSIONS,
          8,
        ),
        maxAddedSections: positiveInteger(
          process.env.SEMANTIC_PLANNER_BUSINESS_MAX_ADDED_SECTIONS,
          64,
        ),
        maxPlannedSections: positiveInteger(
          process.env.SEMANTIC_PLANNER_BUSINESS_MAX_PLANNED_SECTIONS,
          120,
        ),
      },
      context: {
        augmentationVersion: BUSINESS_SEMANTIC_AUGMENTATION_VERSION,
      },
    });
  } catch (error) {
    console.warn(
      "[semantic-planner] business augmentation failed:",
      error?.message || error,
    );

    return {
      ...result,
      executionMeta: {
        ...(result.executionMeta || {}),
        semanticBusinessAugmentation: false,
        semanticOutputPlannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
        semanticAugmentationError: error?.message || String(error),
      },
    };
  }
}

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

  let sections = [];

  switch (templateId) {
    case "hr_monthly_report":
      sections = executeHrMonthlyReport({
        normalizedQueryTables,
        templateCandidate,
      });
      break;

    default:
      sections = executeTemplateSections({
        normalizedQueryTables,
        templateCandidate,
      });
      break;
  }

  if (!sections.length) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_EXECUTION_EMPTY",
      message: "실행 가능한 템플릿 섹션이 없습니다.",
    };
  }

  const result = {
    ok: true,
    resultType: "businessTemplate",
    templateId,
    title: templateCandidate.title || templateId,
    description: templateCandidate.description || "",
    sections,
  };

  return augmentResult({
    result,
    normalizedQueryTables,
    templateCandidate,
  });
}

module.exports = {
  BUSINESS_SEMANTIC_AUGMENTATION_VERSION,
  executeBusinessTemplate,
  executeTemplateSections,
  semanticAugmentationEnabled,
};
