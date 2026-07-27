const path = require("path");
const { executeAnalysisRecipeCandidate } = require("./analysisRecipeExecutor");
const {
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  augmentBusinessTemplateResult,
} = require("./semanticOutputPlanner");

const BUSINESS_SEMANTIC_AUGMENTATION_VERSION =
  "business_semantic_augmentation_v2_observed";
const BUSINESS_TEMPLATE_EXECUTOR_VERSION =
  "business_template_executor_v2_route_observed";

function positiveInteger(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 ? Math.floor(parsed) : fallback;
}

function normalizeBooleanFlag(value) {
  if (value === true) return true;
  if (value === false) return false;
  return null;
}

function semanticAugmentationDecision(templateCandidate = {}) {
  const environmentValue = String(
    process.env.SEMANTIC_PLANNER_BUSINESS_AUGMENT ?? "",
  ).trim();
  const candidateFlag = normalizeBooleanFlag(
    templateCandidate.semanticPlannerAugment,
  );

  if (candidateFlag === false) {
    return {
      enabled: false,
      environmentValue,
      candidateFlag,
      skippedReason: "CANDIDATE_DISABLED",
    };
  }

  if (environmentValue === "0") {
    return {
      enabled: false,
      environmentValue,
      candidateFlag,
      skippedReason: "ENV_DISABLED",
    };
  }

  return {
    enabled: true,
    environmentValue,
    candidateFlag,
    skippedReason: "",
  };
}

function semanticAugmentationEnabled(templateCandidate = {}) {
  return semanticAugmentationDecision(templateCandidate).enabled;
}

function modulePathForObservation(filePath = __filename) {
  const relative = path.relative(process.cwd(), filePath);
  return String(relative || filePath).replace(/\\/g, "/");
}

function buildExecutorObservationMeta({
  templateCandidate = {},
  normalizedQueryTables = [],
  decision = semanticAugmentationDecision(templateCandidate),
} = {}) {
  return {
    businessTemplateExecutorVersion: BUSINESS_TEMPLATE_EXECUTOR_VERSION,
    businessSemanticAugmentationVersion: BUSINESS_SEMANTIC_AUGMENTATION_VERSION,
    semanticOutputPlannerVersion: SEMANTIC_OUTPUT_PLANNER_VERSION,
    executorModulePath: modulePathForObservation(__filename),
    semanticAugmentationEnvironmentValue: decision.environmentValue,
    semanticPlannerCandidateFlag: decision.candidateFlag,
    semanticAugmentationEnabled: decision.enabled,
    semanticAugmentationSkippedReason: decision.skippedReason,
    executorInputTableCount: Array.isArray(normalizedQueryTables)
      ? normalizedQueryTables.length
      : 0,
  };
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
  const decision = semanticAugmentationDecision(templateCandidate);
  const observation = buildExecutorObservationMeta({
    templateCandidate,
    normalizedQueryTables,
    decision,
  });

  if (!decision.enabled) {
    return {
      ...result,
      executionMeta: {
        ...(result.executionMeta || {}),
        ...observation,
        semanticBusinessAugmentation: false,
      },
    };
  }

  try {
    const augmented = augmentBusinessTemplateResult({
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

    return {
      ...augmented,
      executionMeta: {
        ...(augmented.executionMeta || {}),
        ...observation,
        semanticBusinessAugmentation: true,
        semanticAugmentationSkippedReason: "",
      },
    };
  } catch (error) {
    console.warn(
      "[semantic-planner] business augmentation failed:",
      error?.message || error,
    );

    return {
      ...result,
      executionMeta: {
        ...(result.executionMeta || {}),
        ...observation,
        semanticBusinessAugmentation: false,
        semanticAugmentationSkippedReason: "EXECUTION_ERROR",
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
      executionMeta: buildExecutorObservationMeta({
        templateCandidate,
        normalizedQueryTables,
      }),
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
      executionMeta: buildExecutorObservationMeta({
        templateCandidate,
        normalizedQueryTables,
      }),
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
  BUSINESS_TEMPLATE_EXECUTOR_VERSION,
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  buildExecutorObservationMeta,
  executeBusinessTemplate,
  executeTemplateSections,
  semanticAugmentationDecision,
  semanticAugmentationEnabled,
};
