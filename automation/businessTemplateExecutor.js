"use strict";

const crypto = require("crypto");
const path = require("path");
const {
  executeAnalysisRecipeCandidate,
} = require("./analysisRecipeExecutor");
const {
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  augmentBusinessTemplateResult,
} = require("./semanticOutputPlanner");
const {
  buildContractDrivenSummarySections,
} = require("./contractDrivenSummaryRecipeBuilder");

const TEMPLATE_REGISTRY_MODULE_PATH = require.resolve(
  "./businessTemplates/templateRegistry",
);
const {
  getBusinessTemplateRegistryItem,
} = require(TEMPLATE_REGISTRY_MODULE_PATH);

const BUSINESS_SEMANTIC_AUGMENTATION_VERSION =
  "business_semantic_augmentation_v5_review_closure";
const BUSINESS_TEMPLATE_EXECUTOR_VERSION =
  "business_template_executor_v5_review_closure";
const BUSINESS_TEMPLATE_REGISTRY_BRIDGE_VERSION =
  "business_template_registry_bridge_v1";
const BUSINESS_TEMPLATE_CONTRACT_BRIDGE_VERSION =
  "business_template_contract_bridge_v1";
const BUSINESS_SECTION_FINALIZATION_VERSION =
  "business_section_finalization_v1_content_hash_snapshot_label";
const BUSINESS_SECTION_CONTENT_HASH_VERSION =
  "business_section_content_hash_v1";
const INVENTORY_SNAPSHOT_SUM_LABEL =
  "행별 재고 스냅샷 합계";

function positiveInteger(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0
    ? Math.floor(parsed)
    : fallback;
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

function executorName(executor) {
  if (typeof executor !== "function") return "";
  return String(executor.name || "anonymous");
}

function definitionExists(definition) {
  return Boolean(
    definition &&
      typeof definition === "object" &&
      Object.keys(definition).length,
  );
}

function buildRegistryObservation({
  registryItem = {},
  executor = null,
  sectionCount = 0,
  outputKind = "",
} = {}) {
  const definition = registryItem.definition || null;
  const custom = registryItem.hasCustomExecutor === true;

  return {
    businessTemplateRegistryBridgeVersion:
      BUSINESS_TEMPLATE_REGISTRY_BRIDGE_VERSION,
    businessTemplateRegistryConnected: true,
    businessTemplateRegistryModulePath:
      modulePathForObservation(TEMPLATE_REGISTRY_MODULE_PATH),
    businessTemplateDefinitionFound:
      definitionExists(definition),
    businessTemplateHasCustomExecutor: custom,
    businessTemplateRegisteredExecutorName:
      executorName(executor),
    businessTemplateImplementationLevel:
      String(definition?.implementationLevel || ""),
    businessTemplateExecutionPath:
      custom
        ? "template_registry_custom_executor"
        : "template_registry_fallback_executor",
    businessTemplateExecutorOutputKind: outputKind,
    dedicatedBaseSectionCount: Number(sectionCount || 0),
  };
}

function buildExecutorObservationMeta({
  templateCandidate = {},
  normalizedQueryTables = [],
  decision = semanticAugmentationDecision(templateCandidate),
  registryObservation = {},
} = {}) {
  return {
    businessTemplateExecutorVersion:
      BUSINESS_TEMPLATE_EXECUTOR_VERSION,
    businessSemanticAugmentationVersion:
      BUSINESS_SEMANTIC_AUGMENTATION_VERSION,
    semanticOutputPlannerVersion:
      SEMANTIC_OUTPUT_PLANNER_VERSION,
    executorModulePath:
      modulePathForObservation(__filename),
    semanticAugmentationEnvironmentValue:
      decision.environmentValue,
    semanticPlannerCandidateFlag:
      decision.candidateFlag,
    semanticAugmentationEnabled:
      decision.enabled,
    semanticAugmentationSkippedReason:
      decision.skippedReason,
    executorInputTableCount:
      Array.isArray(normalizedQueryTables)
        ? normalizedQueryTables.length
        : 0,
    ...registryObservation,
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

function normalizeRegistryExecutorOutput(output) {
  if (Array.isArray(output)) {
    return {
      ok: true,
      outputKind: "section_array",
      sections: output.filter(Boolean),
      baseResult: null,
    };
  }

  if (output && typeof output === "object") {
    if (output.ok === false) {
      return {
        ok: false,
        outputKind: "error_result",
        errorResult: output,
        sections: [],
        baseResult: null,
      };
    }

    if (Array.isArray(output.sections)) {
      return {
        ok: true,
        outputKind: "business_template_result",
        sections: output.sections.filter(Boolean),
        baseResult: output,
      };
    }
  }

  return {
    ok: false,
    outputKind: "unsupported_output",
    errorResult: {
      ok: false,
      code: "BUSINESS_TEMPLATE_EXECUTOR_OUTPUT_INVALID",
      message:
        "템플릿 registry executor가 section 배열 또는 businessTemplate 결과를 반환하지 않았습니다.",
    },
    sections: [],
    baseResult: null,
  };
}

function executeRegisteredTemplate({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const templateId = templateCandidate.templateId;
  let registryItem;

  try {
    registryItem =
      getBusinessTemplateRegistryItem(templateId) || {};
  } catch (error) {
    return {
      ok: false,
      registryItem: {},
      registryObservation: {
        businessTemplateRegistryBridgeVersion:
          BUSINESS_TEMPLATE_REGISTRY_BRIDGE_VERSION,
        businessTemplateRegistryConnected: false,
        businessTemplateRegistryModulePath:
          modulePathForObservation(TEMPLATE_REGISTRY_MODULE_PATH),
        businessTemplateRegistryError:
          error?.message || String(error),
      },
      errorResult: {
        ok: false,
        code: "BUSINESS_TEMPLATE_REGISTRY_LOOKUP_FAILED",
        message:
          error?.message || String(error),
      },
    };
  }

  const executor =
    typeof registryItem.executor === "function"
      ? registryItem.executor
      : executeTemplateSections;
  const definition =
    registryItem.definition &&
    typeof registryItem.definition === "object"
      ? registryItem.definition
      : {};

  let rawOutput;
  try {
    rawOutput = executor({
      normalizedQueryTables,
      templateCandidate,
      definition,
    });
  } catch (error) {
    const registryObservation = buildRegistryObservation({
      registryItem,
      executor,
      sectionCount: 0,
      outputKind: "executor_exception",
    });

    return {
      ok: false,
      registryItem,
      registryObservation: {
        ...registryObservation,
        businessTemplateRegistryExecutorError:
          error?.message || String(error),
      },
      errorResult: {
        ok: false,
        code: "BUSINESS_TEMPLATE_REGISTERED_EXECUTOR_FAILED",
        message:
          error?.message || String(error),
      },
    };
  }

  if (
    rawOutput &&
    typeof rawOutput.then === "function"
  ) {
    const registryObservation = buildRegistryObservation({
      registryItem,
      executor,
      sectionCount: 0,
      outputKind: "async_output_unsupported",
    });

    return {
      ok: false,
      registryItem,
      registryObservation,
      errorResult: {
        ok: false,
        code: "BUSINESS_TEMPLATE_ASYNC_EXECUTOR_UNSUPPORTED",
        message:
          "업무 템플릿 executor는 동기 결과를 반환해야 합니다.",
      },
    };
  }

  const normalized =
    normalizeRegistryExecutorOutput(rawOutput);
  const registryObservation = buildRegistryObservation({
    registryItem,
    executor,
    sectionCount: normalized.sections.length,
    outputKind: normalized.outputKind,
  });

  if (!normalized.ok) {
    return {
      ok: false,
      registryItem,
      registryObservation,
      errorResult: normalized.errorResult,
    };
  }

  return {
    ok: true,
    registryItem,
    registryObservation,
    sections: normalized.sections,
    baseResult: normalized.baseResult,
  };
}

function uniqueStrings(values = []) {
  return Array.from(
    new Set(
      (Array.isArray(values) ? values : [])
        .map((value) => String(value || "").trim())
        .filter(Boolean),
    ),
  );
}

function collectSectionMetricIds(section = {}) {
  return uniqueStrings([
    ...(Array.isArray(section.metricIds)
      ? section.metricIds
      : []),
    ...(Array.isArray(section.result?.metricIds)
      ? section.result.metricIds
      : []),
    ...(Array.isArray(section.result?.meta?.metricIds)
      ? section.result.meta.metricIds
      : []),
    ...(Array.isArray(section.candidate?.metricIds)
      ? section.candidate.metricIds
      : []),
  ]);
}

function renderedMetricIdsFromSections(sections = []) {
  return uniqueStrings(
    (Array.isArray(sections) ? sections : [])
      .flatMap(collectSectionMetricIds),
  );
}

function coverageWithoutSections(coverage = {}) {
  if (!coverage || typeof coverage !== "object") {
    return {};
  }

  const {
    sections,
    ...rest
  } = coverage;

  void sections;
  return rest;
}

function contractCoverageSupported(coverage = {}) {
  const status = String(
    coverage?.status || "",
  ).trim().toUpperCase();

  return Boolean(
    coverage &&
      typeof coverage === "object" &&
      status !== "UNSUPPORTED" &&
      status !== "SOURCE_UNRESOLVED" &&
      (
        Array.isArray(coverage.sections) &&
        coverage.sections.length > 0 ||
        Array.isArray(coverage.expectedMetricIds) &&
        coverage.expectedMetricIds.length > 0
      ),
  );
}

function buildContractCoverageObservation({
  coverage = {},
  sectionCount = 0,
  error = null,
} = {}) {
  return {
    businessTemplateContractBridgeVersion:
      BUSINESS_TEMPLATE_CONTRACT_BRIDGE_VERSION,
    contractSummaryCoverageExecuted: true,
    contractSummaryCoverageStatus:
      String(
        error
          ? "EXECUTION_ERROR"
          : coverage?.status || "UNSUPPORTED",
      ),
    contractSummaryRecipeVersion:
      String(coverage?.version || ""),
    contractSummaryCatalogVersion:
      String(coverage?.contractCatalogVersion || ""),
    contractSummaryContractId:
      String(coverage?.contractId || ""),
    contractSummarySelectedTableId:
      String(coverage?.selectedTableId || ""),
    contractSummarySelectedSheetName:
      String(coverage?.selectedSheetName || ""),
    contractSummarySectionCount:
      Number(sectionCount || 0),
    contractSummaryExpectedMetricIdCount:
      uniqueStrings(
        coverage?.expectedMetricIds,
      ).length,
    contractSummaryRenderedMetricIdCount:
      uniqueStrings(
        coverage?.renderedMetricIds,
      ).length,
    contractSummaryInactiveMetricIdCount:
      uniqueStrings(
        coverage?.inactiveMetricIds,
      ).length,
    contractSummaryErrorMetricIdCount:
      uniqueStrings(
        coverage?.errorMetricIds,
      ).length,
    contractSummaryCoverageSupported:
      contractCoverageSupported(coverage),
    contractSummaryCoverageError:
      error
        ? error?.message || String(error)
        : "",
  };
}

function executeContractCoverage({
  normalizedQueryTables = [],
  templateId = "",
} = {}) {
  try {
    const coverage =
      buildContractDrivenSummarySections({
        normalizedQueryTables,
        templateId,
      }) || {};

    const sections = Array.isArray(
      coverage.sections,
    )
      ? coverage.sections.filter(Boolean)
      : [];

    return {
      coverage,
      sections,
      observation:
        buildContractCoverageObservation({
          coverage,
          sectionCount: sections.length,
        }),
    };
  } catch (error) {
    return {
      coverage: {
        version: "",
        contractCatalogVersion: "",
        templateId,
        status: "EXECUTION_ERROR",
        expectedMetricIds: [],
        renderedMetricIds: [],
        inactiveMetricIds: [],
        errorMetricIds: [],
      },
      sections: [],
      observation:
        buildContractCoverageObservation({
          coverage: {},
          sectionCount: 0,
          error,
        }),
    };
  }
}

function mergeContractAndSemanticCoverage({
  contractCoverage = {},
  semanticCoverage = {},
  sections = [],
} = {}) {
  const contract =
    coverageWithoutSections(contractCoverage);
  const semantic =
    coverageWithoutSections(semanticCoverage);
  const contractSupported =
    contractCoverageSupported(contractCoverage);

  const expectedMetricIds = uniqueStrings([
    ...(contract.expectedMetricIds || []),
    ...(semantic.expectedMetricIds || []),
  ]);

  const renderedMetricIds = uniqueStrings([
    ...(contract.renderedMetricIds || []),
    ...(semantic.renderedMetricIds || []),
    ...renderedMetricIdsFromSections(sections),
  ]);

  const inactiveMetricIds = uniqueStrings(
    contract.inactiveMetricIds,
  );
  const errorMetricIds = uniqueStrings(
    contract.errorMetricIds,
  );

  const primary = contractSupported
    ? contract
    : semantic;

  return {
    ...primary,
    version: String(primary.version || ""),
    contractCatalogVersion:
      String(
        primary.contractCatalogVersion || "",
      ),
    expectedMetricIds,
    renderedMetricIds,
    inactiveMetricIds,
    errorMetricIds,
    coverageSources: uniqueStrings([
      contractSupported
        ? "contract_driven_summary"
        : "",
      Object.keys(semantic).length
        ? "semantic_output_planner"
        : "",
    ]),
    contractCoverage:
      Object.keys(contract).length
        ? contract
        : undefined,
    semanticCoverage:
      Object.keys(semantic).length
        ? semantic
        : undefined,
    combinedCoverageVersion:
      "contract_semantic_coverage_union_v1",
  };
}

function normalizeIdentityText(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("ko-KR");
}

function canonicalizeHashValue(value) {
  if (value == null) return null;

  if (Array.isArray(value)) {
    return value.map(canonicalizeHashValue);
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      return String(value);
    }
    return Object.is(value, -0) ? 0 : value;
  }

  if (typeof value === "object") {
    return Object.keys(value)
      .sort((left, right) =>
        String(left).localeCompare(
          String(right),
          "ko",
        ),
      )
      .reduce((result, key) => {
        result[key] =
          canonicalizeHashValue(value[key]);
        return result;
      }, {});
  }

  return value;
}

function stableHashJson(value) {
  return JSON.stringify(
    canonicalizeHashValue(value),
  );
}

function resultRowsForContentHash(
  section = {},
) {
  const rows = section?.result?.rows;
  return Array.isArray(rows) ? rows : [];
}

function sectionContentPayload(
  section = {},
) {
  const result = section.result || {};

  return {
    version:
      BUSINESS_SECTION_CONTENT_HASH_VERSION,
    resultType: normalizeIdentityText(
      result.resultType || "",
    ),
    operation: normalizeIdentityText(
      result.operation ||
        result.recipeType ||
        "",
    ),
    metricHeader: normalizeIdentityText(
      result.metric?.header || "",
    ),
    groupHeader: normalizeIdentityText(
      result.groupBy?.header || "",
    ),
    rows: resultRowsForContentHash(
      section,
    ),
  };
}

function sectionContentHash(
  section = {},
) {
  return crypto
    .createHash("sha256")
    .update(
      stableHashJson(
        sectionContentPayload(section),
      ),
      "utf8",
    )
    .digest("hex");
}

function sectionTitleIdentity(
  section = {},
) {
  return normalizeIdentityText(
    section.title ||
      section.sectionId ||
      "",
  );
}

function applyMetricIdsToSection(
  section = {},
  metricIds = [],
) {
  const ids = uniqueStrings(metricIds);
  const result = section.result || {};

  return {
    ...section,
    metricIds: ids,
    result: {
      ...result,
      metricIds: ids,
      meta: {
        ...(result.meta || {}),
        metricIds: ids,
      },
    },
  };
}

function mergeEquivalentSections(
  primary = {},
  duplicate = {},
) {
  return applyMetricIdsToSection(
    primary,
    [
      ...collectSectionMetricIds(
        primary,
      ),
      ...collectSectionMetricIds(
        duplicate,
      ),
    ],
  );
}

function dedupeSectionsByContentHash(
  sections = [],
) {
  const kept = [];
  const seen = new Map();
  const removed = [];

  for (
    const section
    of Array.isArray(sections)
      ? sections
      : []
  ) {
    if (!section) continue;

    const titleKey =
      sectionTitleIdentity(section);
    const contentHash =
      sectionContentHash(section);
    const dedupeKey =
      `${titleKey}|${contentHash}`;

    if (!seen.has(dedupeKey)) {
      const index = kept.length;
      kept.push(section);
      seen.set(dedupeKey, index);
      continue;
    }

    const keptIndex =
      seen.get(dedupeKey);
    kept[keptIndex] =
      mergeEquivalentSections(
        kept[keptIndex],
        section,
      );

    removed.push({
      sectionId:
        String(
          section.sectionId || "",
        ),
      title:
        String(
          section.title || "",
        ),
      contentHash,
      retainedSectionId:
        String(
          kept[keptIndex]
            ?.sectionId || "",
        ),
    });
  }

  return {
    sections: kept,
    removed,
  };
}

const GENERIC_SECTION_OUTPUT_KEYS =
  new Set([
    "작업",
    "지표",
    "값",
    "합계",
    "평균",
    "행수",
    "건수",
    "유효값수",
    "비율",
    "비율percent",
    "순위",
    "metricid",
    "period",
    "count",
    "value",
    "sum",
    "average",
    "numericcount",
  ]);

function firstResultRowKey(
  section = {},
) {
  const firstRow =
    resultRowsForContentHash(
      section,
    )[0];

  if (
    !firstRow ||
    typeof firstRow !== "object" ||
    Array.isArray(firstRow)
  ) {
    return "";
  }

  return (
    Object.keys(firstRow).find(
      (key) =>
        !GENERIC_SECTION_OUTPUT_KEYS.has(
          normalizeIdentityText(key),
        ),
    ) || ""
  );
}

function sectionDimensionHeader(
  section = {},
) {
  const result = section.result || {};

  return String(
    result.groupBy?.header ||
      result.dimension?.header ||
      result.columns?.dimension ||
      section.candidate?.columns
        ?.dimension ||
      firstResultRowKey(section) ||
      "",
  ).trim();
}

function uniqueSectionTitle(
  requestedTitle = "",
  usedTitles = new Set(),
) {
  const base =
    String(
      requestedTitle || "분석 결과",
    ).trim() || "분석 결과";

  if (
    !usedTitles.has(
      normalizeIdentityText(base),
    )
  ) {
    return base;
  }

  for (
    let index = 2;
    index <= 200;
    index += 1
  ) {
    const candidate =
      `${base} (${index})`;
    const key =
      normalizeIdentityText(
        candidate,
      );

    if (!usedTitles.has(key)) {
      return candidate;
    }
  }

  return `${base} (${Date.now()})`;
}

function disambiguateDuplicateSectionTitles(
  sections = [],
) {
  const list = Array.isArray(sections)
    ? sections
    : [];
  const titleCounts = new Map();

  list.forEach((section) => {
    const key =
      sectionTitleIdentity(section);
    titleCounts.set(
      key,
      (titleCounts.get(key) || 0) + 1,
    );
  });

  const usedTitles = new Set();
  const renamed = [];

  const output = list.map(
    (section, index) => {
      const originalTitle =
        String(
          section.title ||
            section.sectionId ||
            `section_${index + 1}`,
        ).trim();
      const identity =
        normalizeIdentityText(
          originalTitle,
        );
      let nextTitle =
        originalTitle;

      if (
        usedTitles.has(identity) &&
        (titleCounts.get(identity) || 0) > 1
      ) {
        const dimension =
          sectionDimensionHeader(
            section,
          );
        const proposed =
          dimension
            ? `${originalTitle} (${dimension} 기준)`
            : `${originalTitle} (추가 분석)`;

        nextTitle =
          uniqueSectionTitle(
            proposed,
            usedTitles,
          );

        renamed.push({
          sectionId:
            String(
              section.sectionId || "",
            ),
          originalTitle,
          resolvedTitle: nextTitle,
          dimension,
        });
      } else {
        nextTitle =
          uniqueSectionTitle(
            originalTitle,
            usedTitles,
          );
      }

      usedTitles.add(
        normalizeIdentityText(
          nextTitle,
        ),
      );

      if (
        nextTitle === originalTitle
      ) {
        return section;
      }

      return {
        ...section,
        title: nextTitle,
        result: {
          ...(section.result || {}),
          title: nextTitle,
        },
        candidate: {
          ...(section.candidate || {}),
          title: nextTitle,
        },
      };
    },
  );

  return {
    sections: output,
    renamed,
  };
}

function normalizeInventorySnapshotOverviewLabels(
  sections = [],
) {
  let adjustedCount = 0;
  const adjustedSectionIds = [];

  const output = (
    Array.isArray(sections)
      ? sections
      : []
  ).map((section) => {
    const sectionType =
      normalizeIdentityText(
        section.sectionType ||
          section.result?.resultType ||
          "",
      );

    if (
      sectionType !==
      "inventory_flow_overview"
    ) {
      return section;
    }

    const rows =
      resultRowsForContentHash(
        section,
      );
    let changed = false;

    const nextRows = rows.map(
      (row) => {
        if (
          !row ||
          typeof row !== "object" ||
          Array.isArray(row) ||
          String(row.지표 || "").trim() !==
            "재고 수량 합계"
        ) {
          return row;
        }

        changed = true;
        adjustedCount += 1;

        return {
          ...row,
          지표:
            INVENTORY_SNAPSHOT_SUM_LABEL,
        };
      },
    );

    if (!changed) return section;

    adjustedSectionIds.push(
      String(
        section.sectionId || "",
      ),
    );

    return {
      ...section,
      result: {
        ...(section.result || {}),
        rows: nextRows,
        meta: {
          ...(
            section.result?.meta ||
            {}
          ),
          inventorySnapshotLabelVersion:
            BUSINESS_SECTION_FINALIZATION_VERSION,
          inventorySnapshotAggregation:
            "row_snapshot_sum",
          inventorySnapshotLabel:
            INVENTORY_SNAPSHOT_SUM_LABEL,
        },
      },
    };
  });

  return {
    sections: output,
    adjustedCount,
    adjustedSectionIds:
      uniqueStrings(
        adjustedSectionIds,
      ),
  };
}

function finalizeBusinessTemplateSections(
  sections = [],
) {
  const inventoryNormalized =
    normalizeInventorySnapshotOverviewLabels(
      sections,
    );
  const deduped =
    dedupeSectionsByContentHash(
      inventoryNormalized.sections,
    );
  const disambiguated =
    disambiguateDuplicateSectionTitles(
      deduped.sections,
    );

  return {
    sections:
      disambiguated.sections,
    meta: {
      businessSectionFinalizationVersion:
        BUSINESS_SECTION_FINALIZATION_VERSION,
      inventorySnapshotLabelAdjustedCount:
        inventoryNormalized
          .adjustedCount,
      inventorySnapshotLabelAdjustedSectionIds:
        inventoryNormalized
          .adjustedSectionIds,
      contentHashDedupeApplied:
        deduped.removed.length > 0,
      deduplicatedSectionCount:
        deduped.removed.length,
      deduplicatedSections:
        deduped.removed,
      duplicateTitleDisambiguationApplied:
        disambiguated
          .renamed.length > 0,
      disambiguatedSectionCount:
        disambiguated
          .renamed.length,
      disambiguatedSections:
        disambiguated.renamed,
    },
  };
}

function finalizeBusinessTemplateResult(
  result = {},
) {
  const finalized =
    finalizeBusinessTemplateSections(
      result.sections || [],
    );

  return {
    ...result,
    sections: finalized.sections,
    executionMeta: {
      ...(result.executionMeta || {}),
      ...finalized.meta,
      finalizedSectionCount:
        finalized.sections.length,
    },
  };
}

function attachContractCoverage({
  result = {},
  contractCoverage = {},
  contractObservation = {},
} = {}) {
  const finalized =
    finalizeBusinessTemplateResult(
      result,
    );
  const sections = Array.isArray(
    finalized.sections,
  )
    ? finalized.sections
    : [];

  return {
    ...finalized,
    contractSummaryCoverage:
      mergeContractAndSemanticCoverage({
        contractCoverage,
        semanticCoverage:
          finalized.contractSummaryCoverage || {},
        sections,
      }),
    executionMeta: {
      ...(finalized.executionMeta || {}),
      ...contractObservation,
      combinedExpectedMetricIdCount:
        uniqueStrings([
          ...(contractCoverage.expectedMetricIds || []),
          ...(
            finalized.contractSummaryCoverage
              ?.expectedMetricIds || []
          ),
        ]).length,
    },
  };
}

function augmentResult({
  result,
  normalizedQueryTables,
  templateCandidate,
  registryObservation = {},
  contractCoverage = {},
  contractObservation = {},
}) {
  const decision =
    semanticAugmentationDecision(templateCandidate);
  const observation = buildExecutorObservationMeta({
    templateCandidate,
    normalizedQueryTables,
    decision,
    registryObservation: {
      ...registryObservation,
      ...contractObservation,
    },
  });

  if (!decision.enabled) {
    return attachContractCoverage({
      result: {
        ...result,
        executionMeta: {
          ...(result.executionMeta || {}),
          ...observation,
          semanticBusinessAugmentation: false,
        },
      },
      contractCoverage,
      contractObservation,
    });
  }

  try {
    const augmented =
      augmentBusinessTemplateResult({
        executionResult: result,
        tables: normalizedQueryTables,
        templateCandidate,
        options: {
          maxDimensionsPerSeries: positiveInteger(
            process.env
              .SEMANTIC_PLANNER_BUSINESS_MAX_DIMENSIONS,
            8,
          ),
          maxAddedSections: positiveInteger(
            process.env
              .SEMANTIC_PLANNER_BUSINESS_MAX_ADDED_SECTIONS,
            64,
          ),
          maxPlannedSections: positiveInteger(
            process.env
              .SEMANTIC_PLANNER_BUSINESS_MAX_PLANNED_SECTIONS,
            120,
          ),
        },
        context: {
          augmentationVersion:
            BUSINESS_SEMANTIC_AUGMENTATION_VERSION,
          registryBridgeVersion:
            BUSINESS_TEMPLATE_REGISTRY_BRIDGE_VERSION,
          registryObservation,
        },
      });

    const semanticExpectedMetricIdCount =
      uniqueStrings(
        augmented.contractSummaryCoverage
          ?.expectedMetricIds,
      ).length;

    return attachContractCoverage({
      result: {
        ...augmented,
        executionMeta: {
          ...(augmented.executionMeta || {}),
          ...observation,
          semanticBusinessAugmentation: true,
          semanticAugmentationSkippedReason: "",
          semanticExpectedMetricIdCount,
        },
      },
      contractCoverage,
      contractObservation,
    });
  } catch (error) {
    console.warn(
      "[semantic-planner] business augmentation failed:",
      error?.message || error,
    );

    return attachContractCoverage({
      result: {
        ...result,
        executionMeta: {
          ...(result.executionMeta || {}),
          ...observation,
          semanticBusinessAugmentation: false,
          semanticAugmentationSkippedReason:
            "EXECUTION_ERROR",
          semanticAugmentationError:
            error?.message || String(error),
        },
      },
      contractCoverage,
      contractObservation,
    });
  }
}

function executeBusinessTemplate({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const templateId =
    String(templateCandidate.templateId || "").trim();

  if (!templateId) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_ID_REQUIRED",
      message: "templateId가 필요합니다.",
      executionMeta:
        buildExecutorObservationMeta({
          templateCandidate,
          normalizedQueryTables,
        }),
    };
  }

  const registered = executeRegisteredTemplate({
    normalizedQueryTables,
    templateCandidate,
  });

  if (!registered.ok) {
    return {
      ...(registered.errorResult || {
        ok: false,
        code: "BUSINESS_TEMPLATE_EXECUTION_FAILED",
        message:
          "업무 템플릿 실행에 실패했습니다.",
      }),
      executionMeta:
        buildExecutorObservationMeta({
          templateCandidate,
          normalizedQueryTables,
          registryObservation:
            registered.registryObservation || {},
        }),
    };
  }

  const registrySections =
    registered.sections || [];

  if (!registrySections.length) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_EXECUTION_EMPTY",
      message:
        "실행 가능한 템플릿 섹션이 없습니다.",
      executionMeta:
        buildExecutorObservationMeta({
          templateCandidate,
          normalizedQueryTables,
          registryObservation:
            registered.registryObservation || {},
        }),
    };
  }

  const contractExecution =
    executeContractCoverage({
      normalizedQueryTables,
      templateId,
    });
  const contractSections =
    contractExecution.sections || [];
  const sections = [
    ...registrySections,
    ...contractSections,
  ];

  const registryAndContractObservation = {
    ...(registered.registryObservation || {}),
    ...(contractExecution.observation || {}),
    registryBaseSectionCount:
      registrySections.length,
    contractBaseSectionCount:
      contractSections.length,
    preSemanticBaseSectionCount:
      sections.length,
  };

  const baseResult =
    registered.baseResult &&
    typeof registered.baseResult === "object"
      ? registered.baseResult
      : {};

  const result = {
    ...baseResult,
    ok: true,
    resultType:
      baseResult.resultType || "businessTemplate",
    templateId:
      baseResult.templateId || templateId,
    title:
      baseResult.title ||
      templateCandidate.title ||
      templateId,
    description:
      baseResult.description ||
      templateCandidate.description ||
      "",
    sections,
    contractSummaryCoverage:
      coverageWithoutSections(
        contractExecution.coverage,
      ),
    executionMeta: {
      ...(baseResult.executionMeta || {}),
      ...registryAndContractObservation,
    },
  };

  return augmentResult({
    result,
    normalizedQueryTables,
    templateCandidate,
    registryObservation:
      registered.registryObservation || {},
    contractCoverage:
      contractExecution.coverage || {},
    contractObservation:
      contractExecution.observation || {},
  });
}

module.exports = {
  BUSINESS_SEMANTIC_AUGMENTATION_VERSION,
  BUSINESS_TEMPLATE_EXECUTOR_VERSION,
  BUSINESS_TEMPLATE_REGISTRY_BRIDGE_VERSION,
  BUSINESS_TEMPLATE_CONTRACT_BRIDGE_VERSION,
  BUSINESS_SECTION_FINALIZATION_VERSION,
  BUSINESS_SECTION_CONTENT_HASH_VERSION,
  INVENTORY_SNAPSHOT_SUM_LABEL,
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  attachContractCoverage,
  buildContractCoverageObservation,
  buildExecutorObservationMeta,
  buildRegistryObservation,
  canonicalizeHashValue,
  dedupeSectionsByContentHash,
  disambiguateDuplicateSectionTitles,
  executeBusinessTemplate,
  executeContractCoverage,
  executeRegisteredTemplate,
  finalizeBusinessTemplateResult,
  finalizeBusinessTemplateSections,
  executeTemplateSections,
  mergeContractAndSemanticCoverage,
  normalizeInventorySnapshotOverviewLabels,
  normalizeRegistryExecutorOutput,
  sectionContentHash,
  sectionContentPayload,
  semanticAugmentationDecision,
  semanticAugmentationEnabled,
};
