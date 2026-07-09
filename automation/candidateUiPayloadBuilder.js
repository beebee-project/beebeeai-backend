const {
  buildSourceTablePolicy,
  getCanonicalTableId,
  getSourceTableId,
  getTableUsage,
  isVirtualQueryTable,
} = require("./sourceTablePolicy");
const {
  normalizeOutputTypes,
  getOutputTypeUiLabel,
} = require("./config/outputArtifactConfig");
const {
  getTemplateExposure,
  getTemplateExposureBadges,
  buildTemplateExposureSummary,
} = require("./config/templateExposureConfig");

const CANDIDATE_UI_PAYLOAD_VERSION = "candidate_ui_payload_v2";
const RECOMMENDED_CANDIDATE_LIMIT = 3;
const ANALYSIS_CANDIDATE_LIMIT = 60;
const GROUP_CANDIDATE_LIMIT = 30;

const CANDIDATE_TYPE_LABELS = Object.freeze({
  businessTemplate: "업무 템플릿",
  analysisRecipe: "분석 후보",
  dashboard: "대시보드 후보",
  automationCategory: "자동화 유형",
  multiSource: "다중 원본 후보",
});

const SOURCE_SCOPE_LABELS = Object.freeze({
  singleTable: "단일 원본",
  multiTable: "다중 원본",
  workbook: "전체 파일",
  virtualLinkedTable: "원본·정규화 연결",
});

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const result = [];

  for (const value of values) {
    const text = String(value ?? "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    result.push(text);
  }

  return result;
}

function cleanText(value = "", max = 300) {
  const text = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "";
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function clampNumber(value, fallback = 0, min = 0, max = 100) {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

function formatConfidencePercent(confidence) {
  const n = Number(confidence);
  if (!Number.isFinite(n) || n <= 0) return null;
  const pct = n <= 1 ? Math.round(n * 100) : Math.round(n);
  return Math.max(0, Math.min(100, pct));
}

function inferCandidateType(candidate = {}) {
  if (candidate.candidateType) return String(candidate.candidateType);
  if (candidate.templateId || candidate.type === "businessTemplate") {
    return "businessTemplate";
  }
  if (candidate.dashboardId || candidate.type === "dashboard")
    return "dashboard";
  if (candidate.categoryId && Array.isArray(candidate.candidates)) {
    return "automationCategory";
  }
  if (
    candidate.multiSourceCandidateVersion ||
    candidate.multiSourceCandidateKind ||
    candidate.type === "multiSource"
  ) {
    return "multiSource";
  }
  return "analysisRecipe";
}

function getCandidateId(candidate = {}, index = 0) {
  return (
    cleanText(
      candidate.candidateId ||
        candidate.id ||
        candidate.templateId ||
        candidate.dashboardId ||
        candidate.categoryId ||
        candidate.recipeId ||
        [candidate.recipeType, candidate.tableId, candidate.title]
          .filter(Boolean)
          .join(":"),
      180,
    ) || `ui_candidate_${index + 1}`
  );
}

function getSourceTableIds(candidate = {}) {
  return unique([
    ...asArray(candidate.sourceTableIds),
    candidate.sourceTableId,
    candidate.tableId,
    ...asArray(candidate.candidates).flatMap((item) => [
      item?.sourceTableId,
      item?.tableId,
      ...(Array.isArray(item?.sourceTableIds) ? item.sourceTableIds : []),
    ]),
  ]);
}

function getSourceSheetNames(candidate = {}) {
  return unique([
    ...asArray(candidate.sourceSheetNames),
    candidate.sourceSheetName,
    ...asArray(candidate.candidates).flatMap((item) => [
      item?.sourceSheetName,
      ...(Array.isArray(item?.sourceSheetNames) ? item.sourceSheetNames : []),
    ]),
  ]);
}

function getRecipeIds(candidate = {}) {
  return unique([
    ...asArray(candidate.recipeIds),
    candidate.recipeId,
    candidate.recipeType,
    candidate.type && candidate.type !== "businessTemplate"
      ? candidate.type
      : "",
    ...asArray(candidate.candidates).flatMap((item) => [
      item?.candidateId,
      item?.id,
      item?.recipeId,
      item?.recipeType,
    ]),
  ]);
}

function getCandidateSourceScope(candidate = {}, sourceTableIds = []) {
  const explicit = String(candidate.sourceScope || "").trim();
  if (explicit) return explicit;
  if (sourceTableIds.length > 1) return "multiTable";
  return "singleTable";
}

function getCandidateOutputTypes(candidate = {}) {
  const fallback =
    inferCandidateType(candidate) === "analysisRecipe"
      ? ["summarySheet"]
      : ["summarySheet", "analysisReport", "ppt"];
  return normalizeOutputTypes(candidate.outputTypes || fallback);
}

function getCandidateMatchedHeaders(candidate = {}) {
  return unique([
    ...asArray(candidate.matchedHeaders),
    candidate.groupHeader,
    candidate.metricHeader,
    candidate.dateHeader,
    candidate.statusHeader,
    ...Object.values(candidate.columns || {}),
  ]).slice(0, 8);
}

function getCandidateColumns(candidate = {}) {
  const columns = candidate.columns || {};
  return {
    dimension:
      columns.dimension ||
      candidate.groupHeader ||
      candidate.dimensionHeader ||
      "",
    metric:
      columns.metric ||
      columns.value ||
      columns.amount ||
      candidate.metricHeader ||
      candidate.valueHeader ||
      "",
    date: columns.date || columns.period || candidate.dateHeader || "",
    status: columns.status || candidate.statusHeader || "",
  };
}

function getTemplateDisplayMeta(candidate = {}) {
  if (!candidate.templateId) return null;
  const exposure = getTemplateExposure(candidate.templateId, candidate);

  return {
    domain: candidate.domain || exposure.domain || "",
    domainLabel: candidate.domainLabel || exposure.domainLabel || "",
    domainGroup: candidate.domainGroup || "",
    implementationLevel:
      candidate.implementationLevel || exposure.implementationLevel || "",
    implementationLevelLabel:
      candidate.implementationLevelLabel ||
      exposure.implementationLevelLabel ||
      "",
    exposureLevel: exposure.exposureLevel,
    exposureLabel: exposure.label,
    exposureShortLabel: exposure.shortLabel,
    exposureDescription: exposure.description,
    exposurePriority: exposure.priority,
    frontExposure: exposure,
    display: {
      title: cleanText(exposure.displayTitle || candidate.title || "", 120),
      subtitle: cleanText(
        exposure.displaySubtitle || candidate.description || "",
        180,
      ),
      order: Number.isFinite(Number(exposure.displayOrder))
        ? Number(exposure.displayOrder)
        : null,
    },
  };
}

function isFrontVisibleCard(card = {}) {
  return card.frontExposure?.hideFromFrontend !== true;
}

function buildRecommendationReason(candidate = {}, sourceSheetNames = []) {
  const explicit = cleanText(candidate.recommendationReason, 240);
  if (explicit) return explicit;

  const candidateType = inferCandidateType(candidate);
  const title = cleanText(
    candidate.title || candidate.templateId || candidate.recipeType,
    80,
  );
  const sourceLabel = sourceSheetNames.length
    ? `${sourceSheetNames.slice(0, 3).join(" · ")} 기준`
    : "업로드 데이터 기준";

  if (candidateType === "businessTemplate") {
    const matchedCount = Array.isArray(candidate.candidates)
      ? candidate.candidates.length
      : Number(candidate.matchedCount || 0);
    return `${sourceLabel}으로 ${matchedCount || "여러"}개 분석 후보가 매칭되어 ${title || "업무 템플릿"} 생성에 적합합니다.`;
  }

  if (candidateType === "multiSource") {
    return `${sourceLabel}의 여러 원본데이터 구조를 함께 비교하거나 통합할 수 있는 후보입니다.`;
  }

  if (candidateType === "dashboard") {
    return `${sourceLabel}으로 관련 분석 후보를 묶어 자동화 시트와 보고서 구성이 가능합니다.`;
  }

  return `${sourceLabel}으로 ${title || "분석 자동화"}를 생성할 수 있습니다.`;
}

function isExecutableCandidate(candidate = {}) {
  const type = inferCandidateType(candidate);
  if (type === "businessTemplate") return Boolean(candidate.templateId);
  if (type === "analysisRecipe")
    return Boolean(candidate.recipeType && candidate.tableId);
  return false;
}

function buildExecutionMeta(candidate = {}) {
  const candidateType = inferCandidateType(candidate);
  const exposure =
    candidateType === "businessTemplate"
      ? getTemplateExposure(candidate.templateId, candidate)
      : null;

  if (exposure?.hideFromFrontend) {
    return {
      executable: false,
      actionType: "hidden",
      endpoint: "",
      disabledReason:
        exposure.disabledReason ||
        "현재 버전에서는 안정성 확인 전까지 프론트 추천에서 제외됩니다.",
    };
  }

  const executable = isExecutableCandidate(candidate);
  const endpoint =
    candidateType === "businessTemplate"
      ? "/api/automation/execute-business-template"
      : candidateType === "analysisRecipe"
        ? "/api/automation/execute-analysis-candidate"
        : "";

  return {
    executable,
    actionType: executable
      ? candidateType === "businessTemplate"
        ? "executeBusinessTemplate"
        : "executeAnalysisCandidate"
      : "viewOnly",
    endpoint: executable ? endpoint : "",
    disabledReason: executable
      ? ""
      : "현재 버전에서는 이 후보를 직접 실행하지 않고 구조 확인용으로 표시합니다.",
  };
}

function candidateSortValue(candidate = {}) {
  const rank = Number(candidate.rank);
  if (Number.isFinite(rank) && rank > 0) return 100000 - rank;
  const rankScore = Number(candidate.rankScore);
  if (Number.isFinite(rankScore)) return rankScore;
  const priority = Number(candidate.priority);
  if (Number.isFinite(priority)) return priority;
  return 0;
}

function buildCandidateCard(candidate = {}, index = 0, context = {}) {
  const candidateType = inferCandidateType(candidate);
  const candidateId = getCandidateId(candidate, index);
  const sourceTableIds = getSourceTableIds(candidate);
  const sourceSheetNames = getSourceSheetNames(candidate);
  const sourceScope = getCandidateSourceScope(candidate, sourceTableIds);
  const outputTypes = getCandidateOutputTypes(candidate);
  const confidencePercent = formatConfidencePercent(candidate.confidence);
  const execution = buildExecutionMeta(candidate);
  const columns = getCandidateColumns(candidate);
  const templateDisplayMeta =
    candidateType === "businessTemplate"
      ? getTemplateDisplayMeta(candidate)
      : null;
  const templateExposureBadges = templateDisplayMeta
    ? getTemplateExposureBadges(candidate)
    : [];

  return {
    uiPayloadVersion: CANDIDATE_UI_PAYLOAD_VERSION,
    uiCandidateId: `${candidateType}:${candidateId}`,
    candidateId,
    candidateType,
    candidateTypeLabel: CANDIDATE_TYPE_LABELS[candidateType] || candidateType,
    templateId: candidate.templateId || "",
    domain: templateDisplayMeta?.domain || candidate.domain || "",
    domainLabel:
      templateDisplayMeta?.domainLabel || candidate.domainLabel || "",
    domainGroup:
      templateDisplayMeta?.domainGroup || candidate.domainGroup || "",
    implementationLevel:
      templateDisplayMeta?.implementationLevel ||
      candidate.implementationLevel ||
      "",
    implementationLevelLabel:
      templateDisplayMeta?.implementationLevelLabel ||
      candidate.implementationLevelLabel ||
      "",
    exposureLevel: templateDisplayMeta?.exposureLevel || "",
    exposureLabel: templateDisplayMeta?.exposureLabel || "",
    exposureShortLabel: templateDisplayMeta?.exposureShortLabel || "",
    exposurePriority: templateDisplayMeta?.exposurePriority ?? null,
    frontExposure: templateDisplayMeta?.frontExposure || null,
    display: templateDisplayMeta?.display || null,
    recipeType: candidate.recipeType || candidate.type || "",
    multiSourceCandidateKind: candidate.multiSourceCandidateKind || "",
    title: cleanText(
      templateDisplayMeta?.display?.title ||
        candidate.title ||
        candidate.name ||
        candidateId,
      120,
    ),
    description: cleanText(
      candidate.description ||
        candidate.summary ||
        templateDisplayMeta?.display?.subtitle ||
        "업로드된 파일 구조를 기반으로 추천된 자동화 후보입니다.",
      260,
    ),
    sourceScope,
    sourceScopeLabel: SOURCE_SCOPE_LABELS[sourceScope] || sourceScope,
    sourceTableIds,
    sourceSheetNames,
    sourceLabel: sourceSheetNames.length
      ? sourceSheetNames.join(" · ")
      : sourceScope === "workbook"
        ? "전체 파일"
        : "원본데이터",
    outputTypes,
    outputTypeLabels: outputTypes.map(getOutputTypeUiLabel).filter(Boolean),
    confidence: candidate.confidence ?? null,
    confidencePercent,
    priority: Number.isFinite(Number(candidate.priority))
      ? Number(candidate.priority)
      : null,
    rank: Number.isFinite(Number(candidate.rank))
      ? Number(candidate.rank)
      : null,
    rankingTier: candidate.rankingTier || "",
    rankScore: Number.isFinite(Number(candidate.rankScore))
      ? Number(candidate.rankScore)
      : null,
    recommendationReason: buildRecommendationReason(
      candidate,
      sourceSheetNames,
    ),
    reasonCodes: asArray(candidate.reasonCodes).slice(0, 12),
    matchedHeaders: getCandidateMatchedHeaders(candidate),
    columns,
    badges: unique([
      ...templateExposureBadges,
      CANDIDATE_TYPE_LABELS[candidateType] || candidateType,
      SOURCE_SCOPE_LABELS[sourceScope] || sourceScope,
      confidencePercent ? `신뢰도 ${confidencePercent}%` : "",
      candidate.aiAssisted ? "AI 보조" : "",
    ]),
    execution,
    score: candidate.score || null,
    ref: {
      candidateId,
      candidateType,
      templateId: candidate.templateId || "",
      recipeType: candidate.recipeType || "",
      tableId: candidate.tableId || candidate.sourceTableId || "",
      sourceTableId: candidate.sourceTableId || sourceTableIds[0] || "",
      sourceSheetName: candidate.sourceSheetName || sourceSheetNames[0] || "",
    },
    diagnostics: {
      source: context.source || "candidate-bundle",
      rawType: candidate.type || "",
    },
  };
}

function rankCards(cards = []) {
  return (Array.isArray(cards) ? cards : []).slice().sort((a, b) => {
    const ar = Number(a.rank || 0);
    const br = Number(b.rank || 0);
    if (ar > 0 || br > 0)
      return (ar || Number.MAX_SAFE_INTEGER) - (br || Number.MAX_SAFE_INTEGER);
    const scoreDiff = Number(b.rankScore || 0) - Number(a.rankScore || 0);
    if (Math.abs(scoreDiff) > 0.00001) return scoreDiff;
    return Number(b.priority || 0) - Number(a.priority || 0);
  });
}

function uniqueCardsById(cards = []) {
  const seen = new Set();
  const result = [];
  for (const card of cards) {
    const key = card.uiCandidateId || card.candidateId;
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(card);
  }
  return result;
}

function buildSourceTableItems(
  normalizedQueryTables = [],
  sourceTablePolicy = null,
) {
  const tables = Array.isArray(normalizedQueryTables)
    ? normalizedQueryTables
    : [];
  const policy =
    sourceTablePolicy ||
    buildSourceTablePolicy({ tables, normalizedQueryTables: tables });

  return tables.map((table, index) => {
    const tableId =
      getCanonicalTableId(table) ||
      table.tableId ||
      table.id ||
      `table_${index + 1}`;
    const sourceTableId = getSourceTableId(table) || tableId;
    const usage = getTableUsage(table);
    const sourceSheetName =
      table.sourceSheetName ||
      policy.sourceSheetByTableId?.[sourceTableId] ||
      policy.sourceSheetByTableId?.[tableId] ||
      table.sheetName ||
      `원본데이터${index + 1}`;
    const columns = Array.isArray(table.columns) ? table.columns : [];
    const isVirtual = isVirtualQueryTable(table);

    return {
      sourceTableId,
      tableId,
      sourceSheetName,
      originalSheetName: table.sheetName || "",
      tableName:
        table.tableName || table.tableTitle || table.sheetName || tableId,
      rowCount: Number(table.rowCount || table.rows?.length || 0),
      columnCount: columns.length,
      isPrimary:
        table.isPrimary === true ||
        policy.primarySourceTableId === sourceTableId ||
        policy.primarySourceTableId === tableId,
      isVirtual,
      sourceScope: isVirtual ? "virtualLinkedTable" : "singleTable",
      tableUsage: {
        queryable: usage.queryable !== false,
        analysisEligible: usage.analysisEligible !== false,
        templateEligible: usage.templateEligible !== false,
        reasons: asArray(usage.reasons).slice(0, 8),
      },
      columns: columns.slice(0, 12).map((column) => ({
        header:
          column.header ||
          column.originalHeader ||
          column.name ||
          column.key ||
          "",
        type: column.type || column.dominantType || "",
        role: column.role || column.inferredRole || "",
        uniqueCount: column.uniqueCount ?? column.profile?.uniqueCount ?? null,
      })),
    };
  });
}

function buildCandidateGroups({
  cardsByGroup = {},
  recommendedCards = [],
} = {}) {
  const groupDefs = [
    {
      groupId: "recommended",
      title: "추천 자동화 후보",
      description: "우선 실행하기 좋은 후보 3개입니다.",
      cards: recommendedCards,
    },
    {
      groupId: "businessTemplates",
      title: "업무 템플릿 후보",
      description: "매출·연구비·인사처럼 목적이 명확한 보고서 패키지입니다.",
      cards: cardsByGroup.businessTemplateCandidates || [],
    },
    {
      groupId: "analysisRecipes",
      title: "분석 후보 전체 보기",
      description:
        "단일 지표·기간·카테고리 기준으로 생성 가능한 분석 후보입니다.",
      cards: cardsByGroup.analysisRecipeCandidates || [],
    },
    {
      groupId: "dashboards",
      title: "대시보드 후보",
      description: "여러 분석 후보를 묶어 자동화 시트로 구성하는 후보입니다.",
      cards: cardsByGroup.dashboardCandidates || [],
    },
    {
      groupId: "multiSource",
      title: "다중 원본 후보",
      description:
        "원본데이터2 이상이 있거나 원본·정규화 표 연결이 필요한 후보입니다.",
      cards: cardsByGroup.multiSourceCandidates || [],
    },
  ];

  return groupDefs
    .filter(
      (group) =>
        group.groupId === "recommended" ||
        (Array.isArray(group.cards) && group.cards.length),
    )
    .map((group) => {
      const cards = Array.isArray(group.cards) ? group.cards : [];
      return {
        groupId: group.groupId,
        title: group.title,
        description: group.description,
        count: cards.length,
        emptyReason:
          cards.length || group.groupId !== "recommended"
            ? ""
            : "NO_RECOMMENDED_CANDIDATES",
        candidates: cards.slice(
          0,
          group.groupId === "analysisRecipes"
            ? ANALYSIS_CANDIDATE_LIMIT
            : GROUP_CANDIDATE_LIMIT,
        ),
      };
    });
}

function buildCandidateUiPayload({
  fileName = "",
  normalizedQueryTables = [],
  candidateBundle = {},
  source = "candidate-ui-payload",
} = {}) {
  const safeBundle =
    candidateBundle && typeof candidateBundle === "object"
      ? candidateBundle
      : {};
  const sourceTablePolicy = buildSourceTablePolicy({
    tables: normalizedQueryTables,
    normalizedQueryTables,
  });
  const sourceTables = buildSourceTableItems(
    normalizedQueryTables,
    sourceTablePolicy,
  );

  const candidateGroups = {
    businessTemplateCandidates: asArray(safeBundle.businessTemplateCandidates),
    dashboardCandidates: asArray(safeBundle.dashboardCandidates),
    multiSourceCandidates: asArray(safeBundle.multiSourceCandidates),
    analysisRecipeCandidates: asArray(safeBundle.analysisRecipeCandidates),
    topCandidates: asArray(safeBundle.topCandidates),
    secondaryCandidates: asArray(safeBundle.secondaryCandidates),
  };

  const cardsByGroup = Object.fromEntries(
    Object.entries(candidateGroups).map(([key, list]) => [
      key,
      uniqueCardsById(
        rankCards(
          list
            .map((candidate, index) =>
              buildCandidateCard(candidate, index, { source }),
            )
            .filter(isFrontVisibleCard),
        ),
      ),
    ]),
  );

  const topExecutable = cardsByGroup.topCandidates.filter(
    (card) => card.execution?.executable,
  );
  const fallbackExecutable = uniqueCardsById([
    ...cardsByGroup.businessTemplateCandidates,
    ...cardsByGroup.analysisRecipeCandidates,
  ]).filter((card) => card.execution?.executable);
  const recommendedCandidates = uniqueCardsById([
    ...topExecutable,
    ...fallbackExecutable,
  ]).slice(0, RECOMMENDED_CANDIDATE_LIMIT);

  const allCards = uniqueCardsById([
    ...recommendedCandidates,
    ...cardsByGroup.businessTemplateCandidates,
    ...cardsByGroup.dashboardCandidates,
    ...cardsByGroup.multiSourceCandidates,
    ...cardsByGroup.analysisRecipeCandidates,
    ...cardsByGroup.secondaryCandidates,
  ]);

  const templateExposureSummary = buildTemplateExposureSummary(
    allCards.filter((card) => card.candidateType === "businessTemplate"),
  );

  const scopeCounts = allCards.reduce((acc, card) => {
    const key = card.sourceScope || "unknown";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  return {
    version: CANDIDATE_UI_PAYLOAD_VERSION,
    generatedAt: new Date().toISOString(),
    fileName,
    source,
    summary: {
      recommendedCount: recommendedCandidates.length,
      allCandidateCount: allCards.length,
      sourceTableCount: sourceTables.filter((table) => !table.isVirtual).length,
      virtualTableCount: sourceTables.filter((table) => table.isVirtual).length,
      sourceScopeCounts: scopeCounts,
      businessTemplateExposureCounts:
        templateExposureSummary.exposureCounts || {},
      emptyReason: recommendedCandidates.length
        ? ""
        : "NO_RECOMMENDED_CANDIDATES",
    },
    recommendedCandidates,
    candidateGroups: buildCandidateGroups({
      cardsByGroup,
      recommendedCards: recommendedCandidates,
    }),
    sourceTables,
    sourceStructure: {
      title: "원본데이터 구조 확인",
      sourceSheetNames: unique(
        sourceTables.map((table) => table.sourceSheetName),
      ),
      primarySourceTableId: sourceTablePolicy.primarySourceTableId || "",
      primarySourceSheetName: sourceTablePolicy.primarySourceSheetName || "",
      scope:
        sourceTablePolicy.scope ||
        sourceTablePolicy.sourceScope ||
        "singleTable",
    },
    displayPolicy: {
      recommendedLimit: RECOMMENDED_CANDIDATE_LIMIT,
      showSourceStructure: true,
      showConfidence: true,
      showRecommendationReason: true,
      showOutputTypes: true,
      showTemplateExposure: true,
      templateExposure: templateExposureSummary,
    },
  };
}

module.exports = {
  CANDIDATE_UI_PAYLOAD_VERSION,
  RECOMMENDED_CANDIDATE_LIMIT,
  buildCandidateUiPayload,
  buildCandidateCard,
};
