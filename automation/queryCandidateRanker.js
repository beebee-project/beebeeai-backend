const { normalizeText, sha256 } = require("./queryCandidateObservation");

const QUERY_CANDIDATE_RANKING_RESOLUTION_VERSION =
  "query_candidate_ranking_resolution_v1";
const QUERY_CANDIDATE_RANKING_ITEM_VERSION = "query_candidate_ranking_item_v1";
const QUERY_CANDIDATE_RANKING_POLICY_VERSION =
  "deterministic_candidate_ranking_policy_v1_1";

const RANKING_DISPOSITION = Object.freeze(["RANKED", "NOT_APPLICABLE"]);
const RANKING_TIER = Object.freeze([
  "PRIMARY",
  "SECONDARY",
  "ADDITIONAL",
  "NOT_APPLICABLE",
]);
const RECOMMENDED_LIMIT = 8;
const PRIMARY_LIMIT = 3;

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const result = [];
  for (const value of asArray(values)) {
    const text = normalizeText(value);
    if (!text || seen.has(text)) continue;
    seen.add(text);
    result.push(text);
  }
  return result;
}

function sortedUnique(values = []) {
  return unique(values).sort((a, b) => a.localeCompare(b, "ko"));
}

function normalizeLoose(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[^가-힣a-z0-9]+/gu, "");
}

function stableClone(value) {
  if (Array.isArray(value)) return value.map(stableClone);
  if (!value || typeof value !== "object") return value;
  return Object.keys(value)
    .sort()
    .reduce((result, key) => {
      result[key] = stableClone(value[key]);
      return result;
    }, {});
}

function stableStringify(value) {
  return JSON.stringify(stableClone(value));
}

function clamp(value, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return min;
  return Math.min(max, Math.max(min, number));
}

function rounded(value) {
  return Number(Number(value || 0).toFixed(3));
}

function issue(path, code, message) {
  return { path, code, message };
}

function resolutionMap(candidateResolution = {}) {
  return new Map(
    asArray(candidateResolution.candidates).map((candidate) => [
      normalizeText(candidate.candidateId || ""),
      candidate,
    ]),
  );
}

function familyMemberMap(candidateFamilyResolution = {}) {
  return new Map(
    asArray(candidateFamilyResolution.candidates).map((candidate) => [
      normalizeText(candidate.candidateId || ""),
      candidate,
    ]),
  );
}

function familyMap(candidateFamilyResolution = {}) {
  return new Map(
    asArray(candidateFamilyResolution.families).map((family) => [
      normalizeText(family.familyId || ""),
      family,
    ]),
  );
}

function feasibilityMap(candidateFeasibilityResolution = {}) {
  return new Map(
    asArray(candidateFeasibilityResolution.candidates).map((candidate) => [
      normalizeText(candidate.candidateId || ""),
      candidate,
    ]),
  );
}

function semanticProfileColumnMap(deterministicSemanticProfile = {}) {
  const result = new Map();
  for (const table of asArray(deterministicSemanticProfile.tables)) {
    const rowCount = Number(table.shape?.rowCount || 0);
    for (const column of asArray(table.columns)) {
      const columnId = normalizeText(column.columnId || "");
      if (!columnId) continue;
      result.set(columnId, {
        columnId,
        tableId: normalizeText(column.tableId || table.tableId || ""),
        rowCount: Number.isFinite(rowCount) ? rowCount : 0,
        sourceHeader: normalizeText(column.sourceHeader || ""),
        normalizedHeader: normalizeText(column.normalizedHeader || ""),
        normalizedMeaning: normalizeText(column.normalizedMeaning || ""),
        semanticRole: normalizeText(column.semanticRole || ""),
        semanticType: normalizeText(column.semanticType || ""),
        metricFamily: normalizeText(column.metricFamily || ""),
        metricAliases: sortedUnique(column.metricAliases),
        roleAliases: sortedUnique(column.roleAliases),
        defaultAggregation: normalizeText(column.defaultAggregation || ""),
        unitSemantic: normalizeText(column.unitSemantic || ""),
        uniqueRatio: Number(column.stats?.uniqueRatio),
        nonEmptyRatio: Number(column.stats?.nonEmptyRatio),
      });
    }
  }
  return result;
}

function candidateEvidenceByColumnId(candidate = {}) {
  const result = new Map();
  const evidence = [
    ...asArray(candidate.evidence),
    ...asArray(candidate.checks?.requiredRoles).flatMap((role) =>
      asArray(role.matched),
    ),
    ...asArray(candidate.checks?.operandBinding?.operands).flatMap((operand) =>
      asArray(operand.matched),
    ),
  ];
  for (const item of evidence) {
    const columnId = normalizeText(item.columnId || "");
    if (!columnId || result.has(columnId)) continue;
    result.set(columnId, item);
  }
  return result;
}

function operationCategory(operation = "") {
  const value = normalizeLoose(operation);
  if (
    ["countrows", "categorycount", "timecount", "genericcountbygroup"].includes(
      value,
    )
  ) {
    return "OVERVIEW";
  }
  if (
    ["groupsum", "groupavg", "groupsummary", "compositionratio"].includes(value)
  ) {
    return "GROUP_AGGREGATION";
  }
  if (["timesum", "timeavg", "cumulativesum"].includes(value)) {
    return "TIME_SERIES";
  }
  if (["crosssum", "crosscount"].includes(value)) return "CROSS_TAB";
  if (["topbottom", "productsalesranking"].includes(value)) return "RANKING";
  if (["singlesourcedashboard", "multisourceschemaunion"].includes(value))
    return "STRUCTURAL";
  return "BUSINESS_TEMPLATE";
}

function operationUtility(operation = "", executorMode = "") {
  const value = normalizeLoose(operation);
  if (normalizeText(executorMode).toUpperCase() === "DECLARED") return 8;
  if (["countrows", "categorycount", "timecount"].includes(value)) return 8;
  if (
    ["groupsummary", "compositionratio", "crosssum", "crosscount"].includes(
      value,
    )
  )
    return 7;
  if (
    ["groupsum", "groupavg", "timesum", "timeavg", "cumulativesum"].includes(
      value,
    )
  )
    return 6;
  if (["topbottom"].includes(value)) return 5;
  return 4;
}

function priorRetrievalScore(candidate = {}) {
  const value = normalizeText(
    candidate.previousRetrievalResult || "",
  ).toUpperCase();
  if (value === "RETRIEVED") return 8;
  if (value === "DEFERRED") return 4;
  return 0;
}

function bindingScore(candidate = {}) {
  return (
    {
      BOUND: 10,
      PARTIAL: 7,
      INFERRED: 5,
      UNBOUND: 2,
    }[normalizeText(candidate.bindingStatus || "").toUpperCase()] || 0
  );
}

function executorScore(feasibility = {}) {
  const mode = normalizeText(
    feasibility.executionPlan?.executorMode || "",
  ).toUpperCase();
  if (mode === "DECLARED") return 10;
  if (mode === "GENERIC") return 7;
  return 0;
}

function identityScore(candidate = {}) {
  return normalizeText(candidate.templateId || "") ? 8 : 4;
}

function contractEvidenceScore(candidate = {}, feasibility = {}) {
  const operandStatus = normalizeText(
    candidate.checks?.operandBinding?.status || "",
  ).toUpperCase();
  const operandBindings = asArray(feasibility.executionPlan?.operandBindings);
  const roleBindings = asArray(feasibility.executionPlan?.requiredRoleBindings);
  if (operandStatus === "PASS" && operandBindings.length > 0) return 8;
  if (roleBindings.length > 0) return 8;
  const operation = normalizeLoose(feasibility.executionPlan?.operation || "");
  if (operation === "countrows") return 6;
  return 3;
}

function domainScore(candidate = {}) {
  const status = normalizeText(
    candidate.checks?.domainAlignment?.status || "",
  ).toUpperCase();
  if (status === "PASS") return 6;
  if (status === "NOT_APPLICABLE") return 3;
  return 0;
}

function metricScore(feasibility = {}) {
  const status = normalizeText(
    feasibility.checks?.metricContract?.status || "",
  ).toUpperCase();
  if (status === "PASS") return 4;
  if (status === "NOT_APPLICABLE") return 2;
  return 0;
}

function sourceSimplicityScore(feasibility = {}) {
  const count = sortedUnique(feasibility.executionPlan?.sourceTableIds).length;
  if (count === 1) return 4;
  if (count === 2) return 2;
  return 0;
}

function normalizedScoreContribution(value, maxPoints) {
  return rounded((clamp(value, 0, 100) / 100) * maxPoints);
}

function metricEvidenceText({
  candidate = {},
  measureColumnIds = [],
  semanticColumns = new Map(),
} = {}) {
  const candidateEvidence = candidateEvidenceByColumnId(candidate);
  const tokens = [];
  for (const columnId of sortedUnique(measureColumnIds)) {
    const profileColumn = semanticColumns.get(columnId) || {};
    const candidateColumn = candidateEvidence.get(columnId) || {};
    tokens.push(
      profileColumn.sourceHeader,
      profileColumn.normalizedHeader,
      profileColumn.normalizedMeaning,
      profileColumn.semanticRole,
      profileColumn.metricFamily,
      profileColumn.defaultAggregation,
      profileColumn.unitSemantic,
      ...asArray(profileColumn.metricAliases),
      ...asArray(profileColumn.roleAliases),
      candidateColumn.columnName,
      candidateColumn.semanticRole,
      candidateColumn.semanticType,
      candidateColumn.dataType,
    );
  }
  for (const operand of asArray(candidate.checks?.operandBinding?.operands)) {
    const kind = normalizeLoose(operand.kind || "");
    if (
      ![
        "measure",
        "amount",
        "revenue",
        "score",
        "rating",
        "satisfaction",
        "quantity",
      ].includes(kind)
    )
      continue;
    tokens.push(operand.expectedToken, operand.identifier);
  }
  for (const role of asArray(candidate.checks?.requiredRoles)) {
    const kind = normalizeLoose(role.role || "");
    if (
      ![
        "measure",
        "amount",
        "revenue",
        "score",
        "rating",
        "satisfaction",
        "quantity",
      ].includes(kind)
    )
      continue;
    tokens.push(role.role, ...asArray(role.aliases));
  }
  return normalizeLoose(tokens.filter(Boolean).join(" "));
}

function metricAggregationClass({
  candidate = {},
  measureColumnIds = [],
  semanticColumns = new Map(),
  source = {},
} = {}) {
  if (!measureColumnIds.length) return "NONE";
  const text = metricEvidenceText({
    candidate,
    measureColumnIds,
    semanticColumns,
  });
  const scaleTokens = [
    "만족도",
    "만족",
    "평가점수",
    "추천점수",
    "평점",
    "점수",
    "척도",
    "satisfaction",
    "rating",
    "score",
    "nps",
    "likert",
    "percent",
    "percentage",
    "비율",
    "율",
  ];
  if (scaleTokens.some((token) => text.includes(normalizeLoose(token))))
    return "SCALE";

  const columns = measureColumnIds.map(
    (columnId) => semanticColumns.get(columnId) || {},
  );
  const additiveByMetadata = columns.some(
    (column) =>
      normalizeLoose(column.defaultAggregation) === "sum" ||
      [
        "currency",
        "amount",
        "count",
        "quantity",
        "volume",
        "area",
        "weight",
      ].includes(normalizeLoose(column.unitSemantic)),
  );
  const additiveTokens = [
    "금액",
    "매출",
    "예산",
    "비용",
    "지출",
    "집행",
    "수량",
    "판매량",
    "건수",
    "인원",
    "면적",
    "중량",
    "총액",
    "합계",
    "amount",
    "revenue",
    "sales",
    "budget",
    "expense",
    "cost",
    "quantity",
    "volume",
    "count",
  ];
  const sourceText = normalizeLoose(
    `${source.primaryDomain || ""} ${source.datasetIntent || ""}`,
  );
  if (
    additiveByMetadata ||
    additiveTokens.some((token) => text.includes(normalizeLoose(token))) ||
    ["financ", "budget", "sales", "revenue", "expense", "cost"].some((token) =>
      sourceText.includes(token),
    )
  ) {
    return "ADDITIVE";
  }
  return "NEUTRAL";
}

function metricAggregationAffinity(operation = "", metricClass = "NONE") {
  const value = normalizeLoose(operation);
  if (metricClass === "SCALE") {
    if (["groupavg", "timeavg"].includes(value)) return 1.5;
    if (value === "topbottom") return 0.5;
    if (["groupsum", "timesum"].includes(value)) return -3;
    if (value === "cumulativesum") return -4;
  }
  if (metricClass === "ADDITIVE") {
    if (["groupsum", "timesum"].includes(value)) return 1;
    if (value === "cumulativesum") return 0.5;
    if (["groupavg", "timeavg"].includes(value)) return -1.5;
    if (value === "topbottom") return 1;
  }
  return 0;
}

function degenerateGroupPenalty({
  operation = "",
  axisColumnIds = [],
  semanticColumns = new Map(),
} = {}) {
  if (normalizeLoose(operation) !== "categorycount") return 0;
  let penalty = 0;
  for (const columnId of sortedUnique(axisColumnIds)) {
    const column = semanticColumns.get(columnId);
    if (!column || !Number.isFinite(column.uniqueRatio) || column.rowCount < 8)
      continue;
    if (column.uniqueRatio >= 0.98) penalty = Math.max(penalty, 12);
    else if (column.uniqueRatio >= 0.9) penalty = Math.max(penalty, 10);
    else if (column.uniqueRatio >= 0.8) penalty = Math.max(penalty, 7);
    else if (column.uniqueRatio >= 0.65) penalty = Math.max(penalty, 4);
  }
  return -penalty;
}

function scoreComponents(candidate = {}, feasibility = {}, context = {}) {
  const executorMode = normalizeText(
    feasibility.executionPlan?.executorMode || "",
  );
  const operation = normalizeText(feasibility.executionPlan?.operation || "");
  const signatures = bindingSignatures(feasibility);
  const metricClass = metricAggregationClass({
    candidate,
    measureColumnIds: signatures.measureColumnIds,
    semanticColumns: context.semanticColumns,
    source: context.source,
  });
  const components = {
    readyBaseline: 20,
    priorRetrieval: priorRetrievalScore(candidate),
    bindingEvidence: bindingScore(candidate),
    executorConfidence: executorScore(feasibility),
    identitySpecificity: identityScore(candidate),
    operationUtility: operationUtility(operation, executorMode),
    metricAggregationAffinity: metricAggregationAffinity(
      operation,
      metricClass,
    ),
    degenerateGroupPenalty: degenerateGroupPenalty({
      operation,
      axisColumnIds: signatures.axisColumnIds,
      semanticColumns: context.semanticColumns,
    }),
    contractEvidence: contractEvidenceScore(candidate, feasibility),
    domainAlignment: domainScore(candidate),
    metricAlignment: metricScore(feasibility),
    sourceSimplicity: sourceSimplicityScore(feasibility),
    resolutionEvidence: normalizedScoreContribution(
      candidate.resolutionScore,
      5,
    ),
    originalEvidence: normalizedScoreContribution(candidate.originalScore, 3),
  };
  const baseScore = rounded(
    Object.values(components).reduce(
      (sum, value) => sum + Number(value || 0),
      0,
    ),
  );
  return {
    components,
    baseScore: rounded(Math.min(100, Math.max(0, baseScore))),
    metricClass,
    signatures,
  };
}

function bindingSignatures(feasibility = {}) {
  const bindings = [
    ...asArray(feasibility.executionPlan?.operandBindings).map((binding) => ({
      kind: normalizeLoose(binding.kind || "unknown"),
      columnIds: sortedUnique(binding.columnIds),
    })),
    ...asArray(feasibility.executionPlan?.requiredRoleBindings).map(
      (binding) => ({
        kind: normalizeLoose(binding.role || "unknown"),
        columnIds: sortedUnique(binding.columnIds),
      }),
    ),
  ];
  const groupKinds = new Set([
    "group",
    "dimension",
    "category",
    "status",
    "product",
    "organization",
  ]);
  const periodKinds = new Set(["period", "date", "datetime"]);
  const measureKinds = new Set([
    "measure",
    "amount",
    "revenue",
    "score",
    "rating",
    "satisfaction",
    "quantity",
    "count",
  ]);
  const axis = sortedUnique(
    bindings
      .filter(
        (binding) =>
          groupKinds.has(binding.kind) || periodKinds.has(binding.kind),
      )
      .flatMap((binding) => binding.columnIds),
  );
  const measure = sortedUnique(
    bindings
      .filter((binding) => measureKinds.has(binding.kind))
      .flatMap((binding) => binding.columnIds),
  );
  return {
    axisColumnIds: axis,
    measureColumnIds: measure,
    axisSignature: axis.join("|"),
    measureSignature: measure.join("|"),
  };
}

function diversityPenalty(item = {}, ranked = []) {
  const sameOperation = ranked.filter(
    (previous) => previous.operation === item.operation,
  ).length;
  const sameAxis = item.axisSignature
    ? ranked.filter((previous) => previous.axisSignature === item.axisSignature)
        .length
    : 0;
  const sameMeasure = item.measureSignature
    ? ranked.filter(
        (previous) => previous.measureSignature === item.measureSignature,
      ).length
    : 0;
  const components = {
    repeatedOperation: Math.min(8, sameOperation * 2),
    repeatedAxis: Math.min(6, sameAxis * 2),
    repeatedMeasure: Math.min(4, sameMeasure),
  };
  return {
    components,
    total: rounded(
      Object.values(components).reduce((sum, value) => sum + value, 0),
    ),
  };
}

function compareEligible(a = {}, b = {}) {
  if (a.adjustedScore !== b.adjustedScore)
    return b.adjustedScore - a.adjustedScore;
  if (a.baseScore !== b.baseScore) return b.baseScore - a.baseScore;
  if (a.resolutionScore !== b.resolutionScore)
    return b.resolutionScore - a.resolutionScore;
  if (a.originalScore !== b.originalScore)
    return b.originalScore - a.originalScore;
  if (a.originalRank !== b.originalRank) return a.originalRank - b.originalRank;
  return a.candidateId.localeCompare(b.candidateId, "ko");
}

function tierForRank(rank) {
  if (rank <= PRIMARY_LIMIT) return "PRIMARY";
  if (rank <= RECOMMENDED_LIMIT) return "SECONDARY";
  return "ADDITIONAL";
}

function rankReadyCandidates(eligible = []) {
  const remaining = [...eligible];
  const ranked = [];
  while (remaining.length) {
    const scored = remaining.map((item) => {
      const penalty = diversityPenalty(item, ranked);
      return {
        ...item,
        diversityPenalty: penalty,
        adjustedScore: rounded(Math.max(0, item.baseScore - penalty.total)),
      };
    });
    scored.sort(compareEligible);
    const selected = scored[0];
    const index = remaining.findIndex(
      (item) => item.candidateId === selected.candidateId,
    );
    remaining.splice(index, 1);
    ranked.push({
      ...selected,
      rank: ranked.length + 1,
      tier: tierForRank(ranked.length + 1),
    });
  }
  return ranked;
}

function buildRankingItem({
  candidate = {},
  familyMember = {},
  family = {},
  feasibility = {},
  ranked = null,
}) {
  const candidateId = normalizeText(
    candidate.candidateId || feasibility.candidateId || "",
  );
  if (!ranked) {
    return {
      version: QUERY_CANDIDATE_RANKING_ITEM_VERSION,
      candidateId,
      familyId: normalizeText(
        familyMember.familyId || feasibility.familyId || "",
      ),
      feasibilityStatus: normalizeText(
        feasibility.feasibilityStatus || "NOT_APPLICABLE",
      ),
      rankingDisposition: "NOT_APPLICABLE",
      rank: null,
      rankingTier: "NOT_APPLICABLE",
      rankingScore: null,
      baseScore: null,
      diversityPenalty: null,
      scoreComponents: {},
      operation: normalizeText(feasibility.executionPlan?.operation || ""),
      operationCategory: operationCategory(
        feasibility.executionPlan?.operation || "",
      ),
      sourceTableIds: sortedUnique(feasibility.executionPlan?.sourceTableIds),
      recipeId: normalizeText(candidate.recipeId || ""),
      templateId: normalizeText(candidate.templateId || ""),
      reasonCode:
        normalizeText(feasibility.feasibilityStatus || "") === "REVIEW"
          ? "FEASIBILITY_REVIEW_NOT_RANKED"
          : normalizeText(feasibility.feasibilityStatus || "") === "UNSUPPORTED"
            ? "FEASIBILITY_UNSUPPORTED_NOT_RANKED"
            : "FEASIBILITY_NOT_READY",
      provenance: {
        feasibilityItemVersion: normalizeText(feasibility.version || ""),
        feasibilityItemSha256: normalizeText(
          feasibility.feasibilityItemSha256 || "",
        ),
        familyMemberVersion: normalizeText(familyMember.version || ""),
        familyMemberSha256: normalizeText(
          familyMember.familyMemberSha256 || "",
        ),
        resolutionItemVersion: normalizeText(candidate.version || ""),
        resolutionItemSha256: normalizeText(
          candidate.resolutionItemSha256 || "",
        ),
        sourceCandidateUnmodified: true,
        feasibilityStatusMutated: false,
        productionRouteChanged: false,
      },
    };
  }
  return {
    version: QUERY_CANDIDATE_RANKING_ITEM_VERSION,
    candidateId,
    familyId: normalizeText(
      familyMember.familyId || feasibility.familyId || "",
    ),
    feasibilityStatus: "READY",
    rankingDisposition: "RANKED",
    rank: ranked.rank,
    rankingTier: ranked.tier,
    rankingScore: ranked.adjustedScore,
    baseScore: ranked.baseScore,
    diversityPenalty: ranked.diversityPenalty,
    scoreComponents: ranked.scoreComponents,
    operation: ranked.operation,
    operationCategory: ranked.operationCategory,
    sourceTableIds: ranked.sourceTableIds,
    recipeId: normalizeText(candidate.recipeId || ""),
    templateId: normalizeText(candidate.templateId || ""),
    reasonCode: "DETERMINISTIC_READY_CANDIDATE_RANKED",
    provenance: {
      feasibilityItemVersion: normalizeText(feasibility.version || ""),
      feasibilityItemSha256: normalizeText(
        feasibility.feasibilityItemSha256 || "",
      ),
      familyMemberVersion: normalizeText(familyMember.version || ""),
      familyMemberSha256: normalizeText(familyMember.familyMemberSha256 || ""),
      resolutionItemVersion: normalizeText(candidate.version || ""),
      resolutionItemSha256: normalizeText(candidate.resolutionItemSha256 || ""),
      sourceCandidateUnmodified: true,
      feasibilityStatusMutated: false,
      productionRouteChanged: false,
    },
  };
}

function buildQueryCandidateRankingResolution({
  candidateResolution = {},
  candidateFamilyResolution = {},
  candidateFeasibilityResolution = {},
  deterministicSemanticProfile = {},
} = {}) {
  const resolutions = resolutionMap(candidateResolution);
  const familyMembers = familyMemberMap(candidateFamilyResolution);
  const families = familyMap(candidateFamilyResolution);
  const feasibilities = feasibilityMap(candidateFeasibilityResolution);
  const semanticColumns = semanticProfileColumnMap(
    deterministicSemanticProfile,
  );
  const rankingContext = {
    semanticColumns,
    source: {
      primaryDomain: normalizeText(
        candidateResolution.source?.primaryDomain || "",
      ),
      datasetIntent: normalizeText(
        candidateResolution.source?.datasetIntent || "",
      ),
    },
  };

  const eligible = [];
  for (const feasibility of asArray(
    candidateFeasibilityResolution.candidates,
  )) {
    if (
      normalizeText(feasibility.feasibilityStatus || "").toUpperCase() !==
      "READY"
    )
      continue;
    const candidateId = normalizeText(feasibility.candidateId || "");
    const candidate = resolutions.get(candidateId) || {};
    const familyMember = familyMembers.get(candidateId) || {};
    const family =
      families.get(
        normalizeText(familyMember.familyId || feasibility.familyId || ""),
      ) || {};
    const { components, baseScore, signatures } = scoreComponents(
      candidate,
      feasibility,
      rankingContext,
    );
    eligible.push({
      candidateId,
      candidate,
      familyMember,
      family,
      feasibility,
      baseScore,
      scoreComponents: components,
      operation: normalizeLoose(
        feasibility.executionPlan?.operation || "unknownoperation",
      ),
      operationCategory: operationCategory(
        feasibility.executionPlan?.operation || "",
      ),
      sourceTableIds: sortedUnique(feasibility.executionPlan?.sourceTableIds),
      axisSignature: signatures.axisSignature,
      measureSignature: signatures.measureSignature,
      resolutionScore: Number(candidate.resolutionScore || 0),
      originalScore: Number(candidate.originalScore || 0),
      originalRank: Number.isInteger(candidate.originalRank)
        ? candidate.originalRank
        : Number.MAX_SAFE_INTEGER,
    });
  }

  const ranked = rankReadyCandidates(eligible);
  const rankedById = new Map(ranked.map((item) => [item.candidateId, item]));
  const candidates = asArray(candidateResolution.candidates).map(
    (candidate) => {
      const candidateId = normalizeText(candidate.candidateId || "");
      const familyMember = familyMembers.get(candidateId) || {};
      const family =
        families.get(normalizeText(familyMember.familyId || "")) || {};
      const feasibility = feasibilities.get(candidateId) || {};
      return buildRankingItem({
        candidate,
        familyMember,
        family,
        feasibility,
        ranked: rankedById.get(candidateId) || null,
      });
    },
  );

  const rankedItems = candidates
    .filter((candidate) => candidate.rankingDisposition === "RANKED")
    .sort((a, b) => a.rank - b.rank);
  const counts = {
    total: candidates.length,
    readyInput: rankedItems.length,
    ranked: rankedItems.length,
    notApplicable: candidates.length - rankedItems.length,
    primary: rankedItems.filter(
      (candidate) => candidate.rankingTier === "PRIMARY",
    ).length,
    secondary: rankedItems.filter(
      (candidate) => candidate.rankingTier === "SECONDARY",
    ).length,
    additional: rankedItems.filter(
      (candidate) => candidate.rankingTier === "ADDITIONAL",
    ).length,
    recommended: Math.min(RECOMMENDED_LIMIT, rankedItems.length),
  };

  const document = {
    version: QUERY_CANDIDATE_RANKING_RESOLUTION_VERSION,
    itemVersion: QUERY_CANDIDATE_RANKING_ITEM_VERSION,
    policy: {
      version: QUERY_CANDIDATE_RANKING_POLICY_VERSION,
      readyCandidatesOnly: true,
      feasibilityRequired: true,
      reviewCandidatesExcluded: true,
      unsupportedCandidatesExcluded: true,
      allReadyCandidatesRanked: true,
      fixedScoreWeights: true,
      deterministicDiversityReranking: true,
      metricAggregationAffinityEnabled: true,
      degenerateGroupPenaltyEnabled: true,
      semanticProfileStatsAreSignalsNotGates: true,
      primaryLimit: PRIMARY_LIMIT,
      recommendedLimit: RECOMMENDED_LIMIT,
      priorScoresAreSignalsNotGates: true,
      sourceCandidatesAreNotRemovedOrMutated: true,
      feasibilityStatusMutation: false,
      productionRouteChanged: false,
    },
    source: {
      caseId: normalizeText(
        candidateResolution.source?.caseId || candidateResolution.caseId || "",
      ),
      fileName: normalizeText(candidateResolution.source?.fileName || ""),
      candidateResolutionVersion: normalizeText(
        candidateResolution.version || "",
      ),
      candidateResolutionPolicyVersion: normalizeText(
        candidateResolution.policy?.version || "",
      ),
      candidateResolutionSha256: normalizeText(
        candidateResolution.resolutionSha256 || "",
      ),
      candidateFamilyResolutionVersion: normalizeText(
        candidateFamilyResolution.version || "",
      ),
      candidateFamilyPolicyVersion: normalizeText(
        candidateFamilyResolution.policy?.version || "",
      ),
      candidateFamilyResolutionSha256: normalizeText(
        candidateFamilyResolution.familyResolutionSha256 || "",
      ),
      candidateFeasibilityResolutionVersion: normalizeText(
        candidateFeasibilityResolution.version || "",
      ),
      candidateFeasibilityPolicyVersion: normalizeText(
        candidateFeasibilityResolution.policy?.version || "",
      ),
      candidateFeasibilityResolutionSha256: normalizeText(
        candidateFeasibilityResolution.feasibilityResolutionSha256 || "",
      ),
      deterministicSemanticProfileVersion: normalizeText(
        deterministicSemanticProfile.version || "",
      ),
      deterministicSemanticProfileSha256: normalizeText(
        deterministicSemanticProfile.profileSha256 || "",
      ),
      primaryDomain: normalizeText(
        candidateResolution.source?.primaryDomain || "",
      ),
      datasetIntent: normalizeText(
        candidateResolution.source?.datasetIntent || "",
      ),
    },
    integrity: {
      sourceCandidateCount: asArray(candidateResolution.candidates).length,
      familyCandidateCount: asArray(candidateFamilyResolution.candidates)
        .length,
      feasibilityCandidateCount: asArray(
        candidateFeasibilityResolution.candidates,
      ).length,
      readyFeasibilityCount: asArray(
        candidateFeasibilityResolution.candidates,
      ).filter(
        (candidate) =>
          normalizeText(candidate.feasibilityStatus || "").toUpperCase() ===
          "READY",
      ).length,
      candidateCoverageComplete: false,
      readyCoverageComplete: false,
      rankSequenceContiguous: false,
      rankedCandidateIdsUnique: false,
      recommendedPrefixComplete: false,
      sourceCandidatesPreserved: true,
      rankingContainsNoSourceRecords: true,
    },
    counts,
    rankedCandidateIds: rankedItems.map((candidate) => candidate.candidateId),
    recommendedCandidateIds: rankedItems
      .slice(0, RECOMMENDED_LIMIT)
      .map((candidate) => candidate.candidateId),
    candidates,
  };

  const validation = validateQueryCandidateRankingResolution(document);
  document.integrity = {
    ...document.integrity,
    candidateCoverageComplete: validation.checks.candidateCoverageComplete,
    readyCoverageComplete: validation.checks.readyCoverageComplete,
    rankSequenceContiguous: validation.checks.rankSequenceContiguous,
    rankedCandidateIdsUnique: validation.checks.rankedCandidateIdsUnique,
    recommendedPrefixComplete: validation.checks.recommendedPrefixComplete,
  };
  document.rankingResolutionSha256 = sha256(
    stableStringify({
      ...document,
      rankingResolutionSha256: undefined,
    }),
  );
  return document;
}

function validateQueryCandidateRankingResolution(document = {}) {
  const errors = [];
  const warnings = [];
  const candidates = asArray(document.candidates);
  const ranked = candidates
    .filter((candidate) => candidate.rankingDisposition === "RANKED")
    .sort((a, b) => Number(a.rank) - Number(b.rank));
  const candidateIds = candidates.map((candidate) =>
    normalizeText(candidate.candidateId || ""),
  );
  const rankedIds = ranked.map((candidate) =>
    normalizeText(candidate.candidateId || ""),
  );
  const readyIds = candidates
    .filter(
      (candidate) =>
        normalizeText(candidate.feasibilityStatus || "").toUpperCase() ===
        "READY",
    )
    .map((candidate) => normalizeText(candidate.candidateId || ""));
  const expectedRanks = ranked.map((_, index) => index + 1);
  const actualRanks = ranked.map((candidate) => Number(candidate.rank));
  const recommended = asArray(document.recommendedCandidateIds).map(
    normalizeText,
  );
  const expectedRecommended = rankedIds.slice(
    0,
    Number(document.policy?.recommendedLimit || RECOMMENDED_LIMIT),
  );

  if (document.version !== QUERY_CANDIDATE_RANKING_RESOLUTION_VERSION) {
    errors.push(
      issue(
        "version",
        "VERSION_MISMATCH",
        "ranking resolution version이 올바르지 않습니다.",
      ),
    );
  }
  if (document.policy?.version !== QUERY_CANDIDATE_RANKING_POLICY_VERSION) {
    errors.push(
      issue(
        "policy.version",
        "POLICY_VERSION_MISMATCH",
        "ranking policy version이 올바르지 않습니다.",
      ),
    );
  }
  for (const [index, candidate] of candidates.entries()) {
    if (!RANKING_DISPOSITION.includes(candidate.rankingDisposition)) {
      errors.push(
        issue(
          `candidates[${index}].rankingDisposition`,
          "INVALID_DISPOSITION",
          "지원하지 않는 ranking disposition입니다.",
        ),
      );
    }
    if (!RANKING_TIER.includes(candidate.rankingTier)) {
      errors.push(
        issue(
          `candidates[${index}].rankingTier`,
          "INVALID_TIER",
          "지원하지 않는 ranking tier입니다.",
        ),
      );
    }
    if (candidate.rankingDisposition === "RANKED") {
      if (candidate.feasibilityStatus !== "READY") {
        errors.push(
          issue(
            `candidates[${index}]`,
            "NON_READY_CANDIDATE_RANKED",
            "READY가 아닌 후보가 순위화됐습니다.",
          ),
        );
      }
      if (!Number.isInteger(candidate.rank) || candidate.rank < 1) {
        errors.push(
          issue(
            `candidates[${index}].rank`,
            "INVALID_RANK",
            "rank는 1 이상의 정수여야 합니다.",
          ),
        );
      }
      if (
        !Number.isFinite(candidate.rankingScore) ||
        candidate.rankingScore < 0 ||
        candidate.rankingScore > 100
      ) {
        errors.push(
          issue(
            `candidates[${index}].rankingScore`,
            "INVALID_SCORE",
            "rankingScore는 0~100 범위여야 합니다.",
          ),
        );
      }
    } else if (candidate.rank != null || candidate.rankingScore != null) {
      errors.push(
        issue(
          `candidates[${index}]`,
          "NOT_APPLICABLE_HAS_RANK",
          "비대상 후보에는 rank와 score가 없어야 합니다.",
        ),
      );
    }
  }

  const checks = {
    candidateCoverageComplete:
      candidates.length === Number(document.counts?.total || 0) &&
      candidateIds.length === new Set(candidateIds).size,
    readyCoverageComplete:
      readyIds.length === rankedIds.length &&
      readyIds.every((candidateId) => rankedIds.includes(candidateId)),
    rankSequenceContiguous:
      JSON.stringify(actualRanks) === JSON.stringify(expectedRanks),
    rankedCandidateIdsUnique: rankedIds.length === new Set(rankedIds).size,
    recommendedPrefixComplete:
      JSON.stringify(recommended) === JSON.stringify(expectedRecommended),
  };
  for (const [key, passed] of Object.entries(checks)) {
    if (!passed)
      errors.push(
        issue(
          `integrity.${key}`,
          "INTEGRITY_CHECK_FAILED",
          `${key} 검사가 실패했습니다.`,
        ),
      );
  }
  if (Number(document.counts?.ranked || 0) !== ranked.length) {
    errors.push(
      issue(
        "counts.ranked",
        "COUNT_MISMATCH",
        "ranked count가 실제 항목 수와 다릅니다.",
      ),
    );
  }
  if (Number(document.counts?.readyInput || 0) !== readyIds.length) {
    errors.push(
      issue(
        "counts.readyInput",
        "COUNT_MISMATCH",
        "readyInput count가 실제 READY 수와 다릅니다.",
      ),
    );
  }
  if (
    Number(document.counts?.ranked || 0) +
      Number(document.counts?.notApplicable || 0) !==
    candidates.length
  ) {
    errors.push(
      issue(
        "counts",
        "STATUS_PARTITION_INCOMPLETE",
        "RANKED/NOT_APPLICABLE 분할이 전체 후보를 덮지 않습니다.",
      ),
    );
  }

  return {
    version: "query_candidate_ranking_validation_v1",
    valid: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings,
    checks,
  };
}

module.exports = {
  QUERY_CANDIDATE_RANKING_RESOLUTION_VERSION,
  QUERY_CANDIDATE_RANKING_ITEM_VERSION,
  QUERY_CANDIDATE_RANKING_POLICY_VERSION,
  RANKING_DISPOSITION,
  RANKING_TIER,
  RECOMMENDED_LIMIT,
  PRIMARY_LIMIT,
  buildQueryCandidateRankingResolution,
  validateQueryCandidateRankingResolution,
};
