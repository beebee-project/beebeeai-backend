const { normalizeText, sha256 } = require("./queryCandidateObservation");
const { assessCandidate } = require("./queryCandidateRetriever");
const {
  runConditionalCandidatePlanner,
  ALLOWED_OPERATIONS,
} = require("./queryCandidatePlanner");
const {
  buildQueryCandidateResolution,
  validateQueryCandidateResolution,
} = require("./queryCandidateResolver");
const {
  buildQueryCandidateFamilyResolution,
  validateQueryCandidateFamilyResolution,
} = require("./queryCandidateFamilyResolver");
const {
  buildQueryCandidateFeasibilityResolution,
  validateQueryCandidateFeasibilityResolution,
} = require("./queryCandidateFeasibilityGate");
const {
  buildQueryCandidateRankingResolution,
  validateQueryCandidateRankingResolution,
} = require("./queryCandidateRanker");

const QUERY_CANDIDATE_PLANNER_SHADOW_RESOLUTION_VERSION =
  "query_candidate_planner_shadow_resolution_v1";
const QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION =
  "query_candidate_planner_reentry_bundle_v1";
const QUERY_CANDIDATE_PLANNER_SHADOW_POLICY_VERSION =
  "conditional_llm_candidate_planner_shadow_policy_v1";
const QUERY_CANDIDATE_PLANNER_SHADOW_MODE = "LIVE_SHADOW";

const SHADOW_STATUS = Object.freeze([
  "SKIPPED",
  "REQUIRED_NOT_RUN",
  "FAILED_SAFE",
  "SHADOW_COMPLETED",
]);

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
  return unique(values).sort((left, right) => left.localeCompare(right, "ko"));
}

function normalizeLoose(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[^가-힣a-z0-9]+/gu, "");
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value == null ? {} : value));
}

function columnMap(profile = {}) {
  const result = new Map();
  for (const table of asArray(profile.tables)) {
    for (const column of asArray(table.columns)) {
      const columnId = normalizeText(column.columnId || "");
      if (!columnId) continue;
      result.set(columnId, {
        ...column,
        tableId: normalizeText(table.tableId || column.tableId || ""),
      });
    }
  }
  return result;
}

function safeOperandToken(column = {}, fallback = "operand") {
  const value = normalizeText(
    column.normalizedHeader ||
      column.sourceHeader ||
      column.normalizedMeaning ||
      fallback,
  )
    .normalize("NFKC")
    .replace(/[_\s]+/gu, "-")
    .replace(/[^가-힣A-Za-z0-9-]+/gu, "")
    .replace(/-+/gu, "-")
    .replace(/^-|-$/gu, "");
  return value || fallback;
}

function canonicalOperation(operation = "") {
  const normalized = normalizeLoose(operation);
  const definition = ALLOWED_OPERATIONS[normalized];
  return definition?.operation || normalizeText(operation || "");
}

function orderedBindings(proposal = {}) {
  const definition =
    ALLOWED_OPERATIONS[normalizeLoose(proposal.operation || "")];
  const remaining = asArray(proposal.operandBindings).map((binding, index) => ({
    index,
    kind: normalizeLoose(binding.kind || ""),
    columnId: normalizeText(binding.columnId || ""),
  }));
  const ordered = [];
  for (const kind of asArray(definition?.kinds)) {
    const matchIndex = remaining.findIndex((binding) => binding.kind === kind);
    if (matchIndex < 0) continue;
    ordered.push(remaining.splice(matchIndex, 1)[0]);
  }
  return [...ordered, ...remaining].map(({ index, ...binding }) => binding);
}

function recipeIdForProposal(proposal = {}, deterministicSemanticProfile = {}) {
  const operation = canonicalOperation(proposal.operation);
  const columns = columnMap(deterministicSemanticProfile);
  const tokens = orderedBindings(proposal).map((binding, index) =>
    safeOperandToken(
      columns.get(binding.columnId),
      `${binding.kind || "operand"}${index + 1}`,
    ),
  );
  return [operation, ...tokens].filter(Boolean).join("_");
}

function roleRequirement(binding = {}, column = {}, index = 0) {
  const kind = normalizeLoose(binding.kind || "operand") || "operand";
  const header = normalizeText(
    column.normalizedHeader ||
      column.sourceHeader ||
      column.normalizedMeaning ||
      "",
  );
  const semanticType = normalizeText(column.semanticType || "unknown");
  const dataType = normalizeText(column.dataType || "unknown");
  return {
    role: `${kind}${index + 1}`,
    aliases: unique([header || kind]),
    dataType,
    semanticType,
    required: true,
    source: "conditional_llm_planner_shadow",
  };
}

function capabilitiesForOperation(operation = "") {
  switch (normalizeLoose(operation)) {
    case "countrows":
      return ["operation:countRows", "single_table"];
    case "categorycount":
      return ["group_by", "operation:countRows", "single_table"];
    case "groupsum":
      return [
        "group_by",
        "operation:sum",
        "metric_kind:aggregate",
        "single_table",
      ];
    case "groupavg":
      return [
        "group_by",
        "operation:average",
        "metric_kind:aggregate",
        "single_table",
      ];
    case "topbottom":
      return [
        "group_by",
        "operation:rank",
        "metric_kind:rank",
        "ranking",
        "single_table",
      ];
    case "timesum":
      return [
        "operation:sum",
        "operation:timeSeries",
        "metric_kind:aggregate",
        "single_table",
      ];
    case "timeavg":
      return [
        "operation:average",
        "operation:timeSeries",
        "metric_kind:aggregate",
        "single_table",
      ];
    case "timecount":
      return ["operation:countRows", "operation:timeSeries", "single_table"];
    case "cumulativesum":
      return [
        "operation:sum",
        "operation:timeSeries",
        "metric_kind:aggregate",
        "single_table",
      ];
    case "crosssum":
      return [
        "group_by",
        "operation:sum",
        "metric_kind:aggregate",
        "single_table",
      ];
    case "crosscount":
      return ["group_by", "operation:countRows", "single_table"];
    default:
      return ["single_table"];
  }
}

function capabilityItemForProposal(
  proposal = {},
  deterministicSemanticProfile = {},
) {
  const columns = columnMap(deterministicSemanticProfile);
  const bindings = orderedBindings(proposal);
  const operation = canonicalOperation(proposal.operation);
  const recipeId = recipeIdForProposal(proposal, deterministicSemanticProfile);
  const item = {
    version: "query_candidate_capability_item_v1",
    candidateId: normalizeText(proposal.candidateId || ""),
    recipeId,
    recipeIds: unique([operation, recipeId]),
    templateId: "",
    candidateType: "ANALYSIS_RECIPE",
    bindingStatus: "INFERRED",
    bindingSource: "CONDITIONAL_LLM_PLANNER_SHADOW",
    bindingKey: recipeId,
    contractIds: [],
    matchedTemplateIds: [],
    requiredColumnRoles: bindings.map((binding, index) =>
      roleRequirement(binding, columns.get(binding.columnId) || {}, index),
    ),
    optionalColumnRoles: [],
    metricContracts: [],
    coreMetricIds: [],
    conditionalMetricIds: [],
    supportedMetricIds: [],
    metricFamilies: [],
    supportedOperations: [operation],
    requiredCapabilities: capabilitiesForOperation(operation),
    executorSupport: {
      status: "GENERIC",
      outputTypes: ["summarySheet"],
      reasons: ["conditional_llm_planner_shadow_reentry"],
    },
    constraints: {
      minimumTableCount: 1,
      maximumTableCount: 1,
      sourceScope: "singleTable",
      minimumRowCount: 1,
    },
    provenance: {
      candidateContractVersion: "query_candidate_item_v1",
      candidateStatus: "UNASSESSED",
      observedClass: "CONDITIONAL",
      plannerItemSha256: normalizeText(proposal.plannerItemSha256 || ""),
      shadowOnly: true,
    },
  };
  item.capabilitySha256 = sha256({ ...item, capabilitySha256: undefined });
  return item;
}

function retrievalItemForProposal(
  proposal = {},
  capability = {},
  resolvedSemanticProfile = {},
) {
  const candidate = {
    version: "query_candidate_item_v1",
    candidateId: normalizeText(proposal.candidateId || ""),
    recipeId: normalizeText(capability.recipeId || ""),
    templateId: "",
    candidateType: "ANALYSIS_RECIPE",
    observedClass: "CONDITIONAL",
    visibility: "HIDDEN",
    rank: null,
    score: Number(proposal.confidence || 0) * 100,
    sourceTableIds: sortedUnique(proposal.sourceTableIds),
    status: "UNASSESSED",
  };
  const assessment = assessCandidate(
    candidate,
    capability,
    resolvedSemanticProfile,
  );
  assessment.provenance = {
    ...(assessment.provenance || {}),
    candidateItemVersion: "query_candidate_item_v1",
    candidateStatus: "UNASSESSED",
    plannerItemSha256: normalizeText(proposal.plannerItemSha256 || ""),
    plannerReentry: true,
    shadowOnly: true,
  };
  assessment.retrievalItemSha256 = sha256({
    ...assessment,
    retrievalItemSha256: undefined,
  });
  return assessment;
}

function buildCapabilityManifest(candidates = [], source = {}) {
  const manifest = {
    version: "query_candidate_capability_manifest_v1",
    itemVersion: "query_candidate_capability_item_v1",
    source: {
      caseId: normalizeText(source.caseId || ""),
      fileName: "",
      contractVersion: "query_candidate_contract_v1",
      contractSha256: normalizeText(source.plannerResolutionSha256 || ""),
      shadowOnly: true,
    },
    counts: {
      total: candidates.length,
      bound: 0,
      partial: 0,
      inferred: candidates.length,
      unbound: 0,
      executorDeclared: 0,
      executorGeneric: candidates.length,
      executorUnknown: 0,
    },
    candidates,
  };
  manifest.manifestSha256 = sha256({ ...manifest, manifestSha256: undefined });
  return manifest;
}

function buildRetrievalDocument(candidates = [], source = {}) {
  const document = {
    version: "query_candidate_retrieval_v1",
    itemVersion: "query_candidate_retrieval_item_v1",
    policy: {
      version: "deterministic_candidate_retrieval_policy_v1",
      explicitBindingStatuses: ["BOUND", "PARTIAL"],
      inferredCandidatesAreDeferred: true,
      unboundCandidatesAreDeferred: true,
      onlyExplicitMissingRequirementsAreExcluded: true,
      candidateStatusMutation: false,
      shadowOnly: true,
    },
    source: {
      caseId: normalizeText(source.caseId || ""),
      fileName: "",
      contractVersion: "query_candidate_contract_v1",
      contractSha256: normalizeText(source.plannerResolutionSha256 || ""),
      capabilityManifestVersion: "query_candidate_capability_manifest_v1",
      capabilityManifestSha256: normalizeText(
        source.capabilityManifestSha256 || "",
      ),
      semanticProfileVersion: normalizeText(
        source.semanticProfileVersion || "",
      ),
      semanticProfileSha256: normalizeText(source.semanticProfileSha256 || ""),
      shadowOnly: true,
    },
    integrity: {
      contractCandidateCount: candidates.length,
      capabilityCandidateCount: candidates.length,
      missingCapabilityCandidateIds: [],
      orphanCapabilityCandidateIds: [],
      candidateCountMatch: true,
    },
    counts: {
      total: candidates.length,
      retrieved: candidates.filter((item) => item.result === "RETRIEVED")
        .length,
      deferred: candidates.filter((item) => item.result === "DEFERRED").length,
      excluded: candidates.filter((item) => item.result === "EXCLUDED").length,
      boundRetrieved: 0,
      partialRetrieved: 0,
      inferredDeferred: candidates.filter(
        (item) =>
          item.result === "DEFERRED" && item.bindingStatus === "INFERRED",
      ).length,
      unboundDeferred: 0,
    },
    candidates,
  };
  document.retrievalSha256 = sha256({
    ...document,
    retrievalSha256: undefined,
  });
  return document;
}

function normalizedResolvedProfile(
  deterministicSemanticProfile = {},
  resolvedSemanticProfile = {},
) {
  const profile = deepClone(
    Object.keys(resolvedSemanticProfile || {}).length
      ? resolvedSemanticProfile
      : deterministicSemanticProfile,
  );
  const deterministicHash = normalizeText(
    deterministicSemanticProfile.profileSha256 ||
      deterministicSemanticProfile.semanticProfileSha256 ||
      sha256(deterministicSemanticProfile),
  );
  profile.source = {
    ...(profile.source || {}),
    deterministicProfileSha256:
      normalizeText(profile.source?.deterministicProfileSha256 || "") ||
      deterministicHash,
  };
  if (!profile.profileSha256) {
    profile.profileSha256 = sha256({ ...profile, profileSha256: undefined });
  }
  return profile;
}

function buildPlannerReentryBundle({
  plannerResolution = {},
  deterministicSemanticProfile = {},
  resolvedSemanticProfile = {},
} = {}) {
  const accepted = asArray(plannerResolution.proposals).filter(
    (proposal) => proposal.disposition === "ACCEPTED_FOR_REVALIDATION",
  );
  const resolvedProfile = normalizedResolvedProfile(
    deterministicSemanticProfile,
    resolvedSemanticProfile,
  );
  const capabilities = accepted.map((proposal) =>
    capabilityItemForProposal(proposal, deterministicSemanticProfile),
  );
  const capabilityManifest = buildCapabilityManifest(capabilities, {
    caseId:
      plannerResolution.source?.caseId ||
      deterministicSemanticProfile.source?.caseId ||
      resolvedProfile.source?.caseId,
    plannerResolutionSha256: plannerResolution.plannerResolutionSha256,
  });
  const capabilityById = new Map(
    capabilities.map((item) => [item.candidateId, item]),
  );
  const retrievalCandidates = accepted.map((proposal) =>
    retrievalItemForProposal(
      proposal,
      capabilityById.get(proposal.candidateId) || {},
      resolvedProfile,
    ),
  );
  const deterministicHash = normalizeText(
    resolvedProfile.source?.deterministicProfileSha256 || "",
  );
  const retrieval = buildRetrievalDocument(retrievalCandidates, {
    caseId:
      plannerResolution.source?.caseId ||
      deterministicSemanticProfile.source?.caseId ||
      resolvedProfile.source?.caseId,
    plannerResolutionSha256: plannerResolution.plannerResolutionSha256,
    capabilityManifestSha256: capabilityManifest.manifestSha256,
    semanticProfileVersion: normalizeText(
      deterministicSemanticProfile.version || "",
    ),
    semanticProfileSha256: deterministicHash,
  });
  return {
    version: QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION,
    shadowOnly: true,
    acceptedProposalCount: accepted.length,
    acceptedCandidateIds: accepted.map((item) => item.candidateId),
    resolvedSemanticProfile: resolvedProfile,
    capabilityManifest,
    retrieval,
    bundleSha256: sha256({
      version: QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION,
      acceptedCandidateIds: accepted.map((item) => item.candidateId),
      capabilityManifestSha256: capabilityManifest.manifestSha256,
      retrievalSha256: retrieval.retrievalSha256,
      resolvedSemanticProfileSha256: resolvedProfile.profileSha256,
    }),
  };
}

function runPlannerResolverReentry({
  plannerResolution = {},
  deterministicSemanticProfile = {},
  resolvedSemanticProfile = {},
} = {}) {
  const bundle = buildPlannerReentryBundle({
    plannerResolution,
    deterministicSemanticProfile,
    resolvedSemanticProfile,
  });
  const candidateResolution = buildQueryCandidateResolution({
    retrieval: bundle.retrieval,
    capabilityManifest: bundle.capabilityManifest,
    resolvedSemanticProfile: bundle.resolvedSemanticProfile,
  });
  const candidateFamilyResolution = buildQueryCandidateFamilyResolution({
    candidateResolution,
  });
  const candidateFeasibilityResolution =
    buildQueryCandidateFeasibilityResolution({
      candidateFamilyResolution,
      candidateResolution,
    });
  const candidateRankingResolution = buildQueryCandidateRankingResolution({
    candidateResolution,
    candidateFamilyResolution,
    candidateFeasibilityResolution,
    deterministicSemanticProfile,
  });
  const validations = {
    resolver: validateQueryCandidateResolution(candidateResolution),
    family: validateQueryCandidateFamilyResolution(candidateFamilyResolution),
    feasibility: validateQueryCandidateFeasibilityResolution(
      candidateFeasibilityResolution,
    ),
    ranking: validateQueryCandidateRankingResolution(
      candidateRankingResolution,
    ),
  };
  const candidateById = new Map(
    asArray(candidateResolution.candidates).map((item) => [
      item.candidateId,
      item,
    ]),
  );
  const feasibilityById = new Map(
    asArray(candidateFeasibilityResolution.candidates).map((item) => [
      item.candidateId,
      item,
    ]),
  );
  const rankingById = new Map(
    asArray(candidateRankingResolution.candidates).map((item) => [
      item.candidateId,
      item,
    ]),
  );
  const items = bundle.acceptedCandidateIds.map((candidateId) => {
    const resolver = candidateById.get(candidateId) || {};
    const feasibility = feasibilityById.get(candidateId) || {};
    const ranking = rankingById.get(candidateId) || {};
    return {
      candidateId,
      resolverResult: normalizeText(resolver.result || ""),
      feasibilityStatus: normalizeText(feasibility.feasibilityStatus || ""),
      rankingDisposition: normalizeText(ranking.rankingDisposition || ""),
      shadowRank: Number.isInteger(ranking.rank) ? ranking.rank : null,
      productionCandidateMerged: false,
      productionReadyAssigned: false,
    };
  });
  return {
    bundle,
    candidateResolution,
    candidateFamilyResolution,
    candidateFeasibilityResolution,
    candidateRankingResolution,
    validations,
    items,
  };
}

function baseShadowDocument({
  plannerResolution = {},
  mode = QUERY_CANDIDATE_PLANNER_SHADOW_MODE,
} = {}) {
  return {
    version: QUERY_CANDIDATE_PLANNER_SHADOW_RESOLUTION_VERSION,
    policy: {
      version: QUERY_CANDIDATE_PLANNER_SHADOW_POLICY_VERSION,
      mode,
      conditionalPlannerRunsFirst: true,
      acceptedProposalsReenterResolver: true,
      resolverFamilyFeasibilityRankerShadowChain: true,
      shadowOnly: true,
      sourceCandidatesAreNotRemovedOrMutated: true,
      sourceCandidateStatusMutation: false,
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      plaintextPersistenceAllowed: false,
    },
    source: {
      caseId: normalizeText(plannerResolution.source?.caseId || ""),
      plannerResolutionVersion: normalizeText(plannerResolution.version || ""),
      plannerResolutionSha256: normalizeText(
        plannerResolution.plannerResolutionSha256 || "",
      ),
    },
    privacy: {
      rawRowsSent: false,
      sampleValuesSent: false,
      originalFileSent: false,
      fileNameSent: false,
    },
  };
}

function finalizeShadow(document) {
  document.shadowResolutionSha256 = sha256({
    ...document,
    shadowResolutionSha256: undefined,
  });
  return document;
}

async function runCandidatePlannerLiveShadow({
  semanticProfile,
  resolvedSemanticProfile,
  candidateResolution,
  candidateFeasibilityResolution,
  candidateRankingResolution,
  provider,
  cache,
  tenantId,
  cacheSecret,
  model,
  reasoningEffort,
  pricing,
  mode = QUERY_CANDIDATE_PLANNER_SHADOW_MODE,
} = {}) {
  const plannerResolution = await runConditionalCandidatePlanner({
    semanticProfile,
    resolvedSemanticProfile,
    candidateResolution,
    candidateFeasibilityResolution,
    candidateRankingResolution,
    provider,
    cache,
    tenantId,
    cacheSecret,
    model,
    reasoningEffort,
    pricing,
  });
  const base = baseShadowDocument({ plannerResolution, mode });
  const acceptedCount = Number(plannerResolution.counts?.accepted || 0);
  if (plannerResolution.invocation?.status === "REQUIRED_NOT_RUN") {
    return finalizeShadow({
      ...base,
      status: "REQUIRED_NOT_RUN",
      plannerResolution,
      reentry: null,
      counts: { accepted: 0, resolved: 0, ready: 0, ranked: 0 },
      integrity: {
        sourceCandidatesPreserved: true,
        productionCandidateMerge: false,
        productionRouteChanged: false,
      },
    });
  }
  if (plannerResolution.invocation?.status === "FAILED_SAFE") {
    return finalizeShadow({
      ...base,
      status: "FAILED_SAFE",
      plannerResolution,
      reentry: null,
      counts: { accepted: 0, resolved: 0, ready: 0, ranked: 0 },
      integrity: {
        sourceCandidatesPreserved: true,
        productionCandidateMerge: false,
        productionRouteChanged: false,
      },
    });
  }
  if (!acceptedCount) {
    return finalizeShadow({
      ...base,
      status: "SKIPPED",
      plannerResolution,
      reentry: null,
      counts: { accepted: 0, resolved: 0, ready: 0, ranked: 0 },
      integrity: {
        sourceCandidatesPreserved: true,
        productionCandidateMerge: false,
        productionRouteChanged: false,
      },
    });
  }
  try {
    const reentry = runPlannerResolverReentry({
      plannerResolution,
      deterministicSemanticProfile: semanticProfile,
      resolvedSemanticProfile,
    });
    const validationList = Object.values(reentry.validations);
    const allValid = validationList.every(
      (validation) => validation.valid === true,
    );
    const resolved = reentry.items.filter(
      (item) => item.resolverResult === "RESOLVED",
    ).length;
    const ready = reentry.items.filter(
      (item) => item.feasibilityStatus === "READY",
    ).length;
    const ranked = reentry.items.filter(
      (item) => item.rankingDisposition === "RANKED",
    ).length;
    return finalizeShadow({
      ...base,
      status: allValid ? "SHADOW_COMPLETED" : "FAILED_SAFE",
      plannerResolution,
      reentry: {
        version: QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION,
        bundleSha256: reentry.bundle.bundleSha256,
        retrievalSha256: reentry.bundle.retrieval.retrievalSha256,
        capabilityManifestSha256:
          reentry.bundle.capabilityManifest.manifestSha256,
        candidateResolutionSha256: reentry.candidateResolution.resolutionSha256,
        candidateFamilyResolutionSha256:
          reentry.candidateFamilyResolution.familyResolutionSha256,
        candidateFeasibilityResolutionSha256:
          reentry.candidateFeasibilityResolution.feasibilityResolutionSha256,
        candidateRankingResolutionSha256:
          reentry.candidateRankingResolution.rankingResolutionSha256,
        validations: reentry.validations,
        items: reentry.items,
      },
      counts: { accepted: acceptedCount, resolved, ready, ranked },
      integrity: {
        sourceCandidatesPreserved: true,
        acceptedProposalCountMatches: acceptedCount === reentry.items.length,
        allReentryValid: allValid,
        productionCandidateMerge: false,
        productionReadyAssignment: false,
        productionRouteChanged: false,
      },
    });
  } catch (error) {
    return finalizeShadow({
      ...base,
      status: "FAILED_SAFE",
      plannerResolution,
      reentry: null,
      failure: {
        code: normalizeText(error?.code || "PLANNER_REENTRY_FAILED"),
        message: normalizeText(error?.message || "Planner re-entry failed"),
      },
      counts: { accepted: acceptedCount, resolved: 0, ready: 0, ranked: 0 },
      integrity: {
        sourceCandidatesPreserved: true,
        productionCandidateMerge: false,
        productionReadyAssignment: false,
        productionRouteChanged: false,
      },
    });
  }
}

function validateCandidatePlannerShadowResolution(document = {}) {
  const errors = [];
  if (document.version !== QUERY_CANDIDATE_PLANNER_SHADOW_RESOLUTION_VERSION) {
    errors.push({ path: "version", code: "INVALID_VERSION" });
  }
  if (
    document.policy?.version !== QUERY_CANDIDATE_PLANNER_SHADOW_POLICY_VERSION
  ) {
    errors.push({ path: "policy.version", code: "INVALID_POLICY_VERSION" });
  }
  if (!SHADOW_STATUS.includes(document.status)) {
    errors.push({ path: "status", code: "INVALID_SHADOW_STATUS" });
  }
  if (
    document.policy?.shadowOnly !== true ||
    document.policy?.productionCandidateMerge !== false ||
    document.policy?.productionReadyAssignment !== false ||
    document.policy?.productionRouteChanged !== false
  ) {
    errors.push({ path: "policy", code: "SHADOW_BOUNDARY_VIOLATION" });
  }
  if (
    document.privacy?.rawRowsSent !== false ||
    document.privacy?.sampleValuesSent !== false ||
    document.privacy?.originalFileSent !== false ||
    document.privacy?.fileNameSent !== false
  ) {
    errors.push({ path: "privacy", code: "PRIVACY_BOUNDARY_VIOLATION" });
  }
  for (const item of asArray(document.reentry?.items)) {
    if (
      item.productionCandidateMerged !== false ||
      item.productionReadyAssigned !== false
    ) {
      errors.push({
        path: `reentry.items.${item.candidateId}`,
        code: "PRODUCTION_MUTATION",
      });
    }
  }
  const expected = sha256({ ...document, shadowResolutionSha256: undefined });
  if (document.shadowResolutionSha256 !== expected) {
    errors.push({ path: "shadowResolutionSha256", code: "SHA_MISMATCH" });
  }
  return {
    version: "query_candidate_planner_shadow_validation_v1",
    valid: errors.length === 0,
    errorCount: errors.length,
    errors,
  };
}

module.exports = {
  QUERY_CANDIDATE_PLANNER_SHADOW_RESOLUTION_VERSION,
  QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION,
  QUERY_CANDIDATE_PLANNER_SHADOW_POLICY_VERSION,
  QUERY_CANDIDATE_PLANNER_SHADOW_MODE,
  SHADOW_STATUS,
  recipeIdForProposal,
  buildPlannerReentryBundle,
  runPlannerResolverReentry,
  runCandidatePlannerLiveShadow,
  validateCandidatePlannerShadowResolution,
};
