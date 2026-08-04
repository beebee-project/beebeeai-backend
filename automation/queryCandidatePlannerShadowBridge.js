const { normalizeText, sha256 } = require("./queryCandidateObservation");
const { assessCandidate } = require("./queryCandidateRetriever");
const {
  runConditionalCandidatePlanner,
  validateQueryCandidatePlannerResolution,
  buildPlannerInput,
  ALLOWED_OPERATIONS,
  QUERY_CANDIDATE_PLANNER_POLICY_VERSION,
  QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION,
  DEFAULT_MODEL,
  DEFAULT_REASONING_EFFORT,
} = require("./queryCandidatePlanner");
const {
  buildQueryCandidateResolution,
  validateQueryCandidateResolution,
  QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION,
} = require("./queryCandidateResolver");
const {
  buildQueryCandidateFamilyResolution,
  validateQueryCandidateFamilyResolution,
  QUERY_CANDIDATE_FAMILY_POLICY_VERSION,
} = require("./queryCandidateFamilyResolver");
const {
  buildQueryCandidateFeasibilityResolution,
  validateQueryCandidateFeasibilityResolution,
  QUERY_CANDIDATE_FEASIBILITY_POLICY_VERSION,
} = require("./queryCandidateFeasibilityGate");
const {
  buildQueryCandidateRankingResolution,
  validateQueryCandidateRankingResolution,
  QUERY_CANDIDATE_RANKING_POLICY_VERSION,
} = require("./queryCandidateRanker");
const {
  CACHE_LAYERS,
  CACHE_ARTIFACT_TYPES,
  CACHE_READ_SOURCE,
  buildHierarchicalCacheIdentity,
  createPlannerProviderHierarchicalCacheAdapter,
} = require("./queryCandidatePlannerHierarchicalEncryptedCache");
const {
  buildUploadInvalidationTags,
} = require("./queryCandidatePlannerCacheOperationalControls");

const QUERY_CANDIDATE_PLANNER_SHADOW_RESOLUTION_VERSION =
  "query_candidate_planner_shadow_resolution_v1";
const QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION =
  "query_candidate_planner_reentry_bundle_v1";
const QUERY_CANDIDATE_PLANNER_SHADOW_POLICY_VERSION =
  "conditional_llm_candidate_planner_shadow_policy_v1";
const QUERY_CANDIDATE_PLANNER_SHADOW_MODE = "LIVE_SHADOW";
const QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_INTEGRATION_VERSION =
  "query_candidate_planner_shadow_cache_integration_v1";
const QUERY_CANDIDATE_PLANNER_SHADOW_REENTRY_CACHE_ARTIFACT_VERSION =
  "query_candidate_planner_shadow_reentry_cache_artifact_v1";
const QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_POLICY_VERSION =
  "encrypted_shadow_reentry_cache_policy_v1";

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

function sanitizeSemanticProfileForPersistentCache(profile = {}) {
  const clone = deepClone(profile);
  function visit(value) {
    if (Array.isArray(value)) return value.map(visit);
    if (!value || typeof value !== "object") return value;
    const output = {};
    for (const [key, item] of Object.entries(value)) {
      const normalizedKey = normalizeLoose(key);
      if (
        [
          "samplevalues",
          "rawrows",
          "originalfile",
          "originalfilebytes",
        ].includes(normalizedKey)
      ) {
        output[key] = Array.isArray(item) ? [] : "";
        continue;
      }
      if (["filename", "originalfilename"].includes(normalizedKey)) {
        output[key] = "";
        continue;
      }
      output[key] = visit(item);
    }
    return output;
  }
  const sanitized = visit(clone);
  delete sanitized.profileSha256;
  delete sanitized.semanticProfileSha256;
  sanitized.profileSha256 = sha256({
    ...sanitized,
    profileSha256: undefined,
    semanticProfileSha256: undefined,
  });
  return sanitized;
}

function persistentCachePrivacyBoundaryValid(value) {
  let valid = true;
  function visit(item) {
    if (!valid || item == null) return;
    if (Array.isArray(item)) {
      item.forEach(visit);
      return;
    }
    if (typeof item !== "object") return;
    for (const [key, nested] of Object.entries(item)) {
      const normalizedKey = normalizeLoose(key);
      if (["samplevalues", "rawrows"].includes(normalizedKey)) {
        if (Array.isArray(nested) ? nested.length > 0 : Boolean(nested))
          valid = false;
      }
      if (
        [
          "filename",
          "originalfilename",
          "originalfile",
          "originalfilebytes",
        ].includes(normalizedKey)
      ) {
        if (
          Array.isArray(nested)
            ? nested.length > 0
            : Boolean(normalizeText(nested || ""))
        )
          valid = false;
      }
      visit(nested);
    }
  }
  visit(value);
  return valid;
}

function plannerProposalSetSha256(plannerResolution = {}) {
  const accepted = asArray(plannerResolution.proposals)
    .filter((proposal) => proposal.disposition === "ACCEPTED_FOR_REVALIDATION")
    .map((proposal) => ({
      proposalIndex: Number(proposal.proposalIndex || 0),
      candidateId: normalizeText(proposal.candidateId || ""),
      plannerItemSha256: normalizeText(proposal.plannerItemSha256 || ""),
    }))
    .sort(
      (left, right) =>
        left.proposalIndex - right.proposalIndex ||
        left.candidateId.localeCompare(right.candidateId, "ko"),
    );
  return sha256({
    version: QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_INTEGRATION_VERSION,
    plannerInputSha256: normalizeText(
      plannerResolution.source?.inputSha256 || "",
    ),
    semanticProfileSha256: normalizeText(
      plannerResolution.source?.semanticProfileSha256 || "",
    ),
    accepted,
  });
}

function cachePrivacyMetadata() {
  return {
    rawRowsIncluded: false,
    sampleValuesIncluded: false,
    originalFileIncluded: false,
    fileNameIncluded: false,
  };
}

function plannerResolutionCacheIdentity({
  plannerResolution = {},
  hierarchicalCacheConfig = {},
} = {}) {
  return buildHierarchicalCacheIdentity({
    tenantId: hierarchicalCacheConfig.tenantId,
    cacheSecret: hierarchicalCacheConfig.cacheSecret,
    layer: CACHE_LAYERS.L3_SEMANTIC,
    artifactType: CACHE_ARTIFACT_TYPES.PLANNER_RESOLUTION,
    semanticProfileSha256: plannerResolution.source?.semanticProfileSha256,
    plannerInputSha256: plannerResolution.source?.inputSha256,
    model: hierarchicalCacheConfig.model,
    reasoningEffort: hierarchicalCacheConfig.reasoningEffort,
    promptVersion: hierarchicalCacheConfig.promptVersion,
    schemaVersion: QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION,
    plannerPolicyVersion: QUERY_CANDIDATE_PLANNER_POLICY_VERSION,
    extraIdentity: {
      integrationVersion:
        QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_INTEGRATION_VERSION,
    },
  });
}

function shadowReentryCacheIdentity({
  proposalSetSha256,
  plannerResolution = {},
  sanitizedSemanticProfile = {},
  sanitizedResolvedSemanticProfile = {},
  hierarchicalCacheConfig = {},
} = {}) {
  return buildHierarchicalCacheIdentity({
    tenantId: hierarchicalCacheConfig.tenantId,
    cacheSecret: hierarchicalCacheConfig.cacheSecret,
    layer: CACHE_LAYERS.L4_REENTRY,
    artifactType: CACHE_ARTIFACT_TYPES.SHADOW_REENTRY,
    plannerProposalSetSha256: proposalSetSha256,
    model: hierarchicalCacheConfig.model,
    reasoningEffort: hierarchicalCacheConfig.reasoningEffort,
    promptVersion: hierarchicalCacheConfig.promptVersion,
    schemaVersion:
      QUERY_CANDIDATE_PLANNER_SHADOW_REENTRY_CACHE_ARTIFACT_VERSION,
    plannerPolicyVersion: QUERY_CANDIDATE_PLANNER_POLICY_VERSION,
    resolverPolicyVersion: QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION,
    familyPolicyVersion: QUERY_CANDIDATE_FAMILY_POLICY_VERSION,
    feasibilityPolicyVersion: QUERY_CANDIDATE_FEASIBILITY_POLICY_VERSION,
    rankerPolicyVersion: QUERY_CANDIDATE_RANKING_POLICY_VERSION,
    extraIdentity: {
      integrationVersion:
        QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_INTEGRATION_VERSION,
      plannerInputSha256: normalizeText(
        plannerResolution.source?.inputSha256 || "",
      ),
      semanticProfileSha256: normalizeText(
        sanitizedSemanticProfile.profileSha256 || "",
      ),
      resolvedSemanticProfileSha256: normalizeText(
        sanitizedResolvedSemanticProfile.profileSha256 || "",
      ),
    },
  });
}

function buildPublicReentry(reentry = {}) {
  return {
    version: QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION,
    bundleSha256: reentry.bundle.bundleSha256,
    retrievalSha256: reentry.bundle.retrieval.retrievalSha256,
    capabilityManifestSha256: reentry.bundle.capabilityManifest.manifestSha256,
    candidateResolutionSha256: reentry.candidateResolution.resolutionSha256,
    candidateFamilyResolutionSha256:
      reentry.candidateFamilyResolution.familyResolutionSha256,
    candidateFeasibilityResolutionSha256:
      reentry.candidateFeasibilityResolution.feasibilityResolutionSha256,
    candidateRankingResolutionSha256:
      reentry.candidateRankingResolution.rankingResolutionSha256,
    validations: reentry.validations,
    items: reentry.items,
  };
}

function reentryCounts(publicReentry = {}) {
  const items = asArray(publicReentry.items);
  return {
    accepted: items.length,
    resolved: items.filter((item) => item.resolverResult === "RESOLVED").length,
    ready: items.filter((item) => item.feasibilityStatus === "READY").length,
    ranked: items.filter((item) => item.rankingDisposition === "RANKED").length,
  };
}

function buildShadowReentryCacheArtifact({
  reentry = {},
  proposalSetSha256,
  plannerResolution = {},
} = {}) {
  const publicReentry = buildPublicReentry(reentry);
  const counts = reentryCounts(publicReentry);
  const artifact = {
    version: QUERY_CANDIDATE_PLANNER_SHADOW_REENTRY_CACHE_ARTIFACT_VERSION,
    policy: {
      version: QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_POLICY_VERSION,
      shadowOnly: true,
      validatedDeterministicChainRequired: true,
      directPlannerReadyAssignment: false,
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      plaintextPersistenceAllowed: false,
    },
    source: {
      caseId: normalizeText(plannerResolution.source?.caseId || ""),
      plannerInputSha256: normalizeText(
        plannerResolution.source?.inputSha256 || "",
      ),
      semanticProfileSha256: normalizeText(
        plannerResolution.source?.semanticProfileSha256 || "",
      ),
      proposalSetSha256,
      canonicalPlannerReentrySourceSha256: proposalSetSha256,
    },
    privacy: {
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      originalFileIncluded: false,
      fileNameIncluded: false,
    },
    documents: {
      bundle: reentry.bundle,
      candidateResolution: reentry.candidateResolution,
      candidateFamilyResolution: reentry.candidateFamilyResolution,
      candidateFeasibilityResolution: reentry.candidateFeasibilityResolution,
      candidateRankingResolution: reentry.candidateRankingResolution,
    },
    reentry: publicReentry,
    counts,
    integrity: {
      allValid: Object.values(reentry.validations).every(
        (validation) => validation.valid === true,
      ),
      allAcceptedResolvedReadyRanked:
        counts.accepted === counts.resolved &&
        counts.accepted === counts.ready &&
        counts.accepted === counts.ranked,
      persistentPrivacyBoundaryValid: false,
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
    },
  };
  artifact.integrity.persistentPrivacyBoundaryValid =
    persistentCachePrivacyBoundaryValid(artifact);
  artifact.artifactSha256 = sha256({ ...artifact, artifactSha256: undefined });
  return artifact;
}

function validateShadowReentryCacheArtifact(
  artifact = {},
  { expectedProposalSetSha256 = "" } = {},
) {
  const errors = [];
  const documents = artifact.documents || {};
  if (
    artifact.version !==
    QUERY_CANDIDATE_PLANNER_SHADOW_REENTRY_CACHE_ARTIFACT_VERSION
  ) {
    errors.push({ code: "INVALID_ARTIFACT_VERSION" });
  }
  if (
    artifact.policy?.version !==
    QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_POLICY_VERSION
  ) {
    errors.push({ code: "INVALID_CACHE_POLICY_VERSION" });
  }
  if (
    expectedProposalSetSha256 &&
    artifact.source?.proposalSetSha256 !== expectedProposalSetSha256
  ) {
    errors.push({ code: "PROPOSAL_SET_SHA_MISMATCH" });
  }
  if (!persistentCachePrivacyBoundaryValid(artifact)) {
    errors.push({ code: "PERSISTENT_PRIVACY_BOUNDARY_VIOLATION" });
  }
  const validations = {
    resolver: validateQueryCandidateResolution(
      documents.candidateResolution || {},
    ),
    family: validateQueryCandidateFamilyResolution(
      documents.candidateFamilyResolution || {},
    ),
    feasibility: validateQueryCandidateFeasibilityResolution(
      documents.candidateFeasibilityResolution || {},
    ),
    ranking: validateQueryCandidateRankingResolution(
      documents.candidateRankingResolution || {},
    ),
  };
  if (
    !Object.values(validations).every((validation) => validation.valid === true)
  ) {
    errors.push({ code: "DETERMINISTIC_VALIDATION_FAILED" });
  }
  const bundle = documents.bundle || {};
  const expectedBundleSha256 = sha256({
    version: QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION,
    acceptedCandidateIds: asArray(bundle.acceptedCandidateIds),
    capabilityManifestSha256: bundle.capabilityManifest?.manifestSha256,
    retrievalSha256: bundle.retrieval?.retrievalSha256,
    resolvedSemanticProfileSha256:
      bundle.resolvedSemanticProfile?.profileSha256,
  });
  if (bundle.bundleSha256 !== expectedBundleSha256) {
    errors.push({ code: "BUNDLE_SHA_MISMATCH" });
  }
  const expectedArtifactSha256 = sha256({
    ...artifact,
    artifactSha256: undefined,
  });
  if (artifact.artifactSha256 !== expectedArtifactSha256) {
    errors.push({ code: "ARTIFACT_SHA_MISMATCH" });
  }
  const publicReentry = {
    version: QUERY_CANDIDATE_PLANNER_REENTRY_BUNDLE_VERSION,
    bundleSha256: bundle.bundleSha256,
    retrievalSha256: bundle.retrieval?.retrievalSha256,
    capabilityManifestSha256: bundle.capabilityManifest?.manifestSha256,
    candidateResolutionSha256: documents.candidateResolution?.resolutionSha256,
    candidateFamilyResolutionSha256:
      documents.candidateFamilyResolution?.familyResolutionSha256,
    candidateFeasibilityResolutionSha256:
      documents.candidateFeasibilityResolution?.feasibilityResolutionSha256,
    candidateRankingResolutionSha256:
      documents.candidateRankingResolution?.rankingResolutionSha256,
    validations,
    items: asArray(artifact.reentry?.items),
  };
  const counts = reentryCounts(publicReentry);
  if (
    counts.accepted !== counts.resolved ||
    counts.accepted !== counts.ready ||
    counts.accepted !== counts.ranked
  ) {
    errors.push({ code: "PARTIAL_REENTRY_NOT_CACHEABLE" });
  }
  return {
    version: "query_candidate_planner_shadow_reentry_cache_validation_v1",
    valid: errors.length === 0,
    errorCount: errors.length,
    errors,
    validations,
    publicReentry,
    counts,
  };
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
  plannerReentrySourceSha256 = "",
} = {}) {
  const accepted = asArray(plannerResolution.proposals).filter(
    (proposal) => proposal.disposition === "ACCEPTED_FOR_REVALIDATION",
  );
  const effectivePlannerReentrySourceSha256 = normalizeText(
    plannerReentrySourceSha256 ||
      plannerResolution.plannerResolutionSha256 ||
      "",
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
    plannerResolutionSha256: effectivePlannerReentrySourceSha256,
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
    plannerResolutionSha256: effectivePlannerReentrySourceSha256,
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
  plannerReentrySourceSha256 = "",
} = {}) {
  const bundle = buildPlannerReentryBundle({
    plannerResolution,
    deterministicSemanticProfile,
    resolvedSemanticProfile,
    plannerReentrySourceSha256,
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
  hierarchicalCache,
  tenantId,
  cacheSecret,
  model,
  reasoningEffort,
  pricing,
  cachePromptVersion = "query_candidate_planner_prompt_v1",
  uploadFingerprintSha256,
  queryJsonSha256,
  plannerResolutionTtlMs,
  reentryTtlMs,
  mode = QUERY_CANDIDATE_PLANNER_SHADOW_MODE,
} = {}) {
  const cacheIntegrationEnabled = Boolean(
    hierarchicalCache &&
    typeof hierarchicalCache.get === "function" &&
    typeof hierarchicalCache.set === "function" &&
    typeof hierarchicalCache.delete === "function",
  );
  const effectiveModel = normalizeText(model || DEFAULT_MODEL);
  const effectiveReasoningEffort = normalizeText(
    reasoningEffort || DEFAULT_REASONING_EFFORT,
  );
  const hierarchicalCacheConfig = cacheIntegrationEnabled
    ? {
        tenantId: normalizeText(tenantId),
        cacheSecret,
        model: effectiveModel,
        reasoningEffort: effectiveReasoningEffort,
        promptVersion: normalizeText(cachePromptVersion),
      }
    : null;
  if (cacheIntegrationEnabled && !hierarchicalCacheConfig.tenantId) {
    throw new Error("Shadow hierarchical cache tenantId가 필요합니다.");
  }
  if (cacheIntegrationEnabled && !hierarchicalCacheConfig.cacheSecret) {
    throw new Error("Shadow hierarchical cache cacheSecret이 필요합니다.");
  }
  const cacheInvalidationTags =
    cacheIntegrationEnabled &&
    (normalizeText(uploadFingerprintSha256 || "") ||
      normalizeText(queryJsonSha256 || ""))
      ? buildUploadInvalidationTags({
          uploadFingerprintSha256,
          queryJsonSha256,
        })
      : {};

  const plannerCache =
    cache ||
    (cacheIntegrationEnabled
      ? createPlannerProviderHierarchicalCacheAdapter({
          hierarchicalCache,
          tenantId: hierarchicalCacheConfig.tenantId,
          cacheSecret: hierarchicalCacheConfig.cacheSecret,
          model: effectiveModel,
          reasoningEffort: effectiveReasoningEffort,
          promptVersion: hierarchicalCacheConfig.promptVersion,
          schemaVersion: QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION,
          plannerPolicyVersion: QUERY_CANDIDATE_PLANNER_POLICY_VERSION,
          invalidationTags: cacheInvalidationTags,
        })
      : undefined);

  const plannerResolution = await runConditionalCandidatePlanner({
    semanticProfile,
    resolvedSemanticProfile,
    candidateResolution,
    candidateFeasibilityResolution,
    candidateRankingResolution,
    provider,
    cache: plannerCache,
    tenantId,
    cacheSecret,
    model: effectiveModel,
    reasoningEffort: effectiveReasoningEffort,
    pricing,
  });
  const plannerValidation =
    validateQueryCandidatePlannerResolution(plannerResolution);
  const base = baseShadowDocument({ plannerResolution, mode });
  const acceptedCount = Number(plannerResolution.counts?.accepted || 0);
  const proposalSetSha256 = acceptedCount
    ? plannerProposalSetSha256(plannerResolution)
    : "";
  const cacheAudit = cacheIntegrationEnabled
    ? {
        version: QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_INTEGRATION_VERSION,
        policyVersion: QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_POLICY_VERSION,
        enabled: true,
        encryptedPersistentOnly: true,
        plaintextPersistenceAllowed: false,
        invalidationTagDigestSha256: Object.keys(cacheInvalidationTags).length
          ? sha256(cacheInvalidationTags)
          : "",
        plannerProvider: {
          invocationStatus: normalizeText(
            plannerResolution.invocation?.status || "",
          ),
          cacheHit: plannerResolution.invocation?.cacheHit === true,
          providerCallCount: Number(
            plannerResolution.invocation?.providerCallCount || 0,
          ),
        },
        plannerResolution: {
          attempted: false,
          hit: false,
          source: CACHE_READ_SOURCE.MISS,
          stored: false,
          valid: plannerValidation.valid === true,
          keyDigest: "",
          payloadSha256: "",
          reason: "NOT_ATTEMPTED",
        },
        reentry: {
          attempted: false,
          hit: false,
          source: CACHE_READ_SOURCE.MISS,
          stored: false,
          valid: false,
          keyDigest: "",
          proposalSetSha256,
          artifactSha256: "",
          reason: "NOT_ATTEMPTED",
          deterministicFallback: false,
        },
      }
    : null;

  if (
    cacheIntegrationEnabled &&
    plannerValidation.valid === true &&
    acceptedCount > 0 &&
    ["CALLED", "CACHE_HIT"].includes(plannerResolution.invocation?.status) &&
    !normalizeText(plannerResolution.invocation?.failureCode || "")
  ) {
    const identity = plannerResolutionCacheIdentity({
      plannerResolution,
      hierarchicalCacheConfig,
    });
    cacheAudit.plannerResolution.attempted = true;
    cacheAudit.plannerResolution.keyDigest = identity.keyDigest;
    const cached = await hierarchicalCache.get({ identity });
    if (cached.hit) {
      const cachedValidation = validateQueryCandidatePlannerResolution(
        cached.value,
      );
      const cachedProposalSetSha256 = plannerProposalSetSha256(cached.value);
      const valid =
        cachedValidation.valid === true &&
        cachedProposalSetSha256 === proposalSetSha256 &&
        persistentCachePrivacyBoundaryValid(cached.value);
      if (valid) {
        cacheAudit.plannerResolution.hit = true;
        cacheAudit.plannerResolution.source = cached.source;
        cacheAudit.plannerResolution.valid = true;
        cacheAudit.plannerResolution.payloadSha256 = cached.payloadSha256;
        cacheAudit.plannerResolution.reason = "VALID_CACHE_HIT";
      } else {
        await hierarchicalCache.delete({ identity });
        cacheAudit.plannerResolution.reason = "INVALID_CACHE_ENTRY_DELETED";
      }
    } else {
      cacheAudit.plannerResolution.reason = normalizeText(
        cached.reason || "MISS",
      );
    }
    if (!cacheAudit.plannerResolution.hit) {
      const stored = await hierarchicalCache.set({
        identity,
        value: plannerResolution,
        ttlMs: plannerResolutionTtlMs,
        metadata: {
          cacheable: true,
          validationValid: plannerValidation.valid === true,
          outcomeStatus: normalizeText(
            plannerResolution.invocation?.status || "",
          ),
          failureCode: normalizeText(
            plannerResolution.invocation?.failureCode || "",
          ),
          privacy: cachePrivacyMetadata(),
          invalidationTags: cacheInvalidationTags,
        },
      });
      cacheAudit.plannerResolution.stored = stored.stored === true;
      cacheAudit.plannerResolution.payloadSha256 = normalizeText(
        stored.payloadSha256 || "",
      );
      cacheAudit.plannerResolution.reason = stored.stored
        ? "STORED_AFTER_MISS"
        : normalizeText(stored.reason || "NOT_STORED");
    }
  }

  function finalizeWithOptionalCache(document) {
    return finalizeShadow(
      cacheAudit ? { ...document, cache: cacheAudit } : document,
    );
  }

  if (plannerResolution.invocation?.status === "REQUIRED_NOT_RUN") {
    return finalizeWithOptionalCache({
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
    return finalizeWithOptionalCache({
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
    return finalizeWithOptionalCache({
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

  const sanitizedSemanticProfile = cacheIntegrationEnabled
    ? sanitizeSemanticProfileForPersistentCache(semanticProfile)
    : semanticProfile;
  const sanitizedResolvedSemanticProfile = cacheIntegrationEnabled
    ? sanitizeSemanticProfileForPersistentCache(
        Object.keys(resolvedSemanticProfile || {}).length
          ? resolvedSemanticProfile
          : semanticProfile,
      )
    : resolvedSemanticProfile;
  const canonicalPlannerReentrySourceSha256 = cacheIntegrationEnabled
    ? proposalSetSha256
    : "";

  try {
    if (cacheIntegrationEnabled) {
      const identity = shadowReentryCacheIdentity({
        proposalSetSha256,
        plannerResolution,
        sanitizedSemanticProfile,
        sanitizedResolvedSemanticProfile,
        hierarchicalCacheConfig,
      });
      cacheAudit.reentry.attempted = true;
      cacheAudit.reentry.keyDigest = identity.keyDigest;
      const cached = await hierarchicalCache.get({ identity });
      if (cached.hit) {
        const cachedValidation = validateShadowReentryCacheArtifact(
          cached.value,
          {
            expectedProposalSetSha256: proposalSetSha256,
          },
        );
        if (cachedValidation.valid) {
          cacheAudit.reentry.hit = true;
          cacheAudit.reentry.source = cached.source;
          cacheAudit.reentry.valid = true;
          cacheAudit.reentry.payloadSha256 = cached.payloadSha256;
          cacheAudit.reentry.artifactSha256 = normalizeText(
            cached.value.artifactSha256 || "",
          );
          cacheAudit.reentry.reason = "VALID_CACHE_HIT";
          return finalizeWithOptionalCache({
            ...base,
            status: "SHADOW_COMPLETED",
            plannerResolution,
            reentry: cachedValidation.publicReentry,
            counts: cachedValidation.counts,
            integrity: {
              sourceCandidatesPreserved: true,
              acceptedProposalCountMatches:
                acceptedCount === cachedValidation.counts.accepted,
              allReentryValid: true,
              productionCandidateMerge: false,
              productionReadyAssignment: false,
              productionRouteChanged: false,
            },
          });
        }
        await hierarchicalCache.delete({ identity });
        cacheAudit.reentry.reason = "INVALID_CACHE_ENTRY_DELETED";
        cacheAudit.reentry.deterministicFallback = true;
      } else {
        cacheAudit.reentry.reason = normalizeText(cached.reason || "MISS");
        cacheAudit.reentry.deterministicFallback =
          normalizeText(cached.reason || "") === "CORRUPT_ENTRY";
      }

      const reentry = runPlannerResolverReentry({
        plannerResolution,
        deterministicSemanticProfile: sanitizedSemanticProfile,
        resolvedSemanticProfile: sanitizedResolvedSemanticProfile,
        plannerReentrySourceSha256: canonicalPlannerReentrySourceSha256,
      });
      const validationList = Object.values(reentry.validations);
      const allValid = validationList.every(
        (validation) => validation.valid === true,
      );
      const publicReentry = buildPublicReentry(reentry);
      const counts = reentryCounts(publicReentry);
      const allAcceptedResolvedReadyRanked =
        counts.accepted === counts.resolved &&
        counts.accepted === counts.ready &&
        counts.accepted === counts.ranked;
      if (allValid && allAcceptedResolvedReadyRanked) {
        const artifact = buildShadowReentryCacheArtifact({
          reentry,
          proposalSetSha256,
          plannerResolution,
        });
        const artifactValidation = validateShadowReentryCacheArtifact(
          artifact,
          {
            expectedProposalSetSha256: proposalSetSha256,
          },
        );
        if (artifactValidation.valid) {
          const stored = await hierarchicalCache.set({
            identity,
            value: artifact,
            ttlMs: reentryTtlMs,
            metadata: {
              cacheable: true,
              validationValid: true,
              outcomeStatus: "SHADOW_COMPLETED",
              failureCode: "",
              privacy: cachePrivacyMetadata(),
              invalidationTags: cacheInvalidationTags,
            },
          });
          cacheAudit.reentry.stored = stored.stored === true;
          cacheAudit.reentry.valid = true;
          cacheAudit.reentry.payloadSha256 = normalizeText(
            stored.payloadSha256 || "",
          );
          cacheAudit.reentry.artifactSha256 = artifact.artifactSha256;
          cacheAudit.reentry.reason = stored.stored
            ? "STORED_AFTER_DETERMINISTIC_REENTRY"
            : normalizeText(stored.reason || "NOT_STORED");
        }
      }
      return finalizeWithOptionalCache({
        ...base,
        status: allValid ? "SHADOW_COMPLETED" : "FAILED_SAFE",
        plannerResolution,
        reentry: publicReentry,
        counts,
        integrity: {
          sourceCandidatesPreserved: true,
          acceptedProposalCountMatches: acceptedCount === counts.accepted,
          allReentryValid: allValid,
          productionCandidateMerge: false,
          productionReadyAssignment: false,
          productionRouteChanged: false,
        },
      });
    }

    const reentry = runPlannerResolverReentry({
      plannerResolution,
      deterministicSemanticProfile: semanticProfile,
      resolvedSemanticProfile,
    });
    const validationList = Object.values(reentry.validations);
    const allValid = validationList.every(
      (validation) => validation.valid === true,
    );
    const publicReentry = buildPublicReentry(reentry);
    const counts = reentryCounts(publicReentry);
    return finalizeWithOptionalCache({
      ...base,
      status: allValid ? "SHADOW_COMPLETED" : "FAILED_SAFE",
      plannerResolution,
      reentry: publicReentry,
      counts,
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
    if (cacheAudit) {
      cacheAudit.reentry.reason = normalizeText(
        error?.code || "PLANNER_REENTRY_FAILED",
      );
    }
    return finalizeWithOptionalCache({
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
  if (document.cache) {
    if (
      document.cache.version !==
        QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_INTEGRATION_VERSION ||
      document.cache.policyVersion !==
        QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_POLICY_VERSION ||
      document.cache.encryptedPersistentOnly !== true ||
      document.cache.plaintextPersistenceAllowed !== false
    ) {
      errors.push({ path: "cache", code: "CACHE_BOUNDARY_VIOLATION" });
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
  QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_INTEGRATION_VERSION,
  QUERY_CANDIDATE_PLANNER_SHADOW_REENTRY_CACHE_ARTIFACT_VERSION,
  QUERY_CANDIDATE_PLANNER_SHADOW_CACHE_POLICY_VERSION,
  SHADOW_STATUS,
  recipeIdForProposal,
  sanitizeSemanticProfileForPersistentCache,
  persistentCachePrivacyBoundaryValid,
  plannerProposalSetSha256,
  buildShadowReentryCacheArtifact,
  validateShadowReentryCacheArtifact,
  buildPlannerReentryBundle,
  runPlannerResolverReentry,
  runCandidatePlannerLiveShadow,
  validateCandidatePlannerShadowResolution,
};
