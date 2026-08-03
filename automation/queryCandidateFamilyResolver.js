"use strict";

const { normalizeText, sha256 } = require("./queryCandidateObservation");

const QUERY_CANDIDATE_FAMILY_RESOLUTION_VERSION =
  "query_candidate_family_resolution_v1";
const QUERY_CANDIDATE_FAMILY_ITEM_VERSION =
  "query_candidate_family_item_v1";
const QUERY_CANDIDATE_FAMILY_MEMBER_VERSION =
  "query_candidate_family_member_v1";
const QUERY_CANDIDATE_FAMILY_POLICY_VERSION =
  "deterministic_candidate_family_policy_v1";

const FAMILY_TYPE = Object.freeze([
  "NAMED_TEMPLATE",
  "STRUCTURAL_RECIPE",
]);
const FAMILY_DISPOSITION = Object.freeze([
  "SELECTED",
  "SUPPRESSED",
  "NOT_APPLICABLE",
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

function normalizeLoose(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[^가-힣a-z0-9]+/gu, "");
}

function sortedUnique(values = []) {
  return unique(values).sort((a, b) => a.localeCompare(b, "ko"));
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

function sourceRootIds(candidate = {}) {
  return sortedUnique(
    asArray(candidate.checks?.sourceScope?.selectedRootIds).length
      ? candidate.checks.sourceScope.selectedRootIds
      : asArray(candidate.matchedPhysicalTableIds).length
        ? candidate.matchedPhysicalTableIds
        : candidate.sourceTableIds,
  );
}

function canonicalMatchedColumns(operand = {}) {
  const values = asArray(operand.matched).map((match) => {
    const columnName = normalizeLoose(match?.columnName || "");
    if (columnName) return columnName;
    return normalizeLoose(match?.columnId || "");
  });
  return sortedUnique(values);
}

function canonicalOperands(candidate = {}) {
  return asArray(candidate.checks?.operandBinding?.operands)
    .map((operand, index) => ({
      index,
      kind: normalizeLoose(operand?.kind || "unknown"),
      expectedToken: normalizeLoose(operand?.expectedToken || ""),
      matchedColumns: canonicalMatchedColumns(operand),
    }))
    .sort((a, b) => {
      const kind = a.kind.localeCompare(b.kind, "en");
      if (kind) return kind;
      const token = a.expectedToken.localeCompare(b.expectedToken, "ko");
      if (token) return token;
      return a.index - b.index;
    })
    .map(({ index, ...operand }) => operand);
}

function canonicalOperation(candidate = {}) {
  const explicit = normalizeLoose(
    candidate.checks?.operandBinding?.operation || "",
  );
  if (explicit) return explicit;
  return normalizeLoose(candidate.recipeId || "unknownrecipe") || "unknownrecipe";
}

function canonicalRecipeClass(candidate = {}) {
  const operation = normalizeLoose(
    candidate.checks?.operandBinding?.operation || "",
  );
  if (operation) return operation;
  return normalizeLoose(candidate.recipeId || "unknownrecipe") || "unknownrecipe";
}

function canonicalTemplateAnchor(candidate = {}) {
  return normalizeLoose(candidate.templateId || "");
}

function familyCategory(operation = "") {
  const value = normalizeLoose(operation);
  if (["singlesourcedashboard"].includes(value)) return "DASHBOARD";
  if (["multisourceschemaunion"].includes(value)) return "MULTI_SOURCE";
  if (["timesum", "timeavg", "cumulativesum"].includes(value)) return "TIME_SERIES";
  if (["groupsum", "groupavg", "groupsummary", "compositionratio"].includes(value)) return "GROUP_AGGREGATION";
  if (["topbottom"].includes(value)) return "RANKING";
  if (["crosssum", "crosscount"].includes(value)) return "CROSS_TAB";
  if (["categorycount", "countrows"].includes(value)) return "COUNTING";
  return "OTHER";
}

function familyDescriptor(candidate = {}) {
  const templateAnchor = canonicalTemplateAnchor(candidate);
  const descriptor = {
    familyType: templateAnchor ? "NAMED_TEMPLATE" : "STRUCTURAL_RECIPE",
    templateAnchor,
    sourceRootIds: sourceRootIds(candidate),
    operation: canonicalOperation(candidate),
    familyCategory: familyCategory(canonicalOperation(candidate)),
    recipeClass: canonicalRecipeClass(candidate),
    operands: canonicalOperands(candidate),
    outputTypes: sortedUnique(
      asArray(candidate.checks?.executorSupport?.outputTypes).map(normalizeLoose),
    ),
  };
  const signature = stableStringify(descriptor);
  return {
    ...descriptor,
    signature,
    familyId: `family_${sha256(signature).slice(0, 20)}`,
  };
}

function bindingPriority(value = "") {
  return {
    BOUND: 4,
    PARTIAL: 3,
    INFERRED: 2,
    UNBOUND: 1,
  }[normalizeText(value).toUpperCase()] || 0;
}

function executorPriority(candidate = {}) {
  return {
    DECLARED: 3,
    GENERIC: 2,
    UNKNOWN: 1,
  }[
    normalizeText(candidate.checks?.executorSupport?.declaredStatus || "")
      .toUpperCase()
  ] || 0;
}

function operandPriority(candidate = {}) {
  return {
    PASS: 3,
    NOT_APPLICABLE: 2,
    UNKNOWN: 1,
    FAIL: 0,
  }[
    normalizeText(candidate.checks?.operandBinding?.status || "NOT_APPLICABLE")
      .toUpperCase()
  ] || 0;
}

function candidatePriority(candidate = {}) {
  return {
    priorRetrieved:
      normalizeText(candidate.previousRetrievalResult).toUpperCase() ===
      "RETRIEVED"
        ? 1
        : 0,
    binding: bindingPriority(candidate.bindingStatus),
    executor: executorPriority(candidate),
    operand: operandPriority(candidate),
    resolutionScore: Number(candidate.resolutionScore || 0),
    originalScore: Number(candidate.originalScore || 0),
    originalRank: Number.isInteger(candidate.originalRank)
      ? candidate.originalRank
      : Number.MAX_SAFE_INTEGER,
    candidateId: normalizeText(candidate.candidateId || ""),
  };
}

function compareCandidates(a = {}, b = {}) {
  const pa = candidatePriority(a);
  const pb = candidatePriority(b);
  for (const key of [
    "priorRetrieved",
    "binding",
    "executor",
    "operand",
    "resolutionScore",
    "originalScore",
  ]) {
    if (pa[key] !== pb[key]) return pb[key] - pa[key];
  }
  if (pa.originalRank !== pb.originalRank) {
    return pa.originalRank - pb.originalRank;
  }
  return pa.candidateId.localeCompare(pb.candidateId, "ko");
}

function representativeDecision(selected = {}, suppressed = {}) {
  const winner = candidatePriority(selected);
  const loser = candidatePriority(suppressed);
  const checks = [
    ["priorRetrieved", "PRIOR_RETRIEVED_PREFERRED"],
    ["binding", "STRONGER_BINDING_PREFERRED"],
    ["executor", "DECLARED_EXECUTOR_PREFERRED"],
    ["operand", "STRONGER_OPERAND_EVIDENCE_PREFERRED"],
    ["resolutionScore", "HIGHER_RESOLUTION_SCORE"],
    ["originalScore", "HIGHER_ORIGINAL_SCORE"],
  ];
  for (const [key, reasonCode] of checks) {
    if (winner[key] !== loser[key]) return reasonCode;
  }
  if (winner.originalRank !== loser.originalRank) {
    return "LOWER_ORIGINAL_RANK";
  }
  return "LEXICOGRAPHIC_TIE_BREAK";
}

function buildFamily(familyId, descriptor, candidates) {
  const ordered = [...candidates].sort(compareCandidates);
  const selected = ordered[0];
  const members = ordered.map((candidate, index) => ({
    candidateId: normalizeText(candidate.candidateId || ""),
    disposition: index === 0 ? "SELECTED" : "SUPPRESSED",
    representativePriority: candidatePriority(candidate),
    suppressionReason:
      index === 0 ? "" : representativeDecision(selected, candidate),
  }));
  const family = {
    version: QUERY_CANDIDATE_FAMILY_ITEM_VERSION,
    familyId,
    familyType: descriptor.familyType,
    signature: descriptor.signature,
    templateAnchor: descriptor.templateAnchor,
    sourceRootIds: descriptor.sourceRootIds,
    operation: descriptor.operation,
    familyCategory: descriptor.familyCategory,
    recipeClass: descriptor.recipeClass,
    operands: descriptor.operands,
    outputTypes: descriptor.outputTypes,
    memberCount: members.length,
    selectedCandidateId: normalizeText(selected.candidateId || ""),
    suppressedCandidateIds: members
      .filter((member) => member.disposition === "SUPPRESSED")
      .map((member) => member.candidateId),
    members,
  };
  family.familySha256 = sha256({ ...family, familySha256: undefined });
  return family;
}

function familySort(a = {}, b = {}) {
  const selectedA = a.members?.[0]?.representativePriority || {};
  const selectedB = b.members?.[0]?.representativePriority || {};
  for (const key of [
    "priorRetrieved",
    "binding",
    "executor",
    "operand",
    "resolutionScore",
    "originalScore",
  ]) {
    const left = Number(selectedA[key] || 0);
    const right = Number(selectedB[key] || 0);
    if (left !== right) return right - left;
  }
  const rankA = Number(selectedA.originalRank ?? Number.MAX_SAFE_INTEGER);
  const rankB = Number(selectedB.originalRank ?? Number.MAX_SAFE_INTEGER);
  if (rankA !== rankB) return rankA - rankB;
  return normalizeText(a.familyId).localeCompare(normalizeText(b.familyId), "en");
}

function buildCandidateMember(candidate, family, familyRank) {
  const member = asArray(family?.members).find(
    (item) => item.candidateId === candidate.candidateId,
  );
  const disposition = member?.disposition || "NOT_APPLICABLE";
  const item = {
    version: QUERY_CANDIDATE_FAMILY_MEMBER_VERSION,
    candidateId: normalizeText(candidate.candidateId || ""),
    resolutionResult: normalizeText(candidate.result || ""),
    familyDisposition: disposition,
    familyId: family?.familyId || "",
    familyRank: Number.isInteger(familyRank) ? familyRank : null,
    selectedCandidateId: family?.selectedCandidateId || "",
    suppressionReason:
      disposition === "SUPPRESSED"
        ? member?.suppressionReason || "DUPLICATE_FAMILY_MEMBER"
        : disposition === "NOT_APPLICABLE"
          ? "RESOLUTION_NOT_RESOLVED"
          : "",
    recipeId: normalizeText(candidate.recipeId || ""),
    templateId: normalizeText(candidate.templateId || ""),
    originalRank: Number.isInteger(candidate.originalRank)
      ? candidate.originalRank
      : null,
    originalScore: Number.isFinite(Number(candidate.originalScore))
      ? Number(candidate.originalScore)
      : null,
    resolutionScore: Number(candidate.resolutionScore || 0),
    sourceRootIds: family?.sourceRootIds || sourceRootIds(candidate),
    operation: family?.operation || canonicalOperation(candidate),
    provenance: {
      resolutionItemVersion: normalizeText(candidate.version || ""),
      resolutionItemSha256: normalizeText(candidate.resolutionItemSha256 || ""),
      sourceCandidateUnmodified: true,
      finalReadyStatusAssigned: false,
    },
  };
  item.familyMemberSha256 = sha256({
    ...item,
    familyMemberSha256: undefined,
  });
  return item;
}

function buildQueryCandidateFamilyResolution({ candidateResolution = {} } = {}) {
  const sourceCandidates = asArray(candidateResolution.candidates);
  const resolvedCandidates = sourceCandidates.filter(
    (candidate) => candidate.result === "RESOLVED",
  );
  const grouped = new Map();
  for (const candidate of resolvedCandidates) {
    const descriptor = familyDescriptor(candidate);
    if (!grouped.has(descriptor.familyId)) {
      grouped.set(descriptor.familyId, { descriptor, candidates: [] });
    }
    grouped.get(descriptor.familyId).candidates.push(candidate);
  }

  const families = [...grouped.entries()]
    .map(([familyId, group]) =>
      buildFamily(familyId, group.descriptor, group.candidates),
    )
    .sort(familySort)
    .map((family, index) => {
      const ranked = { ...family, familyRank: index + 1 };
      ranked.familySha256 = sha256({
        ...ranked,
        familySha256: undefined,
      });
      return ranked;
    });
  const familyByCandidateId = new Map();
  for (const family of families) {
    for (const member of family.members) {
      familyByCandidateId.set(member.candidateId, family);
    }
  }

  const candidates = sourceCandidates.map((candidate) => {
    const family = familyByCandidateId.get(candidate.candidateId);
    return buildCandidateMember(candidate, family, family?.familyRank);
  });
  const selectedCandidateIds = families.map(
    (family) => family.selectedCandidateId,
  );
  const suppressedCandidateIds = families.flatMap(
    (family) => family.suppressedCandidateIds,
  );

  const result = {
    version: QUERY_CANDIDATE_FAMILY_RESOLUTION_VERSION,
    itemVersion: QUERY_CANDIDATE_FAMILY_ITEM_VERSION,
    memberVersion: QUERY_CANDIDATE_FAMILY_MEMBER_VERSION,
    policy: {
      version: QUERY_CANDIDATE_FAMILY_POLICY_VERSION,
      onlyResolvedCandidatesAreGrouped: true,
      exactGenerationIntentSignature: true,
      namedTemplatesAreIsolatedByTemplateAnchor: true,
      sourceScopeParticipatesInFamilyIdentity: true,
      operationParticipatesInFamilyIdentity: true,
      operandAxisAndMeasureParticipateInFamilyIdentity: true,
      outputTypesParticipateInFamilyIdentity: true,
      oneRepresentativePerFamily: true,
      sourceCandidatesAreNotRemovedOrMutated: true,
      finalReadyStatusAssigned: false,
      candidateStatusMutation: false,
    },
    source: {
      caseId: normalizeText(candidateResolution.source?.caseId || ""),
      fileName: normalizeText(candidateResolution.source?.fileName || ""),
      candidateResolutionVersion: normalizeText(candidateResolution.version || ""),
      candidateResolutionPolicyVersion: normalizeText(
        candidateResolution.policy?.version || "",
      ),
      candidateResolutionSha256: normalizeText(
        candidateResolution.resolutionSha256 || "",
      ),
      primaryDomain: normalizeText(
        candidateResolution.source?.primaryDomain || "UNKNOWN",
      ),
      datasetIntent: normalizeText(
        candidateResolution.source?.datasetIntent || "UNKNOWN",
      ),
    },
    integrity: {
      sourceCandidateCount: sourceCandidates.length,
      sourceResolvedCount: resolvedCandidates.length,
      sourceCandidateIdsUnique:
        unique(sourceCandidates.map((candidate) => candidate.candidateId)).length ===
        sourceCandidates.length,
      resolvedCoverageComplete:
        selectedCandidateIds.length + suppressedCandidateIds.length ===
        resolvedCandidates.length,
      exactlyOneSelectedPerFamily: families.every(
        (family) =>
          family.members.filter((member) => member.disposition === "SELECTED")
            .length === 1,
      ),
      nonResolvedCandidatesPreserved:
        candidates.filter(
          (candidate) => candidate.familyDisposition === "NOT_APPLICABLE",
        ).length ===
        sourceCandidates.filter((candidate) => candidate.result !== "RESOLVED")
          .length,
    },
    counts: {
      total: sourceCandidates.length,
      resolvedInput: resolvedCandidates.length,
      stillDeferred: sourceCandidates.filter(
        (candidate) => candidate.result === "STILL_DEFERRED",
      ).length,
      excluded: sourceCandidates.filter(
        (candidate) => candidate.result === "EXCLUDED",
      ).length,
      familyCount: families.length,
      selected: selectedCandidateIds.length,
      suppressed: suppressedCandidateIds.length,
      duplicateFamilyCount: families.filter((family) => family.memberCount > 1)
        .length,
      singletonFamilyCount: families.filter((family) => family.memberCount === 1)
        .length,
      nonResolvedPassThrough: sourceCandidates.filter(
        (candidate) => candidate.result !== "RESOLVED",
      ).length,
    },
    selectedCandidateIds,
    suppressedCandidateIds,
    families,
    candidates,
  };
  result.familyResolutionSha256 = sha256({
    ...result,
    familyResolutionSha256: undefined,
  });
  return result;
}

function issue(path, code, message) {
  return { path, code, message };
}

function validateQueryCandidateFamilyResolution(document = {}) {
  const errors = [];
  const warnings = [];
  if (document.version !== QUERY_CANDIDATE_FAMILY_RESOLUTION_VERSION) {
    errors.push(issue("version", "invalid_version", "family resolution version이 유효하지 않습니다."));
  }
  if (document.itemVersion !== QUERY_CANDIDATE_FAMILY_ITEM_VERSION) {
    errors.push(issue("itemVersion", "invalid_version", "family item version이 유효하지 않습니다."));
  }
  if (document.memberVersion !== QUERY_CANDIDATE_FAMILY_MEMBER_VERSION) {
    errors.push(issue("memberVersion", "invalid_version", "family member version이 유효하지 않습니다."));
  }
  if (document.policy?.version !== QUERY_CANDIDATE_FAMILY_POLICY_VERSION) {
    errors.push(issue("policy.version", "invalid_version", "family policy version이 유효하지 않습니다."));
  }
  const families = asArray(document.families);
  const candidates = asArray(document.candidates);
  const familyIds = new Set();
  const resolvedMemberIds = new Set();
  for (const [index, family] of families.entries()) {
    const path = `families[${index}]`;
    if (family.version !== QUERY_CANDIDATE_FAMILY_ITEM_VERSION) {
      errors.push(issue(`${path}.version`, "invalid_version", "family item version이 유효하지 않습니다."));
    }
    if (!FAMILY_TYPE.includes(family.familyType)) {
      errors.push(issue(`${path}.familyType`, "invalid_enum", "familyType이 유효하지 않습니다."));
    }
    if (!normalizeText(family.familyId)) {
      errors.push(issue(`${path}.familyId`, "required", "familyId가 필요합니다."));
    }
    if (familyIds.has(family.familyId)) {
      errors.push(issue(`${path}.familyId`, "duplicate", "familyId가 중복됩니다."));
    }
    familyIds.add(family.familyId);
    const selected = asArray(family.members).filter(
      (member) => member.disposition === "SELECTED",
    );
    if (selected.length !== 1) {
      errors.push(issue(`${path}.members`, "selected_count", "family마다 대표 후보가 정확히 하나여야 합니다."));
    }
    if (selected[0]?.candidateId !== family.selectedCandidateId) {
      errors.push(issue(`${path}.selectedCandidateId`, "selected_mismatch", "selectedCandidateId가 member와 일치하지 않습니다."));
    }
    if (Number(family.memberCount) !== asArray(family.members).length) {
      errors.push(issue(`${path}.memberCount`, "count_mismatch", "memberCount가 실제 member 수와 다릅니다."));
    }
    for (const member of asArray(family.members)) {
      if (resolvedMemberIds.has(member.candidateId)) {
        errors.push(issue(`${path}.members`, "duplicate_membership", "RESOLVED 후보가 여러 family에 포함됐습니다."));
      }
      resolvedMemberIds.add(member.candidateId);
    }
    const expectedFamilySha = sha256({ ...family, familySha256: undefined });
    if (family.familySha256 !== expectedFamilySha) {
      errors.push(issue(`${path}.familySha256`, "sha_mismatch", "family SHA-256이 일치하지 않습니다."));
    }
  }

  const candidateIds = new Set();
  for (const [index, candidate] of candidates.entries()) {
    const path = `candidates[${index}]`;
    if (candidate.version !== QUERY_CANDIDATE_FAMILY_MEMBER_VERSION) {
      errors.push(issue(`${path}.version`, "invalid_version", "family member version이 유효하지 않습니다."));
    }
    if (!FAMILY_DISPOSITION.includes(candidate.familyDisposition)) {
      errors.push(issue(`${path}.familyDisposition`, "invalid_enum", "family disposition이 유효하지 않습니다."));
    }
    if (candidateIds.has(candidate.candidateId)) {
      errors.push(issue(`${path}.candidateId`, "duplicate", "candidateId가 중복됩니다."));
    }
    candidateIds.add(candidate.candidateId);
    if (
      candidate.resolutionResult === "RESOLVED" &&
      candidate.familyDisposition === "NOT_APPLICABLE"
    ) {
      errors.push(issue(path, "resolved_not_grouped", "RESOLVED 후보는 family에 포함돼야 합니다."));
    }
    if (
      candidate.resolutionResult !== "RESOLVED" &&
      candidate.familyDisposition !== "NOT_APPLICABLE"
    ) {
      errors.push(issue(path, "non_resolved_grouped", "RESOLVED가 아닌 후보는 family에 포함하면 안 됩니다."));
    }
    const expectedMemberSha = sha256({
      ...candidate,
      familyMemberSha256: undefined,
    });
    if (candidate.familyMemberSha256 !== expectedMemberSha) {
      errors.push(issue(`${path}.familyMemberSha256`, "sha_mismatch", "family member SHA-256이 일치하지 않습니다."));
    }
  }

  const expectedCounts = {
    total: candidates.length,
    resolvedInput: candidates.filter(
      (candidate) => candidate.resolutionResult === "RESOLVED",
    ).length,
    stillDeferred: candidates.filter(
      (candidate) => candidate.resolutionResult === "STILL_DEFERRED",
    ).length,
    excluded: candidates.filter(
      (candidate) => candidate.resolutionResult === "EXCLUDED",
    ).length,
    familyCount: families.length,
    selected: candidates.filter(
      (candidate) => candidate.familyDisposition === "SELECTED",
    ).length,
    suppressed: candidates.filter(
      (candidate) => candidate.familyDisposition === "SUPPRESSED",
    ).length,
    duplicateFamilyCount: families.filter((family) => family.memberCount > 1)
      .length,
    singletonFamilyCount: families.filter((family) => family.memberCount === 1)
      .length,
    nonResolvedPassThrough: candidates.filter(
      (candidate) => candidate.resolutionResult !== "RESOLVED",
    ).length,
  };
  for (const [key, expected] of Object.entries(expectedCounts)) {
    if (Number(document.counts?.[key] || 0) !== expected) {
      errors.push(issue(`counts.${key}`, "count_mismatch", `${key} count가 실제 값과 다릅니다.`));
    }
  }
  if (!document.integrity?.sourceCandidateIdsUnique) {
    errors.push(issue("integrity.sourceCandidateIdsUnique", "source_duplicate", "source candidateId가 중복됩니다."));
  }
  if (!document.integrity?.resolvedCoverageComplete) {
    errors.push(issue("integrity.resolvedCoverageComplete", "coverage_incomplete", "RESOLVED 후보 family coverage가 불완전합니다."));
  }
  if (!document.integrity?.exactlyOneSelectedPerFamily) {
    errors.push(issue("integrity.exactlyOneSelectedPerFamily", "selected_count", "family 대표 후보 수가 유효하지 않습니다."));
  }
  if (!document.integrity?.nonResolvedCandidatesPreserved) {
    errors.push(issue("integrity.nonResolvedCandidatesPreserved", "pass_through_mismatch", "비RESOLVED 후보가 그대로 보존되지 않았습니다."));
  }
  const expectedSha = sha256({
    ...document,
    familyResolutionSha256: undefined,
  });
  if (document.familyResolutionSha256 !== expectedSha) {
    errors.push(issue("familyResolutionSha256", "sha_mismatch", "family resolution SHA-256이 일치하지 않습니다."));
  }
  return {
    valid: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings,
  };
}

module.exports = {
  QUERY_CANDIDATE_FAMILY_RESOLUTION_VERSION,
  QUERY_CANDIDATE_FAMILY_ITEM_VERSION,
  QUERY_CANDIDATE_FAMILY_MEMBER_VERSION,
  QUERY_CANDIDATE_FAMILY_POLICY_VERSION,
  FAMILY_TYPE,
  FAMILY_DISPOSITION,
  buildQueryCandidateFamilyResolution,
  validateQueryCandidateFamilyResolution,
  familyDescriptor,
  candidatePriority,
  compareCandidates,
  representativeDecision,
  familyCategory,
};
