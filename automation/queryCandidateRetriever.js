const { normalizeText, sha256 } = require("./queryCandidateObservation");

const QUERY_CANDIDATE_RETRIEVAL_VERSION = "query_candidate_retrieval_v1";
const QUERY_CANDIDATE_RETRIEVAL_ITEM_VERSION =
  "query_candidate_retrieval_item_v1";
const QUERY_CANDIDATE_RETRIEVAL_POLICY_VERSION =
  "deterministic_candidate_retrieval_policy_v1";

const RETRIEVAL_RESULT = Object.freeze(["RETRIEVED", "DEFERRED", "EXCLUDED"]);
const CHECK_STATUS = Object.freeze([
  "PASS",
  "FAIL",
  "UNKNOWN",
  "NOT_APPLICABLE",
]);
const REASON_LEVEL = Object.freeze(["INFO", "WARNING", "BLOCKING"]);
const EXPLICIT_BINDING_STATUS = new Set(["BOUND", "PARTIAL"]);

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
    .replace(/[^가-힣a-z0-9]+/g, "");
}

function normalizeCapability(value = "") {
  return normalizeText(value).normalize("NFKC").toLowerCase();
}

function normalizeDataType(value = "") {
  const key = normalizeText(value).toLowerCase();
  const aliases = {
    period: "date",
    datetime: "date",
    integer: "number",
    float: "number",
    decimal: "number",
    numeric: "number",
    text: "string",
    category: "string",
  };
  return aliases[key] || key;
}

function candidateMap(items = []) {
  const map = new Map();
  for (const item of asArray(items)) {
    const id = normalizeText(item?.candidateId || "");
    if (id && !map.has(id)) map.set(id, item);
  }
  return map;
}

function physicalRootId(table = {}) {
  return normalizeText(
    table.flags?.virtual && table.sourceTableId
      ? table.sourceTableId
      : table.tableId || table.sourceTableId || "",
  );
}

function eligibleTables(profile = {}) {
  return asArray(profile.tables).filter(
    (table) => table?.flags?.analysisEligible === true,
  );
}

function tableAliasSet(table = {}) {
  return new Set(
    unique([table.tableId, table.sourceTableId]).map(normalizeLoose),
  );
}

function matchExplicitSourceTables(sourceTableIds = [], tables = []) {
  const sourceKeys = unique(sourceTableIds).map(normalizeLoose).filter(Boolean);
  if (!sourceKeys.length)
    return { tables: [], matchedKeys: [], missingKeys: [] };
  const matchedKeys = [];
  const matchedTables = tables.filter((table) => {
    const aliases = tableAliasSet(table);
    const tableMatches = sourceKeys.filter((key) => aliases.has(key));
    matchedKeys.push(...tableMatches);
    return tableMatches.length > 0;
  });
  const matchedSet = new Set(matchedKeys);
  return {
    tables: matchedTables,
    matchedKeys: unique(matchedKeys),
    missingKeys: sourceKeys.filter((key) => !matchedSet.has(key)),
  };
}

function selectCandidateScope(candidate = {}, capability = {}, profile = {}) {
  const allEligible = eligibleTables(profile);
  const explicitIds = unique(candidate.sourceTableIds);
  const explicitMatch = matchExplicitSourceTables(explicitIds, allEligible);

  if (explicitIds.length) {
    if (explicitMatch.tables.length && !explicitMatch.missingKeys.length) {
      return {
        status: "PASS",
        mode: "EXPLICIT_SOURCE",
        tables: explicitMatch.tables,
        sourceTableIds: explicitIds,
        reasonCode: "SOURCE_TABLE_MATCHED",
      };
    }
    return {
      status: "UNKNOWN",
      mode: "EXPLICIT_SOURCE_UNRESOLVED",
      tables: explicitMatch.tables.length ? explicitMatch.tables : allEligible,
      sourceTableIds: explicitIds,
      missingSourceKeys: explicitMatch.missingKeys,
      reasonCode: "SOURCE_TABLE_REFERENCE_UNRESOLVED",
    };
  }

  const physicalRoots = new Set(
    allEligible.map(physicalRootId).filter(Boolean),
  );
  if (physicalRoots.size > 1) {
    return {
      status: "UNKNOWN",
      mode: "ALL_ELIGIBLE_AMBIGUOUS",
      tables: allEligible,
      sourceTableIds: [],
      reasonCode: "MULTIPLE_ELIGIBLE_TABLES_AMBIGUOUS",
    };
  }

  const primary = allEligible.filter((table) => table.flags?.primary === true);
  if (primary.length === 1) {
    const root = physicalRootId(primary[0]);
    const related = allEligible.filter(
      (table) => physicalRootId(table) === root,
    );
    return {
      status: "PASS",
      mode: "PRIMARY_SOURCE",
      tables: related.length ? related : primary,
      sourceTableIds: [],
      reasonCode: "PRIMARY_TABLE_SELECTED",
    };
  }

  if (allEligible.length === 1) {
    return {
      status: "PASS",
      mode: "SINGLE_ELIGIBLE_TABLE",
      tables: allEligible,
      sourceTableIds: [],
      reasonCode: "SINGLE_ELIGIBLE_TABLE_SELECTED",
    };
  }

  if (allEligible.length > 1 && physicalRoots.size === 1) {
    return {
      status: "PASS",
      mode: "SINGLE_PHYSICAL_SOURCE",
      tables: allEligible,
      sourceTableIds: [],
      reasonCode: "SINGLE_PHYSICAL_SOURCE_SELECTED",
    };
  }

  if (allEligible.length > 1) {
    return {
      status: "UNKNOWN",
      mode: "ALL_ELIGIBLE_AMBIGUOUS",
      tables: allEligible,
      sourceTableIds: [],
      reasonCode: "MULTIPLE_ELIGIBLE_TABLES_AMBIGUOUS",
    };
  }

  return {
    status: "FAIL",
    mode: "NONE",
    tables: [],
    sourceTableIds: [],
    reasonCode: "NO_ANALYSIS_ELIGIBLE_TABLE",
  };
}

function columnTokens(column = {}) {
  return new Set(
    unique([
      column.semanticRole,
      ...asArray(column.roleAliases),
      column.sourceHeader,
      column.normalizedHeader,
    ]).map(normalizeLoose),
  );
}

function roleMatchesColumn(requirement = {}, column = {}) {
  const roleTokens = unique([requirement.role, ...asArray(requirement.aliases)])
    .map(normalizeLoose)
    .filter(Boolean);
  const tokens = columnTokens(column);
  const roleMatch = roleTokens.some((token) => tokens.has(token));
  if (!roleMatch) return false;

  const expectedDataType = normalizeDataType(requirement.dataType);
  const actualDataType = normalizeDataType(column.dataType);
  if (expectedDataType && expectedDataType !== actualDataType) return false;

  const expectedSemanticType = normalizeText(
    requirement.semanticType,
  ).toLowerCase();
  const actualSemanticType = normalizeText(column.semanticType).toLowerCase();
  if (
    expectedSemanticType &&
    expectedSemanticType !== "unknown" &&
    actualSemanticType &&
    expectedSemanticType !== actualSemanticType
  ) {
    return false;
  }
  return true;
}

function roleChecks(requiredRoles = [], tables = []) {
  return asArray(requiredRoles).map((requirement) => {
    const matched = [];
    for (const table of tables) {
      for (const column of asArray(table.columns)) {
        if (!roleMatchesColumn(requirement, column)) continue;
        matched.push({
          tableId: normalizeText(table.tableId || ""),
          columnId: normalizeText(column.columnId || ""),
          columnName: normalizeText(
            column.normalizedHeader || column.sourceHeader || "",
          ),
          semanticRole: normalizeText(column.semanticRole || ""),
          dataType: normalizeText(column.dataType || ""),
          semanticType: normalizeText(column.semanticType || ""),
        });
      }
    }
    return {
      role: normalizeText(requirement.role || ""),
      aliases: unique(requirement.aliases),
      expectedDataType: normalizeText(requirement.dataType || ""),
      expectedSemanticType: normalizeText(requirement.semanticType || ""),
      status: matched.length ? "PASS" : "FAIL",
      matched,
    };
  });
}

function scopeCapabilities(tables = [], scopeMode = "") {
  const values = new Set();
  for (const table of tables) {
    for (const capability of asArray(table.capabilities)) {
      values.add(normalizeCapability(capability));
    }
  }
  const physicalRoots = new Set(tables.map(physicalRootId).filter(Boolean));
  if (physicalRoots.size === 1) values.add("single_table");
  if (physicalRoots.size > 1) values.add("multi_table");
  values.add(`table_count:${physicalRoots.size}`);
  if (scopeMode === "ALL_ELIGIBLE_AMBIGUOUS") values.add("scope:ambiguous");
  return { values, physicalTableCount: physicalRoots.size };
}

function isRoleCapability(value = "") {
  return normalizeCapability(value).startsWith("column_role:");
}

function capabilityChecks(
  requiredCapabilities = [],
  tables = [],
  scopeMode = "",
) {
  const scope = scopeCapabilities(tables, scopeMode);
  const checks = [];
  for (const capability of unique(requiredCapabilities)) {
    if (isRoleCapability(capability)) continue;
    const normalized = normalizeCapability(capability);
    checks.push({
      capability,
      status: scope.values.has(normalized) ? "PASS" : "FAIL",
    });
  }
  return { checks, physicalTableCount: scope.physicalTableCount };
}

function metricFamilyCheck(capability = {}, tables = []) {
  const requiredFamilies = unique(capability.metricFamilies).filter(
    (family) => !["generic", "inferred"].includes(family.toLowerCase()),
  );
  const requiresMeasure = asArray(capability.requiredColumnRoles).some(
    (role) => normalizeText(role.semanticType).toLowerCase() === "measure",
  );
  if (!requiredFamilies.length || !requiresMeasure) {
    return {
      status: "NOT_APPLICABLE",
      required: requiredFamilies,
      available: unique(
        tables.flatMap((table) => asArray(table.metricFamilies)),
      ),
    };
  }
  const available = unique(
    tables.flatMap((table) => asArray(table.metricFamilies)),
  );
  const availableSet = new Set(available.map((value) => value.toLowerCase()));
  return {
    status: requiredFamilies.some((family) =>
      availableSet.has(family.toLowerCase()),
    )
      ? "PASS"
      : "FAIL",
    required: requiredFamilies,
    available,
  };
}

function constraintChecks(
  capability = {},
  tables = [],
  physicalTableCount = 0,
) {
  const constraints = capability.constraints || {};
  const rows = tables.map((table) => Number(table.shape?.rowCount || 0));
  const maxRows = rows.length ? Math.max(...rows) : 0;
  const checks = [];
  const minimumTableCount = Number(constraints.minimumTableCount || 0);
  const maximumTableCount = Number(constraints.maximumTableCount || 0);
  const minimumRowCount = Number(constraints.minimumRowCount || 0);

  if (minimumTableCount > 0) {
    checks.push({
      constraint: "minimumTableCount",
      expected: minimumTableCount,
      actual: physicalTableCount,
      status: physicalTableCount >= minimumTableCount ? "PASS" : "FAIL",
    });
  }
  if (maximumTableCount > 0) {
    checks.push({
      constraint: "maximumTableCount",
      expected: maximumTableCount,
      actual: physicalTableCount,
      status: physicalTableCount <= maximumTableCount ? "PASS" : "FAIL",
    });
  }
  if (minimumRowCount > 0) {
    checks.push({
      constraint: "minimumRowCount",
      expected: minimumRowCount,
      actual: maxRows,
      status: maxRows >= minimumRowCount ? "PASS" : "FAIL",
    });
  }
  return checks;
}

function reason(code, level, message, details = {}) {
  return { code, level, message, details };
}

function deterministicScore({
  capability,
  scope,
  roles,
  capabilities,
  metricFamily,
  constraints,
}) {
  const bindingBase =
    {
      BOUND: 40,
      PARTIAL: 32,
      INFERRED: 16,
      UNBOUND: 0,
    }[capability.bindingStatus] || 0;
  const roleRatio = roles.length
    ? roles.filter((item) => item.status === "PASS").length / roles.length
    : 1;
  const capabilityRatio = capabilities.length
    ? capabilities.filter((item) => item.status === "PASS").length /
      capabilities.length
    : 1;
  const constraintRatio = constraints.length
    ? constraints.filter((item) => item.status === "PASS").length /
      constraints.length
    : 1;
  const executorPoints =
    capability.executorSupport?.status === "DECLARED"
      ? 10
      : capability.executorSupport?.status === "GENERIC"
        ? 5
        : 0;
  const metricPoints =
    metricFamily.status === "PASS" || metricFamily.status === "NOT_APPLICABLE"
      ? 5
      : 0;
  const scopePenalty =
    scope.status === "UNKNOWN" ? 12 : scope.status === "FAIL" ? 25 : 0;
  const value =
    bindingBase +
    roleRatio * 20 +
    capabilityRatio * 15 +
    constraintRatio * 10 +
    executorPoints +
    metricPoints -
    scopePenalty;
  return Number(Math.max(0, Math.min(100, value)).toFixed(6));
}

function assessCandidate(candidate = {}, capability = {}, profile = {}) {
  const scope = selectCandidateScope(candidate, capability, profile);
  const roles = roleChecks(capability.requiredColumnRoles, scope.tables);
  const capabilityResult = capabilityChecks(
    capability.requiredCapabilities,
    scope.tables,
    scope.mode,
  );
  const metricFamily = metricFamilyCheck(capability, scope.tables);
  const constraints = constraintChecks(
    capability,
    scope.tables,
    capabilityResult.physicalTableCount,
  );
  const reasons = [];

  if (scope.status === "FAIL") {
    reasons.push(
      reason(
        scope.reasonCode,
        "BLOCKING",
        "분석 가능한 테이블이 없어 후보를 생성할 수 없습니다.",
        {},
      ),
    );
  } else if (scope.status === "UNKNOWN") {
    reasons.push(
      reason(
        scope.reasonCode,
        "WARNING",
        "후보가 사용할 원본 테이블을 결정론적으로 확정할 수 없습니다.",
        {
          sourceTableIds: scope.sourceTableIds,
          missingSourceKeys: scope.missingSourceKeys || [],
        },
      ),
    );
  }

  for (const item of roles.filter((check) => check.status === "FAIL")) {
    reasons.push(
      reason(
        "REQUIRED_COLUMN_ROLE_MISSING",
        "BLOCKING",
        `필수 열 역할 '${item.role}'을 찾지 못했습니다.`,
        { role: item.role, aliases: item.aliases },
      ),
    );
  }
  for (const item of capabilityResult.checks.filter(
    (check) => check.status === "FAIL",
  )) {
    reasons.push(
      reason(
        "REQUIRED_CAPABILITY_MISSING",
        "BLOCKING",
        `필수 capability '${item.capability}'를 충족하지 못했습니다.`,
        { capability: item.capability },
      ),
    );
  }
  if (metricFamily.status === "FAIL") {
    reasons.push(
      reason(
        "METRIC_FAMILY_MISMATCH",
        "BLOCKING",
        "후보가 요구하는 측정값 계열이 데이터에 없습니다.",
        { required: metricFamily.required, available: metricFamily.available },
      ),
    );
  }
  for (const item of constraints.filter((check) => check.status === "FAIL")) {
    reasons.push(
      reason(
        "CONSTRAINT_NOT_SATISFIED",
        "BLOCKING",
        `후보 제약조건 '${item.constraint}'을 충족하지 못했습니다.`,
        item,
      ),
    );
  }
  if (capability.bindingStatus === "INFERRED") {
    reasons.push(
      reason(
        "CAPABILITY_BINDING_INFERRED",
        "WARNING",
        "식별자 기반 추론 capability이므로 후속 의미 검토가 필요합니다.",
        {},
      ),
    );
  }
  if (capability.bindingStatus === "UNBOUND") {
    reasons.push(
      reason(
        "CAPABILITY_BINDING_UNBOUND",
        "WARNING",
        "후보 capability가 manifest에 연결되지 않았습니다.",
        {},
      ),
    );
  }
  if (capability.executorSupport?.status !== "DECLARED") {
    reasons.push(
      reason(
        "EXECUTOR_SUPPORT_NOT_DECLARED",
        "WARNING",
        "실행기 지원이 명시적으로 선언되지 않았습니다.",
        { status: capability.executorSupport?.status || "UNKNOWN" },
      ),
    );
  }

  const blocking = reasons.some((item) => item.level === "BLOCKING");
  let result = "DEFERRED";
  if (EXPLICIT_BINDING_STATUS.has(capability.bindingStatus)) {
    if (
      scope.status === "UNKNOWN" ||
      capability.executorSupport?.status === "UNKNOWN"
    ) {
      result = "DEFERRED";
    } else {
      result = blocking ? "EXCLUDED" : "RETRIEVED";
    }
  }

  const matchedColumns = roles.flatMap((item) => item.matched);
  const missingRequirements = reasons
    .filter((item) => item.level === "BLOCKING")
    .map((item) => ({
      code: item.code,
      message: item.message,
      details: item.details,
    }));

  const item = {
    version: QUERY_CANDIDATE_RETRIEVAL_ITEM_VERSION,
    candidateId: normalizeText(
      candidate.candidateId || capability.candidateId || "",
    ),
    recipeId: normalizeText(candidate.recipeId || capability.recipeId || ""),
    templateId: normalizeText(
      candidate.templateId || capability.templateId || "",
    ),
    candidateType: normalizeText(
      candidate.candidateType || capability.candidateType || "UNKNOWN",
    ),
    result,
    bindingStatus: normalizeText(capability.bindingStatus || "UNBOUND"),
    bindingSource: normalizeText(capability.bindingSource || "NONE"),
    observedClass: normalizeText(candidate.observedClass || "UNKNOWN"),
    visibility: normalizeText(candidate.visibility || "VISIBLE"),
    originalRank: Number.isInteger(candidate.rank) ? candidate.rank : null,
    originalScore: Number.isFinite(Number(candidate.score))
      ? Number(candidate.score)
      : null,
    retrievalScore: deterministicScore({
      capability,
      scope,
      roles,
      capabilities: capabilityResult.checks,
      metricFamily,
      constraints,
    }),
    sourceTableIds: unique(candidate.sourceTableIds),
    matchedTableIds: unique(scope.tables.map((table) => table.tableId)),
    matchedPhysicalTableIds: unique(scope.tables.map(physicalRootId)),
    matchedColumnIds: unique(matchedColumns.map((column) => column.columnId)),
    checks: {
      sourceScope: {
        status: scope.status,
        mode: scope.mode,
        requestedSourceTableIds: scope.sourceTableIds,
        matchedTableIds: unique(scope.tables.map((table) => table.tableId)),
        missingSourceKeys: unique(scope.missingSourceKeys),
      },
      executorSupport: {
        status:
          capability.executorSupport?.status === "DECLARED"
            ? "PASS"
            : capability.executorSupport?.status === "UNKNOWN"
              ? "UNKNOWN"
              : "PASS",
        declaredStatus: normalizeText(
          capability.executorSupport?.status || "UNKNOWN",
        ),
        outputTypes: unique(capability.executorSupport?.outputTypes),
      },
      requiredRoles: roles,
      requiredCapabilities: capabilityResult.checks,
      metricFamily,
      constraints,
    },
    reasons,
    missingRequirements,
    evidence: matchedColumns,
    provenance: {
      candidateItemVersion: normalizeText(candidate.version || ""),
      capabilityItemVersion: normalizeText(capability.version || ""),
      semanticProfileVersion: normalizeText(profile.version || ""),
      candidateStatus: normalizeText(candidate.status || ""),
    },
  };
  item.retrievalItemSha256 = sha256({
    ...item,
    retrievalItemSha256: undefined,
  });
  return item;
}

function buildQueryCandidateRetrieval({
  contract = {},
  capabilityManifest = {},
  semanticProfile = {},
} = {}) {
  const contractCandidates = asArray(contract.candidates);
  const capabilityCandidates = asArray(capabilityManifest.candidates);
  const capabilityById = candidateMap(capabilityCandidates);
  const contractIds = unique(
    contractCandidates.map((candidate) => candidate.candidateId),
  );
  const capabilityIds = unique(
    capabilityCandidates.map((candidate) => candidate.candidateId),
  );
  const contractIdSet = new Set(contractIds);
  const capabilityIdSet = new Set(capabilityIds);
  const candidates = contractCandidates.map((candidate) => {
    const capability = capabilityById.get(
      normalizeText(candidate.candidateId),
    ) || {
      version: "",
      candidateId: candidate.candidateId,
      recipeId: candidate.recipeId,
      templateId: candidate.templateId,
      candidateType: candidate.candidateType,
      bindingStatus: "UNBOUND",
      bindingSource: "NONE",
      requiredColumnRoles: [],
      requiredCapabilities: [],
      metricFamilies: [],
      executorSupport: { status: "UNKNOWN", outputTypes: [], reasons: [] },
      constraints: {},
    };
    return assessCandidate(candidate, capability, semanticProfile);
  });

  const retrieval = {
    version: QUERY_CANDIDATE_RETRIEVAL_VERSION,
    itemVersion: QUERY_CANDIDATE_RETRIEVAL_ITEM_VERSION,
    policy: {
      version: QUERY_CANDIDATE_RETRIEVAL_POLICY_VERSION,
      explicitBindingStatuses: [...EXPLICIT_BINDING_STATUS],
      inferredCandidatesAreDeferred: true,
      unboundCandidatesAreDeferred: true,
      onlyExplicitMissingRequirementsAreExcluded: true,
      candidateStatusMutation: false,
    },
    source: {
      caseId: normalizeText(
        contract.source?.caseId || semanticProfile.source?.caseId || "",
      ),
      fileName: normalizeText(
        contract.source?.fileName || semanticProfile.source?.fileName || "",
      ),
      contractVersion: normalizeText(contract.version || ""),
      contractSha256: normalizeText(contract.contractSha256 || ""),
      capabilityManifestVersion: normalizeText(
        capabilityManifest.version || "",
      ),
      capabilityManifestSha256: normalizeText(
        capabilityManifest.manifestSha256 || "",
      ),
      semanticProfileVersion: normalizeText(semanticProfile.version || ""),
      semanticProfileSha256: normalizeText(semanticProfile.profileSha256 || ""),
    },
    integrity: {
      contractCandidateCount: contractCandidates.length,
      capabilityCandidateCount: capabilityCandidates.length,
      missingCapabilityCandidateIds: contractIds.filter(
        (id) => !capabilityIdSet.has(id),
      ),
      orphanCapabilityCandidateIds: capabilityIds.filter(
        (id) => !contractIdSet.has(id),
      ),
      candidateCountMatch:
        contractCandidates.length === capabilityCandidates.length &&
        contractIds.every((id) => capabilityIdSet.has(id)),
    },
    counts: {
      total: candidates.length,
      retrieved: candidates.filter((item) => item.result === "RETRIEVED")
        .length,
      deferred: candidates.filter((item) => item.result === "DEFERRED").length,
      excluded: candidates.filter((item) => item.result === "EXCLUDED").length,
      boundRetrieved: candidates.filter(
        (item) => item.result === "RETRIEVED" && item.bindingStatus === "BOUND",
      ).length,
      partialRetrieved: candidates.filter(
        (item) =>
          item.result === "RETRIEVED" && item.bindingStatus === "PARTIAL",
      ).length,
      inferredDeferred: candidates.filter(
        (item) =>
          item.result === "DEFERRED" && item.bindingStatus === "INFERRED",
      ).length,
      unboundDeferred: candidates.filter(
        (item) =>
          item.result === "DEFERRED" && item.bindingStatus === "UNBOUND",
      ).length,
    },
    candidates,
  };
  retrieval.retrievalSha256 = sha256({
    ...retrieval,
    retrievalSha256: undefined,
  });
  return retrieval;
}

function validationIssue(path, code, message) {
  return { path, code, message };
}

function validateRetrievalItem(item = {}, index = 0) {
  const path = `candidates[${index}]`;
  const errors = [];
  const warnings = [];
  if (item.version !== QUERY_CANDIDATE_RETRIEVAL_ITEM_VERSION) {
    errors.push(
      validationIssue(
        `${path}.version`,
        "invalid_version",
        "retrieval item version이 유효하지 않습니다.",
      ),
    );
  }
  if (!normalizeText(item.candidateId)) {
    errors.push(
      validationIssue(
        `${path}.candidateId`,
        "required",
        "candidateId가 필요합니다.",
      ),
    );
  }
  if (!RETRIEVAL_RESULT.includes(item.result)) {
    errors.push(
      validationIssue(
        `${path}.result`,
        "invalid_enum",
        "result가 유효하지 않습니다.",
      ),
    );
  }
  if (
    !Number.isFinite(Number(item.retrievalScore)) ||
    item.retrievalScore < 0 ||
    item.retrievalScore > 100
  ) {
    errors.push(
      validationIssue(
        `${path}.retrievalScore`,
        "invalid_range",
        "retrievalScore는 0~100이어야 합니다.",
      ),
    );
  }
  for (const reasonItem of asArray(item.reasons)) {
    if (!REASON_LEVEL.includes(reasonItem.level)) {
      errors.push(
        validationIssue(
          `${path}.reasons`,
          "invalid_reason_level",
          "reason level이 유효하지 않습니다.",
        ),
      );
    }
  }
  for (const check of [
    item.checks?.sourceScope,
    item.checks?.executorSupport,
  ]) {
    if (!CHECK_STATUS.includes(check?.status)) {
      errors.push(
        validationIssue(
          `${path}.checks`,
          "invalid_check_status",
          "check status가 유효하지 않습니다.",
        ),
      );
    }
  }
  const expectedSha = sha256({ ...item, retrievalItemSha256: undefined });
  if (item.retrievalItemSha256 !== expectedSha) {
    errors.push(
      validationIssue(
        `${path}.retrievalItemSha256`,
        "sha_mismatch",
        "retrieval item SHA-256이 일치하지 않습니다.",
      ),
    );
  }
  if (item.result === "DEFERRED") {
    warnings.push(
      validationIssue(
        path,
        "candidate_retrieval_deferred",
        "후속 의미 판단 또는 manifest 보강이 필요합니다.",
      ),
    );
  }
  return { errors, warnings };
}

function validateQueryCandidateRetrieval(retrieval = {}) {
  const errors = [];
  const warnings = [];
  if (retrieval.version !== QUERY_CANDIDATE_RETRIEVAL_VERSION) {
    errors.push(
      validationIssue(
        "version",
        "invalid_version",
        "retrieval version이 유효하지 않습니다.",
      ),
    );
  }
  if (retrieval.itemVersion !== QUERY_CANDIDATE_RETRIEVAL_ITEM_VERSION) {
    errors.push(
      validationIssue(
        "itemVersion",
        "invalid_version",
        "item version이 유효하지 않습니다.",
      ),
    );
  }
  if (retrieval.policy?.version !== QUERY_CANDIDATE_RETRIEVAL_POLICY_VERSION) {
    errors.push(
      validationIssue(
        "policy.version",
        "invalid_version",
        "retrieval policy version이 유효하지 않습니다.",
      ),
    );
  }
  if (!Array.isArray(retrieval.candidates)) {
    errors.push(
      validationIssue(
        "candidates",
        "invalid_type",
        "candidates는 배열이어야 합니다.",
      ),
    );
  } else {
    const ids = new Set();
    retrieval.candidates.forEach((item, index) => {
      const validation = validateRetrievalItem(item, index);
      errors.push(...validation.errors);
      warnings.push(...validation.warnings);
      if (ids.has(item.candidateId)) {
        errors.push(
          validationIssue(
            `candidates[${index}].candidateId`,
            "duplicate",
            "candidateId가 중복됩니다.",
          ),
        );
      }
      ids.add(item.candidateId);
    });
  }
  const counts = retrieval.counts || {};
  const candidates = asArray(retrieval.candidates);
  const expectedCounts = {
    total: candidates.length,
    retrieved: candidates.filter((item) => item.result === "RETRIEVED").length,
    deferred: candidates.filter((item) => item.result === "DEFERRED").length,
    excluded: candidates.filter((item) => item.result === "EXCLUDED").length,
    boundRetrieved: candidates.filter(
      (item) => item.result === "RETRIEVED" && item.bindingStatus === "BOUND",
    ).length,
    partialRetrieved: candidates.filter(
      (item) => item.result === "RETRIEVED" && item.bindingStatus === "PARTIAL",
    ).length,
    inferredDeferred: candidates.filter(
      (item) => item.result === "DEFERRED" && item.bindingStatus === "INFERRED",
    ).length,
    unboundDeferred: candidates.filter(
      (item) => item.result === "DEFERRED" && item.bindingStatus === "UNBOUND",
    ).length,
  };
  for (const [key, value] of Object.entries(expectedCounts)) {
    if (Number(counts[key] || 0) !== value) {
      errors.push(
        validationIssue(
          `counts.${key}`,
          "count_mismatch",
          `${key} count가 실제 후보 수와 다릅니다.`,
        ),
      );
    }
  }
  const integrity = retrieval.integrity || {};
  if (
    !Array.isArray(integrity.missingCapabilityCandidateIds) ||
    !Array.isArray(integrity.orphanCapabilityCandidateIds)
  ) {
    errors.push(
      validationIssue(
        "integrity",
        "invalid_type",
        "integrity 후보 ID 목록은 배열이어야 합니다.",
      ),
    );
  }
  if (integrity.missingCapabilityCandidateIds?.length) {
    warnings.push(
      validationIssue(
        "integrity.missingCapabilityCandidateIds",
        "capability_candidates_missing",
        "일부 계약 후보의 capability item이 없습니다.",
      ),
    );
  }
  if (integrity.orphanCapabilityCandidateIds?.length) {
    warnings.push(
      validationIssue(
        "integrity.orphanCapabilityCandidateIds",
        "capability_candidates_orphaned",
        "계약에 없는 capability item이 있습니다.",
      ),
    );
  }
  if (
    candidates.some(
      (item) =>
        !["UNASSESSED", ""].includes(item.provenance?.candidateStatus || ""),
    )
  ) {
    warnings.push(
      validationIssue(
        "candidates",
        "candidate_status_unexpected",
        "후보 계약 status는 retriever에서 변경하지 않아야 합니다.",
      ),
    );
  }
  const expectedSha = sha256({ ...retrieval, retrievalSha256: undefined });
  if (retrieval.retrievalSha256 !== expectedSha) {
    errors.push(
      validationIssue(
        "retrievalSha256",
        "sha_mismatch",
        "retrieval SHA-256이 일치하지 않습니다.",
      ),
    );
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
  QUERY_CANDIDATE_RETRIEVAL_VERSION,
  QUERY_CANDIDATE_RETRIEVAL_ITEM_VERSION,
  QUERY_CANDIDATE_RETRIEVAL_POLICY_VERSION,
  RETRIEVAL_RESULT,
  CHECK_STATUS,
  buildQueryCandidateRetrieval,
  validateQueryCandidateRetrieval,
  assessCandidate,
  selectCandidateScope,
};
