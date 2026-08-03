const { normalizeText, sha256 } = require("./queryCandidateObservation");

const QUERY_CANDIDATE_FEASIBILITY_RESOLUTION_VERSION =
  "query_candidate_feasibility_resolution_v1";
const QUERY_CANDIDATE_FEASIBILITY_ITEM_VERSION =
  "query_candidate_feasibility_item_v1";
const QUERY_CANDIDATE_EXECUTION_PLAN_VERSION =
  "query_candidate_execution_plan_v1";
const QUERY_CANDIDATE_FEASIBILITY_POLICY_VERSION =
  "deterministic_candidate_feasibility_policy_v1_1";

const FEASIBILITY_STATUS = Object.freeze([
  "READY",
  "REVIEW",
  "UNSUPPORTED",
  "NOT_APPLICABLE",
]);

const CHECK_STATUS = Object.freeze([
  "PASS",
  "REVIEW",
  "FAIL",
  "NOT_APPLICABLE",
]);

const SUPPORTED_OUTPUT_TYPE = "summarysheet";

const GENERIC_READY_OPERATIONS = Object.freeze({
  countrows: [],
  categorycount: ["group"],
  groupsum: ["group", "measure"],
  groupavg: ["group", "measure"],
  groupsummary: ["group", "measure"],
  compositionratio: ["group", "measure"],
  topbottom: ["group", "measure"],
  timesum: ["period", "measure"],
  timeavg: ["period", "measure"],
  timecount: ["period"],
  cumulativesum: ["period", "measure"],
  crosssum: ["dimension", "dimension", "measure"],
  crosscount: ["dimension", "dimension"],
});

const GENERIC_REVIEW_OPERATIONS = Object.freeze({
  singlesourcedashboard: "GENERIC_DASHBOARD_PLAN_REQUIRES_CONFIRMATION",
  multisourceschemaunion: "MULTI_SOURCE_SCHEMA_PLAN_REQUIRES_CONFIRMATION",
});

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

function issue(path, code, message) {
  return { path, code, message };
}

function reason(level, code, message, details = {}) {
  return { level, code, message, details };
}

function candidateMap(candidateResolution = {}) {
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

function normalizeOperation(candidate = {}, family = {}) {
  return (
    normalizeLoose(family.operation || "") ||
    normalizeLoose(candidate.checks?.operandBinding?.operation || "") ||
    normalizeLoose(candidate.recipeId || "") ||
    "unknownoperation"
  );
}

function sourceRootIds(candidate = {}, family = {}) {
  return sortedUnique(
    asArray(family.sourceRootIds).length
      ? family.sourceRootIds
      : asArray(candidate.checks?.sourceScope?.selectedRootIds).length
        ? candidate.checks.sourceScope.selectedRootIds
        : asArray(candidate.matchedPhysicalTableIds).length
          ? candidate.matchedPhysicalTableIds
          : candidate.sourceTableIds,
  );
}

function columnBelongsToSource(columnId = "", rootIds = []) {
  const value = normalizeText(columnId);
  if (!value) return false;
  return rootIds.some((rootId) => {
    const root = normalizeText(rootId);
    return (
      value === root ||
      value.startsWith(`${root}.`) ||
      value.startsWith(`${root}#`)
    );
  });
}

function operandRows(candidate = {}) {
  return asArray(candidate.checks?.operandBinding?.operands).map((operand) => ({
    kind: normalizeLoose(operand.kind || "unknown"),
    expectedToken: normalizeText(operand.expectedToken || ""),
    status: normalizeText(operand.status || "UNKNOWN").toUpperCase(),
    matched: asArray(operand.matched).map((match) => ({
      columnId: normalizeText(match?.columnId || ""),
      columnName: normalizeText(match?.columnName || ""),
      dataType: normalizeText(match?.dataType || ""),
      semanticType: normalizeText(match?.semanticType || ""),
      semanticRole: normalizeText(match?.semanticRole || ""),
    })),
  }));
}

function requiredKindCounts(operation = "") {
  const required = GENERIC_READY_OPERATIONS[normalizeLoose(operation)] || [];
  return required.reduce((counts, kind) => {
    counts[kind] = Number(counts[kind] || 0) + 1;
    return counts;
  }, {});
}

function matchedOperandBindings(candidate = {}, rootIds = []) {
  return operandRows(candidate).map((operand) => ({
    kind: operand.kind,
    expectedToken: operand.expectedToken,
    columnIds: sortedUnique(
      operand.matched
        .map((match) => match.columnId)
        .filter((columnId) => columnBelongsToSource(columnId, rootIds)),
    ),
    columnNames: sortedUnique(operand.matched.map((match) => match.columnName)),
  }));
}

function roleBindings(candidate = {}, rootIds = []) {
  return asArray(candidate.checks?.requiredRoles).map((role) => ({
    role: normalizeText(role.role || ""),
    status: normalizeText(role.status || "UNKNOWN").toUpperCase(),
    columnIds: sortedUnique(
      asArray(role.matched)
        .map((match) => normalizeText(match?.columnId || ""))
        .filter((columnId) => columnBelongsToSource(columnId, rootIds)),
    ),
    columnNames: sortedUnique(
      asArray(role.matched).map((match) =>
        normalizeText(match?.columnName || ""),
      ),
    ),
  }));
}

function timeCountPeriodFallbackBinding(candidate = {}, rootIds = []) {
  const roleCandidates = roleBindings(candidate, rootIds).filter(
    (binding) =>
      normalizeLoose(binding.role) === "period" &&
      binding.status === "PASS" &&
      binding.columnIds.length > 0,
  );
  const roleColumnIds = sortedUnique(
    roleCandidates.flatMap((binding) => binding.columnIds),
  );
  const roleColumnNames = sortedUnique(
    roleCandidates.flatMap((binding) => binding.columnNames),
  );

  const matchedPeriodColumns = asArray(candidate.matchedColumns).filter(
    (column) => {
      const columnId = normalizeText(column?.columnId || "");
      if (!columnBelongsToSource(columnId, rootIds)) return false;
      const semanticRole = normalizeLoose(column?.semanticRole || "");
      const dataType = normalizeLoose(column?.dataType || "");
      return (
        semanticRole === "period" ||
        dataType === "date" ||
        dataType === "datetime"
      );
    },
  );
  const matchedColumnIds = sortedUnique(
    matchedPeriodColumns.map((column) => column.columnId),
  );
  const matchedColumnNames = sortedUnique(
    matchedPeriodColumns.map(
      (column) => column.header || column.columnName || "",
    ),
  );

  const candidateColumnIds = roleColumnIds.length
    ? roleColumnIds
    : matchedColumnIds;
  const candidateColumnNames = roleColumnIds.length
    ? roleColumnNames
    : matchedColumnNames;

  if (candidateColumnIds.length !== 1) return null;
  return {
    kind: "period",
    expectedToken: candidateColumnNames[0] || "period",
    columnIds: candidateColumnIds,
    columnNames: candidateColumnNames,
    bindingSource: roleColumnIds.length
      ? "REQUIRED_ROLE_PERIOD"
      : "MATCHED_PERIOD_COLUMN",
  };
}

function familySelectionCheck(familyMember = {}, family = {}) {
  const disposition = normalizeText(
    familyMember.familyDisposition || "",
  ).toUpperCase();
  if (disposition !== "SELECTED") {
    return {
      status: "NOT_APPLICABLE",
      reasonCode:
        disposition === "SUPPRESSED"
          ? "DUPLICATE_FAMILY_MEMBER_SUPPRESSED"
          : "CANDIDATE_NOT_SELECTED_FOR_FEASIBILITY",
      familyId: normalizeText(familyMember.familyId || ""),
      selectedCandidateId: normalizeText(
        familyMember.selectedCandidateId || "",
      ),
    };
  }
  const selectedCandidateId = normalizeText(family.selectedCandidateId || "");
  const candidateId = normalizeText(familyMember.candidateId || "");
  return {
    status: selectedCandidateId === candidateId ? "PASS" : "FAIL",
    reasonCode:
      selectedCandidateId === candidateId
        ? "FAMILY_REPRESENTATIVE_SELECTED"
        : "FAMILY_REPRESENTATIVE_MISMATCH",
    familyId: normalizeText(family.familyId || familyMember.familyId || ""),
    selectedCandidateId,
  };
}

function sourceScopeCheck(candidate = {}, family = {}) {
  const rootIds = sourceRootIds(candidate, family);
  const sourceStatus = normalizeText(
    candidate.checks?.sourceScope?.status || "",
  ).toUpperCase();
  const matchedPhysical = sortedUnique(candidate.matchedPhysicalTableIds);
  const matched = sortedUnique([
    ...asArray(candidate.matchedTableIds),
    ...matchedPhysical,
  ]);
  const missing = rootIds.filter(
    (rootId) =>
      !matched.some(
        (tableId) =>
          tableId === rootId ||
          tableId.startsWith(`${rootId}#`) ||
          rootId.startsWith(`${tableId}#`),
      ),
  );
  const status =
    sourceStatus === "PASS" && rootIds.length > 0 && missing.length === 0
      ? "PASS"
      : "FAIL";
  return {
    status,
    reasonCode:
      status === "PASS"
        ? "SOURCE_SCOPE_EXECUTABLE"
        : rootIds.length === 0
          ? "SOURCE_SCOPE_EMPTY"
          : missing.length
            ? "SOURCE_TABLE_NOT_MATCHED"
            : "SOURCE_SCOPE_NOT_RESOLVED",
    sourceRootIds: rootIds,
    matchedPhysicalTableIds: matchedPhysical,
    missingSourceRootIds: missing,
  };
}

function outputContractCheck(candidate = {}, family = {}) {
  const requested = sortedUnique(
    asArray(family.outputTypes).length
      ? family.outputTypes
      : candidate.checks?.executorSupport?.outputTypes,
  ).map(normalizeLoose);
  const supported = requested.includes(SUPPORTED_OUTPUT_TYPE)
    ? [SUPPORTED_OUTPUT_TYPE]
    : [];
  return {
    status: supported.length ? "PASS" : "FAIL",
    reasonCode: supported.length
      ? "SUMMARY_SHEET_OUTPUT_SUPPORTED"
      : "SUMMARY_SHEET_OUTPUT_NOT_SUPPORTED",
    requestedOutputTypes: requested,
    supportedOutputTypes: supported,
    selectedOutputType: supported.length ? SUPPORTED_OUTPUT_TYPE : "",
  };
}

function executorCheck(candidate = {}) {
  const checkStatus = normalizeText(
    candidate.checks?.executorSupport?.status || "",
  ).toUpperCase();
  const declaredStatus = normalizeText(
    candidate.checks?.executorSupport?.declaredStatus || "UNKNOWN",
  ).toUpperCase();
  if (checkStatus !== "PASS") {
    return {
      status: "FAIL",
      declaredStatus,
      reasonCode: "EXECUTOR_SUPPORT_CHECK_FAILED",
    };
  }
  if (declaredStatus === "DECLARED") {
    return {
      status: "PASS",
      declaredStatus,
      reasonCode: "DECLARED_EXECUTOR_AVAILABLE",
    };
  }
  if (declaredStatus === "GENERIC") {
    return {
      status: "PASS",
      declaredStatus,
      reasonCode: "GENERIC_EXECUTOR_AVAILABLE",
    };
  }
  return {
    status: "FAIL",
    declaredStatus,
    reasonCode: "EXECUTOR_NOT_DECLARED_OR_GENERIC",
  };
}

function operationContractCheck(candidate = {}, family = {}, executor = {}) {
  const operation = normalizeOperation(candidate, family);
  const rootIds = sourceRootIds(candidate, family);
  if (executor.declaredStatus === "DECLARED") {
    return {
      status: "PASS",
      operation,
      executionMode: "DECLARED",
      requiredOperandKinds: [],
      reasonCode: "DECLARED_EXECUTOR_OPERATION_CONTRACT",
    };
  }
  if (
    Object.prototype.hasOwnProperty.call(GENERIC_READY_OPERATIONS, operation)
  ) {
    return {
      status: "PASS",
      operation,
      executionMode: "GENERIC_OPERATION",
      requiredOperandKinds: [...GENERIC_READY_OPERATIONS[operation]],
      reasonCode: "GENERIC_OPERATION_SUPPORTED",
    };
  }
  if (
    Object.prototype.hasOwnProperty.call(GENERIC_REVIEW_OPERATIONS, operation)
  ) {
    const invalidSourceCount =
      (operation === "singlesourcedashboard" && rootIds.length !== 1) ||
      (operation === "multisourceschemaunion" && rootIds.length < 2);
    return {
      status: invalidSourceCount ? "FAIL" : "REVIEW",
      operation,
      executionMode: "GENERIC_STRUCTURAL",
      requiredOperandKinds: [],
      reasonCode: invalidSourceCount
        ? "STRUCTURAL_RECIPE_SOURCE_COUNT_INVALID"
        : GENERIC_REVIEW_OPERATIONS[operation],
    };
  }
  return {
    status: "FAIL",
    operation,
    executionMode: "UNKNOWN",
    requiredOperandKinds: [],
    reasonCode: "GENERIC_OPERATION_NOT_SUPPORTED",
  };
}

function operandContractCheck(
  candidate = {},
  family = {},
  operationCheck = {},
) {
  const rootIds = sourceRootIds(candidate, family);
  const bindings = matchedOperandBindings(candidate, rootIds);
  const operation = normalizeLoose(operationCheck.operation || "");
  if (operation === "timecount") {
    const hasPeriodBinding = bindings.some(
      (binding) => binding.kind === "period" && binding.columnIds.length > 0,
    );
    if (!hasPeriodBinding) {
      const fallback = timeCountPeriodFallbackBinding(candidate, rootIds);
      if (fallback) bindings.push(fallback);
    }
  }
  if (operationCheck.executionMode === "DECLARED") {
    const operandStatus = normalizeText(
      candidate.checks?.operandBinding?.status || "NOT_APPLICABLE",
    ).toUpperCase();
    if (operandStatus === "FAIL") {
      return {
        status: "FAIL",
        reasonCode: "DECLARED_RECIPE_OPERAND_BINDING_FAILED",
        requiredKindCounts: {},
        bindings,
        missingKinds: [],
      };
    }
    return {
      status: operandStatus === "PASS" ? "PASS" : "NOT_APPLICABLE",
      reasonCode:
        operandStatus === "PASS"
          ? "DECLARED_RECIPE_OPERANDS_BOUND"
          : "DECLARED_RECIPE_USES_ROLE_CONTRACT",
      requiredKindCounts: {},
      bindings,
      missingKinds: [],
    };
  }
  if (operationCheck.status === "REVIEW") {
    return {
      status: "NOT_APPLICABLE",
      reasonCode: "STRUCTURAL_RECIPE_HAS_NO_EXPLICIT_OPERAND_CONTRACT",
      requiredKindCounts: {},
      bindings,
      missingKinds: [],
    };
  }
  const requiredCounts = requiredKindCounts(operationCheck.operation);
  const missingKinds = [];
  for (const [kind, count] of Object.entries(requiredCounts)) {
    const matching = bindings.filter(
      (binding) => binding.kind === kind && binding.columnIds.length > 0,
    ).length;
    if (matching < count) {
      for (let index = matching; index < count; index += 1)
        missingKinds.push(kind);
    }
  }
  const operandStatus = normalizeText(
    candidate.checks?.operandBinding?.status || "NOT_APPLICABLE",
  ).toUpperCase();
  const timeCountFallbackBound =
    operation === "timecount" &&
    bindings.some(
      (binding) =>
        binding.kind === "period" &&
        binding.columnIds.length === 1 &&
        ["REQUIRED_ROLE_PERIOD", "MATCHED_PERIOD_COLUMN"].includes(
          binding.bindingSource,
        ),
    );
  const status =
    missingKinds.length === 0 &&
    (operandStatus === "PASS" || timeCountFallbackBound)
      ? "PASS"
      : "FAIL";
  return {
    status,
    reasonCode:
      status === "PASS"
        ? timeCountFallbackBound
          ? "TIME_COUNT_PERIOD_CONTRACT_BOUND"
          : "GENERIC_OPERATION_OPERANDS_EXECUTABLE"
        : operation === "timecount" && missingKinds.includes("period")
          ? "TIME_COUNT_PERIOD_COLUMN_MISSING_OR_AMBIGUOUS"
          : operandStatus !== "PASS"
            ? "OPERAND_BINDING_NOT_PASSED"
            : "REQUIRED_OPERAND_COLUMN_MISSING",
    requiredKindCounts: requiredCounts,
    bindings,
    missingKinds,
  };
}

function roleContractCheck(candidate = {}, family = {}) {
  const bindings = roleBindings(candidate, sourceRootIds(candidate, family));
  const failed = bindings.filter(
    (binding) => binding.status !== "PASS" || binding.columnIds.length === 0,
  );
  return {
    status: failed.length ? "FAIL" : "PASS",
    reasonCode: failed.length
      ? "REQUIRED_ROLE_BINDING_FAILED"
      : "REQUIRED_ROLE_BINDINGS_EXECUTABLE",
    bindings,
    failedRoles: failed.map((binding) => binding.role),
  };
}

function capabilityContractCheck(candidate = {}) {
  const capabilities = asArray(candidate.checks?.requiredCapabilities).map(
    (item) => ({
      capability: normalizeText(item.capability || ""),
      status: normalizeText(item.status || "UNKNOWN").toUpperCase(),
    }),
  );
  const failed = capabilities.filter((item) => item.status !== "PASS");
  return {
    status: failed.length ? "FAIL" : "PASS",
    reasonCode: failed.length
      ? "REQUIRED_CAPABILITY_FAILED"
      : "REQUIRED_CAPABILITIES_EXECUTABLE",
    capabilities,
    failedCapabilities: failed.map((item) => item.capability),
  };
}

function constraintContractCheck(candidate = {}) {
  const constraints = asArray(candidate.checks?.constraints).map((item) => ({
    constraint: normalizeText(item.constraint || ""),
    expected: item.expected ?? null,
    actual: item.actual ?? null,
    status: normalizeText(item.status || "UNKNOWN").toUpperCase(),
  }));
  const failed = constraints.filter((item) => item.status !== "PASS");
  return {
    status: failed.length ? "FAIL" : "PASS",
    reasonCode: failed.length
      ? "EXECUTION_CONSTRAINT_FAILED"
      : "EXECUTION_CONSTRAINTS_PASSED",
    constraints,
    failedConstraints: failed.map((item) => item.constraint),
  };
}

function metricContractCheck(candidate = {}) {
  const source = candidate.checks?.metricFamily || {};
  const status = normalizeText(source.status || "NOT_APPLICABLE").toUpperCase();
  const resultStatus =
    status === "FAIL" ? "FAIL" : status === "PASS" ? "PASS" : "NOT_APPLICABLE";
  return {
    status: resultStatus,
    reasonCode:
      resultStatus === "FAIL"
        ? "METRIC_FAMILY_CONTRACT_FAILED"
        : resultStatus === "PASS"
          ? "METRIC_FAMILY_CONTRACT_PASSED"
          : "METRIC_FAMILY_NOT_REQUIRED",
    required: sortedUnique(source.required),
    available: sortedUnique(source.available),
  };
}

function buildExecutionPlan({ candidate = {}, family = {}, checks = {} } = {}) {
  const sourceIds =
    checks.sourceScope?.sourceRootIds || sourceRootIds(candidate, family);
  const plan = {
    version: QUERY_CANDIDATE_EXECUTION_PLAN_VERSION,
    candidateId: normalizeText(candidate.candidateId || ""),
    familyId: normalizeText(family.familyId || ""),
    recipeId: normalizeText(candidate.recipeId || ""),
    templateId: normalizeText(candidate.templateId || ""),
    executorMode: normalizeText(
      checks.operationContract?.executionMode || "UNKNOWN",
    ),
    operation: normalizeText(checks.operationContract?.operation || ""),
    sourceTableIds: sourceIds,
    outputType: normalizeText(checks.outputContract?.selectedOutputType || ""),
    operandBindings: asArray(checks.operandContract?.bindings).map(
      (binding) => ({
        kind: binding.kind,
        expectedToken: binding.expectedToken,
        columnIds: binding.columnIds,
      }),
    ),
    requiredRoleBindings: asArray(checks.roleContract?.bindings).map(
      (binding) => ({
        role: binding.role,
        columnIds: binding.columnIds,
      }),
    ),
    requiresManualConfirmation: checks.operationContract?.status === "REVIEW",
    confirmationReasonCodes:
      checks.operationContract?.status === "REVIEW"
        ? [checks.operationContract.reasonCode]
        : [],
  };
  plan.executionPlanSha256 = sha256({
    ...plan,
    executionPlanSha256: undefined,
  });
  return plan;
}

function statusFromChecks(checks = {}) {
  const hardChecks = [
    checks.familySelection,
    checks.sourceScope,
    checks.outputContract,
    checks.executor,
    checks.operationContract,
    checks.operandContract,
    checks.roleContract,
    checks.capabilityContract,
    checks.constraintContract,
    checks.metricContract,
  ];
  if (hardChecks.some((check) => check?.status === "FAIL"))
    return "UNSUPPORTED";
  if (hardChecks.some((check) => check?.status === "REVIEW")) return "REVIEW";
  return "READY";
}

function reasonsFromChecks(checks = {}, status = "") {
  const result = [];
  const entries = [
    ["sourceScope", checks.sourceScope],
    ["outputContract", checks.outputContract],
    ["executor", checks.executor],
    ["operationContract", checks.operationContract],
    ["operandContract", checks.operandContract],
    ["roleContract", checks.roleContract],
    ["capabilityContract", checks.capabilityContract],
    ["constraintContract", checks.constraintContract],
    ["metricContract", checks.metricContract],
  ];
  for (const [name, check] of entries) {
    if (!check || check.status === "PASS" || check.status === "NOT_APPLICABLE")
      continue;
    result.push(
      reason(
        check.status === "FAIL" ? "BLOCKING" : "WARNING",
        check.reasonCode || `${name.toUpperCase()}_${check.status}`,
        check.status === "FAIL"
          ? `${name} 실행 계약을 충족하지 못했습니다.`
          : `${name} 실행 계획을 추가 확인해야 합니다.`,
        stableClone(check),
      ),
    );
  }
  if (status === "READY") {
    result.push(
      reason(
        "INFO",
        "DETERMINISTIC_FEASIBILITY_READY",
        "source·executor·operation·operand·출력 계약이 결정론적으로 확인됐습니다.",
      ),
    );
  }
  return result;
}

function buildSelectedItem({
  candidate = {},
  familyMember = {},
  family = {},
} = {}) {
  const familySelection = familySelectionCheck(familyMember, family);
  const sourceScope = sourceScopeCheck(candidate, family);
  const outputContract = outputContractCheck(candidate, family);
  const executor = executorCheck(candidate);
  const operationContract = operationContractCheck(candidate, family, executor);
  const operandContract = operandContractCheck(
    candidate,
    family,
    operationContract,
  );
  const roleContract = roleContractCheck(candidate, family);
  const capabilityContract = capabilityContractCheck(candidate);
  const constraintContract = constraintContractCheck(candidate);
  const metricContract = metricContractCheck(candidate);
  const checks = {
    familySelection,
    sourceScope,
    outputContract,
    executor,
    operationContract,
    operandContract,
    roleContract,
    capabilityContract,
    constraintContract,
    metricContract,
  };
  const feasibilityStatus = statusFromChecks(checks);
  const executionPlan = buildExecutionPlan({ candidate, family, checks });
  const item = {
    version: QUERY_CANDIDATE_FEASIBILITY_ITEM_VERSION,
    candidateId: normalizeText(candidate.candidateId || ""),
    familyId: normalizeText(family.familyId || ""),
    familyRank: Number.isInteger(family.familyRank) ? family.familyRank : null,
    familyDisposition: "SELECTED",
    feasibilityStatus,
    recipeId: normalizeText(candidate.recipeId || ""),
    templateId: normalizeText(candidate.templateId || ""),
    checks,
    executionPlan,
    reasons: reasonsFromChecks(checks, feasibilityStatus),
    missingRequirements: reasonsFromChecks(checks, feasibilityStatus)
      .filter((itemReason) => itemReason.level === "BLOCKING")
      .map((itemReason) => ({
        code: itemReason.code,
        message: itemReason.message,
      })),
    provenance: {
      familyMemberVersion: normalizeText(familyMember.version || ""),
      familyMemberSha256: normalizeText(familyMember.familyMemberSha256 || ""),
      resolutionItemVersion: normalizeText(candidate.version || ""),
      resolutionItemSha256: normalizeText(candidate.resolutionItemSha256 || ""),
      sourceCandidateUnmodified: true,
      rankScoreUsedForFeasibility: false,
      productionRouteChanged: false,
    },
  };
  item.feasibilityItemSha256 = sha256({
    ...item,
    feasibilityItemSha256: undefined,
  });
  return item;
}

function buildNotApplicableItem({ candidate = {}, familyMember = {} } = {}) {
  const item = {
    version: QUERY_CANDIDATE_FEASIBILITY_ITEM_VERSION,
    candidateId: normalizeText(
      candidate.candidateId || familyMember.candidateId || "",
    ),
    familyId: normalizeText(familyMember.familyId || ""),
    familyRank: Number.isInteger(familyMember.familyRank)
      ? familyMember.familyRank
      : null,
    familyDisposition: normalizeText(
      familyMember.familyDisposition || "NOT_APPLICABLE",
    ),
    feasibilityStatus: "NOT_APPLICABLE",
    recipeId: normalizeText(candidate.recipeId || familyMember.recipeId || ""),
    templateId: normalizeText(
      candidate.templateId || familyMember.templateId || "",
    ),
    checks: {
      familySelection: familySelectionCheck(familyMember, {}),
      sourceScope: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
      outputContract: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
      executor: { status: "NOT_APPLICABLE", reasonCode: "FEASIBILITY_NOT_RUN" },
      operationContract: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
      operandContract: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
      roleContract: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
      capabilityContract: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
      constraintContract: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
      metricContract: {
        status: "NOT_APPLICABLE",
        reasonCode: "FEASIBILITY_NOT_RUN",
      },
    },
    executionPlan: null,
    reasons: [
      reason(
        "INFO",
        "FEASIBILITY_NOT_APPLICABLE",
        "Family SELECTED 대표 후보가 아니므로 Feasibility Gate를 실행하지 않았습니다.",
        {
          familyDisposition: normalizeText(
            familyMember.familyDisposition || "NOT_APPLICABLE",
          ),
        },
      ),
    ],
    missingRequirements: [],
    provenance: {
      familyMemberVersion: normalizeText(familyMember.version || ""),
      familyMemberSha256: normalizeText(familyMember.familyMemberSha256 || ""),
      resolutionItemVersion: normalizeText(candidate.version || ""),
      resolutionItemSha256: normalizeText(candidate.resolutionItemSha256 || ""),
      sourceCandidateUnmodified: true,
      rankScoreUsedForFeasibility: false,
      productionRouteChanged: false,
    },
  };
  item.feasibilityItemSha256 = sha256({
    ...item,
    feasibilityItemSha256: undefined,
  });
  return item;
}

function buildQueryCandidateFeasibilityResolution({
  candidateFamilyResolution = {},
  candidateResolution = {},
} = {}) {
  const resolutionById = candidateMap(candidateResolution);
  const familyMemberById = familyMemberMap(candidateFamilyResolution);
  const familiesById = familyMap(candidateFamilyResolution);
  const sourceCandidates = asArray(candidateResolution.candidates);
  const candidates = sourceCandidates.map((candidate) => {
    const familyMember = familyMemberById.get(candidate.candidateId) || {};
    const family = familiesById.get(familyMember.familyId) || {};
    return normalizeText(familyMember.familyDisposition).toUpperCase() ===
      "SELECTED"
      ? buildSelectedItem({ candidate, familyMember, family })
      : buildNotApplicableItem({ candidate, familyMember });
  });

  const selectedInputCount = asArray(
    candidateFamilyResolution.selectedCandidateIds,
  ).length;
  const selectedItems = candidates.filter(
    (candidate) => candidate.familyDisposition === "SELECTED",
  );
  const statusIds = (status) =>
    candidates
      .filter((candidate) => candidate.feasibilityStatus === status)
      .map((candidate) => candidate.candidateId);
  const familyCandidateIds = new Set(
    asArray(candidateFamilyResolution.candidates).map(
      (candidate) => candidate.candidateId,
    ),
  );
  const missingFamilyCandidateIds = sourceCandidates
    .map((candidate) => candidate.candidateId)
    .filter((candidateId) => !familyCandidateIds.has(candidateId));
  const orphanFamilyCandidateIds = [...familyCandidateIds].filter(
    (candidateId) => !resolutionById.has(candidateId),
  );

  const result = {
    version: QUERY_CANDIDATE_FEASIBILITY_RESOLUTION_VERSION,
    itemVersion: QUERY_CANDIDATE_FEASIBILITY_ITEM_VERSION,
    executionPlanVersion: QUERY_CANDIDATE_EXECUTION_PLAN_VERSION,
    policy: {
      version: QUERY_CANDIDATE_FEASIBILITY_POLICY_VERSION,
      selectedFamilyRepresentativesOnly: true,
      rankIndependent: true,
      rankScoreUsedForFeasibility: false,
      summarySheetMvpBoundary: true,
      declaredExecutorMayBeReadyFromRoleContract: true,
      genericOperationsRequireExplicitOperandColumns: true,
      timeCountRequiresSinglePeriodOperand: true,
      genericDashboardRequiresReview: true,
      multiSourceSchemaUnionRequiresReview: true,
      sourceCandidatesAreNotRemovedOrMutated: true,
      sourceCandidateStatusMutation: false,
      productionRouteChanged: false,
      feasibilityStatusAssigned: true,
    },
    source: {
      caseId: normalizeText(candidateResolution.source?.caseId || ""),
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
      primaryDomain: normalizeText(
        candidateResolution.source?.primaryDomain || "UNKNOWN",
      ),
      datasetIntent: normalizeText(
        candidateResolution.source?.datasetIntent || "UNKNOWN",
      ),
    },
    integrity: {
      sourceCandidateCount: sourceCandidates.length,
      familyCandidateCount: asArray(candidateFamilyResolution.candidates)
        .length,
      selectedFamilyRepresentativeCount: selectedInputCount,
      missingFamilyCandidateIds,
      orphanFamilyCandidateIds,
      candidateCoverageComplete:
        missingFamilyCandidateIds.length === 0 &&
        orphanFamilyCandidateIds.length === 0,
      selectedCoverageComplete: selectedItems.length === selectedInputCount,
      statusPartitionComplete:
        candidates.length ===
        statusIds("READY").length +
          statusIds("REVIEW").length +
          statusIds("UNSUPPORTED").length +
          statusIds("NOT_APPLICABLE").length,
      sourceCandidatesPreserved: candidates.length === sourceCandidates.length,
      executionPlansContainNoSourceRecords: candidates.every((candidate) => {
        const serialized = stableStringify(candidate.executionPlan || {});
        return !/rawrows|samplevalues|\"rows\"/i.test(serialized);
      }),
    },
    counts: {
      total: candidates.length,
      selectedInput: selectedInputCount,
      ready: statusIds("READY").length,
      review: statusIds("REVIEW").length,
      unsupported: statusIds("UNSUPPORTED").length,
      notApplicable: statusIds("NOT_APPLICABLE").length,
      executionPlanCount: candidates.filter(
        (candidate) => candidate.executionPlan,
      ).length,
    },
    readyCandidateIds: statusIds("READY"),
    reviewCandidateIds: statusIds("REVIEW"),
    unsupportedCandidateIds: statusIds("UNSUPPORTED"),
    notApplicableCandidateIds: statusIds("NOT_APPLICABLE"),
    candidates,
  };
  result.feasibilityResolutionSha256 = sha256({
    ...result,
    feasibilityResolutionSha256: undefined,
  });
  return result;
}

function validateQueryCandidateFeasibilityResolution(document = {}) {
  const errors = [];
  const warnings = [];
  if (document.version !== QUERY_CANDIDATE_FEASIBILITY_RESOLUTION_VERSION) {
    errors.push(
      issue(
        "version",
        "invalid_version",
        "feasibility resolution version이 유효하지 않습니다.",
      ),
    );
  }
  if (document.itemVersion !== QUERY_CANDIDATE_FEASIBILITY_ITEM_VERSION) {
    errors.push(
      issue(
        "itemVersion",
        "invalid_version",
        "feasibility item version이 유효하지 않습니다.",
      ),
    );
  }
  if (
    document.executionPlanVersion !== QUERY_CANDIDATE_EXECUTION_PLAN_VERSION
  ) {
    errors.push(
      issue(
        "executionPlanVersion",
        "invalid_version",
        "execution plan version이 유효하지 않습니다.",
      ),
    );
  }
  if (document.policy?.version !== QUERY_CANDIDATE_FEASIBILITY_POLICY_VERSION) {
    errors.push(
      issue(
        "policy.version",
        "invalid_version",
        "feasibility policy version이 유효하지 않습니다.",
      ),
    );
  }
  const candidates = asArray(document.candidates);
  const ids = new Set();
  for (const [index, candidate] of candidates.entries()) {
    const path = `candidates[${index}]`;
    if (candidate.version !== QUERY_CANDIDATE_FEASIBILITY_ITEM_VERSION) {
      errors.push(
        issue(
          `${path}.version`,
          "invalid_version",
          "feasibility item version이 유효하지 않습니다.",
        ),
      );
    }
    if (!FEASIBILITY_STATUS.includes(candidate.feasibilityStatus)) {
      errors.push(
        issue(
          `${path}.feasibilityStatus`,
          "invalid_enum",
          "feasibilityStatus가 유효하지 않습니다.",
        ),
      );
    }
    if (ids.has(candidate.candidateId)) {
      errors.push(
        issue(`${path}.candidateId`, "duplicate", "candidateId가 중복됩니다."),
      );
    }
    ids.add(candidate.candidateId);
    for (const [checkName, check] of Object.entries(candidate.checks || {})) {
      if (!CHECK_STATUS.includes(check?.status)) {
        errors.push(
          issue(
            `${path}.checks.${checkName}.status`,
            "invalid_enum",
            "check status가 유효하지 않습니다.",
          ),
        );
      }
    }
    if (
      candidate.feasibilityStatus === "NOT_APPLICABLE" &&
      candidate.executionPlan != null
    ) {
      errors.push(
        issue(
          `${path}.executionPlan`,
          "unexpected_plan",
          "NOT_APPLICABLE 후보에는 execution plan이 없어야 합니다.",
        ),
      );
    }
    if (
      candidate.feasibilityStatus !== "NOT_APPLICABLE" &&
      !candidate.executionPlan
    ) {
      errors.push(
        issue(
          `${path}.executionPlan`,
          "required",
          "평가된 후보에는 execution plan이 필요합니다.",
        ),
      );
    }
    if (candidate.executionPlan) {
      const expectedPlanHash = sha256({
        ...candidate.executionPlan,
        executionPlanSha256: undefined,
      });
      if (candidate.executionPlan.executionPlanSha256 !== expectedPlanHash) {
        errors.push(
          issue(
            `${path}.executionPlan.executionPlanSha256`,
            "hash_mismatch",
            "execution plan hash가 일치하지 않습니다.",
          ),
        );
      }
    }
    if (candidate.feasibilityStatus === "READY") {
      const failed = Object.values(candidate.checks || {}).filter(
        (check) => check?.status === "FAIL" || check?.status === "REVIEW",
      );
      if (failed.length) {
        errors.push(
          issue(
            path,
            "ready_with_nonpass_check",
            "READY 후보에 FAIL 또는 REVIEW 검사가 있습니다.",
          ),
        );
      }
      if (candidate.executionPlan?.outputType !== SUPPORTED_OUTPUT_TYPE) {
        errors.push(
          issue(
            `${path}.executionPlan.outputType`,
            "unsupported_output",
            "READY 후보는 summarySheet 실행 계획이어야 합니다.",
          ),
        );
      }
      if (candidate.executionPlan?.requiresManualConfirmation !== false) {
        errors.push(
          issue(
            `${path}.executionPlan.requiresManualConfirmation`,
            "manual_confirmation",
            "READY 후보는 수동 확인이 없어야 합니다.",
          ),
        );
      }
    }
    if (
      candidate.feasibilityStatus === "REVIEW" &&
      candidate.executionPlan?.requiresManualConfirmation !== true
    ) {
      errors.push(
        issue(
          `${path}.executionPlan.requiresManualConfirmation`,
          "review_confirmation",
          "REVIEW 후보는 수동 확인 사유가 필요합니다.",
        ),
      );
    }
    const expectedItemHash = sha256({
      ...candidate,
      feasibilityItemSha256: undefined,
    });
    if (candidate.feasibilityItemSha256 !== expectedItemHash) {
      errors.push(
        issue(
          `${path}.feasibilityItemSha256`,
          "hash_mismatch",
          "feasibility item hash가 일치하지 않습니다.",
        ),
      );
    }
    const serialized = stableStringify(candidate);
    if (/rawrows|samplevalues|\"rows\"/i.test(serialized)) {
      errors.push(
        issue(
          path,
          "privacy_boundary",
          "raw row 또는 sample value가 포함되면 안 됩니다.",
        ),
      );
    }
  }
  const counts = document.counts || {};
  const expectedTotal =
    Number(counts.ready || 0) +
    Number(counts.review || 0) +
    Number(counts.unsupported || 0) +
    Number(counts.notApplicable || 0);
  if (expectedTotal !== Number(counts.total || 0)) {
    errors.push(
      issue(
        "counts",
        "count_mismatch",
        "feasibility 상태 합계가 total과 다릅니다.",
      ),
    );
  }
  if (
    Number(counts.selectedInput || 0) !==
    Number(counts.ready || 0) +
      Number(counts.review || 0) +
      Number(counts.unsupported || 0)
  ) {
    errors.push(
      issue(
        "counts.selectedInput",
        "count_mismatch",
        "selectedInput이 평가 상태 합계와 다릅니다.",
      ),
    );
  }
  if (!document.integrity?.candidateCoverageComplete) {
    errors.push(
      issue(
        "integrity.candidateCoverageComplete",
        "coverage",
        "family와 resolution 후보 coverage가 불완전합니다.",
      ),
    );
  }
  if (!document.integrity?.selectedCoverageComplete) {
    errors.push(
      issue(
        "integrity.selectedCoverageComplete",
        "coverage",
        "SELECTED 대표 후보 coverage가 불완전합니다.",
      ),
    );
  }
  if (!document.integrity?.statusPartitionComplete) {
    errors.push(
      issue(
        "integrity.statusPartitionComplete",
        "partition",
        "feasibility 상태 partition이 불완전합니다.",
      ),
    );
  }
  if (!document.integrity?.sourceCandidatesPreserved) {
    errors.push(
      issue(
        "integrity.sourceCandidatesPreserved",
        "preservation",
        "source 후보가 보존되지 않았습니다.",
      ),
    );
  }
  if (!document.integrity?.executionPlansContainNoSourceRecords) {
    errors.push(
      issue(
        "integrity.executionPlansContainNoSourceRecords",
        "privacy_boundary",
        "execution plan에 raw row가 포함됐습니다.",
      ),
    );
  }
  const expectedHash = sha256({
    ...document,
    feasibilityResolutionSha256: undefined,
  });
  if (document.feasibilityResolutionSha256 !== expectedHash) {
    errors.push(
      issue(
        "feasibilityResolutionSha256",
        "hash_mismatch",
        "feasibility resolution hash가 일치하지 않습니다.",
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
  QUERY_CANDIDATE_FEASIBILITY_RESOLUTION_VERSION,
  QUERY_CANDIDATE_FEASIBILITY_ITEM_VERSION,
  QUERY_CANDIDATE_EXECUTION_PLAN_VERSION,
  QUERY_CANDIDATE_FEASIBILITY_POLICY_VERSION,
  FEASIBILITY_STATUS,
  GENERIC_READY_OPERATIONS,
  GENERIC_REVIEW_OPERATIONS,
  buildQueryCandidateFeasibilityResolution,
  validateQueryCandidateFeasibilityResolution,
};
