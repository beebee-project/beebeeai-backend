const { sha256, normalizeText } = require("./queryCandidateObservation");
const {
  validateQueryJsonSemanticProfile,
  operationCapabilities,
} = require("./queryJsonSemanticNormalizer");
const { validateSemanticProfile } = require("./querySemanticProfiler");

const RESOLVED_SEMANTIC_PROFILE_VERSION = "resolved_semantic_profile_v1";
const RESOLVED_SEMANTIC_POLICY_VERSION = "semantic_profile_merge_policy_v1";
const RESOLVED_SEMANTIC_SCHEMA_VERSION = "resolved_semantic_profile_schema_v1";

const ROLE_ACCEPT_THRESHOLD = 0.85;
const HEADER_REPLACE_THRESHOLD = 0.98;
const METADATA_ACCEPT_THRESHOLD = 0.8;
const RELATION_ACCEPT_THRESHOLD = 0.75;

const GENERIC_ROLES = new Set([
  "",
  "unknown",
  "dimension",
  "measure",
  "entity",
  "category",
  "group",
  "string",
  "number",
  "metric",
  "text",
]);

const TEMPORAL_ROLES = new Set(["date", "period"]);
const FINANCIAL_ROLES = new Set([
  "amount",
  "revenue",
  "cost",
  "budget",
  "target",
  "actual",
  "variance",
]);
const COUNT_ROLES = new Set(["count", "quantity"]);
const SALES_METRICS = new Set(["sales", "revenue"]);
const EMPTY_METRICS = new Set(["", "NONE", "none", "unknown"]);

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function unique(values = []) {
  const seen = new Set();
  const output = [];
  for (const value of asArray(values)) {
    const normalized = normalizeText(value);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    output.push(normalized);
  }
  return output;
}

function uniqueSorted(values = []) {
  return unique(values).sort((a, b) => a.localeCompare(b));
}

function round(value, digits = 6) {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  const factor = 10 ** digits;
  return Math.round(number * factor) / factor;
}

function clamp(value, minimum = 0, maximum = 1) {
  const number = Number(value);
  if (!Number.isFinite(number)) return minimum;
  return Math.max(minimum, Math.min(maximum, number));
}

function normalizeRole(value = "") {
  return normalizeText(value).toLowerCase();
}

function normalizeMetric(value = "") {
  const normalized = normalizeText(value);
  return EMPTY_METRICS.has(normalized) ? "" : normalized.toLowerCase();
}

function isGenericRole(role = "") {
  return GENERIC_ROLES.has(normalizeRole(role));
}

function isProtectedExplicit(column = {}) {
  return (
    column.roleSource === "explicit" &&
    !isGenericRole(column.semanticRole) &&
    normalizeText(column.evidence?.explicitSemanticRole || "") !== ""
  );
}

function sameRoleFamily(left = "", right = "") {
  const a = normalizeRole(left);
  const b = normalizeRole(right);
  if (!a || !b) return false;
  if (a === b) return true;
  if (TEMPORAL_ROLES.has(a) && TEMPORAL_ROLES.has(b)) return true;
  if (FINANCIAL_ROLES.has(a) && FINANCIAL_ROLES.has(b)) return true;
  if (COUNT_ROLES.has(a) && COUNT_ROLES.has(b)) return true;
  return false;
}

function sameMetricFamily(left = "", right = "") {
  const a = normalizeMetric(left);
  const b = normalizeMetric(right);
  if (!a || !b) return false;
  if (a === b) return true;
  if (SALES_METRICS.has(a) && SALES_METRICS.has(b)) return true;
  return false;
}

function compatibleSemanticType(
  dataType = "unknown",
  semanticType = "unknown",
) {
  const data = normalizeRole(dataType);
  const semantic = normalizeRole(semanticType);
  if (!semantic || semantic === "unknown") return true;
  if (data === "date") return ["temporal", "dimension"].includes(semantic);
  if (data === "number") {
    return ["measure", "dimension", "identifier"].includes(semantic);
  }
  if (data === "boolean") return ["boolean", "dimension"].includes(semantic);
  if (data === "string") {
    return ["dimension", "identifier", "text", "boolean", "temporal"].includes(
      semantic,
    );
  }
  return true;
}

function resolvedSemanticType({ baseColumn, llmColumn, roleAccepted }) {
  const baseType =
    normalizeRole(baseColumn.semanticType || "unknown") || "unknown";
  const llmType =
    normalizeRole(llmColumn.semanticType || "unknown") || "unknown";
  if (
    roleAccepted &&
    llmColumn.confidence >= METADATA_ACCEPT_THRESHOLD &&
    compatibleSemanticType(baseColumn.dataType, llmType)
  ) {
    return llmType;
  }
  return baseType;
}

function acceptedDefaultAggregation({ llmColumn, semanticType }) {
  const aggregation =
    normalizeText(llmColumn.defaultAggregation || "none") || "none";
  if (llmColumn.confidence < METADATA_ACCEPT_THRESHOLD) return "none";
  if (semanticType === "measure") {
    return [
      "sum",
      "average",
      "countRows",
      "countDistinct",
      "min",
      "max",
      "median",
      "ratio",
      "last",
      "first",
    ].includes(aggregation)
      ? aggregation
      : "none";
  }
  if (
    ["dimension", "identifier", "temporal", "text", "boolean"].includes(
      semanticType,
    )
  ) {
    return [
      "none",
      "countRows",
      "countDistinct",
      "min",
      "max",
      "last",
      "first",
    ].includes(aggregation)
      ? aggregation
      : "none";
  }
  return "none";
}

function issue(code, level, message, details = {}) {
  return {
    code: normalizeText(code),
    level: ["INFO", "WARNING", "BLOCKING"].includes(level) ? level : "WARNING",
    message: normalizeText(message),
    details: details && typeof details === "object" ? details : {},
  };
}

function indexProfiles(deterministicProfile, llmProfile) {
  const baseTables = new Map();
  const baseColumns = new Map();
  for (const table of asArray(deterministicProfile.tables)) {
    baseTables.set(table.tableId, table);
    for (const column of asArray(table.columns)) {
      baseColumns.set(column.columnId, { table, column });
    }
  }
  const llmTables = new Map(
    asArray(llmProfile.tableSemantics).map((item) => [item.tableId, item]),
  );
  const llmColumns = new Map(
    asArray(llmProfile.columnSemantics).map((item) => [item.columnId, item]),
  );
  return { baseTables, baseColumns, llmTables, llmColumns };
}

function assertMergeInputs(deterministicProfile, llmProfile) {
  const deterministicValidation =
    validateQueryJsonSemanticProfile(deterministicProfile);
  if (!deterministicValidation.valid) {
    const error = new Error("결정론적 Semantic Profile 검증에 실패했습니다.");
    error.code = "DETERMINISTIC_SEMANTIC_PROFILE_INVALID";
    error.validation = deterministicValidation;
    throw error;
  }
  const llmValidation = validateSemanticProfile(llmProfile);
  if (!llmValidation.valid) {
    const error = new Error("LLM Semantic Profile 검증에 실패했습니다.");
    error.code = "LLM_SEMANTIC_PROFILE_INVALID";
    error.validation = llmValidation;
    throw error;
  }
  if (
    llmProfile.source?.semanticProfileSha256 !==
    deterministicProfile.profileSha256
  ) {
    const error = new Error(
      "LLM profile이 다른 결정론적 profile을 참조합니다.",
    );
    error.code = "SEMANTIC_PROFILE_SOURCE_HASH_MISMATCH";
    throw error;
  }
  const { baseTables, baseColumns, llmTables, llmColumns } = indexProfiles(
    deterministicProfile,
    llmProfile,
  );
  if (llmTables.size > baseTables.size) {
    const error = new Error(
      "LLM table semantics 수가 결정론적 table 수보다 많습니다.",
    );
    error.code = "TABLE_SEMANTIC_COUNT_OVERFLOW";
    throw error;
  }
  if (llmColumns.size > baseColumns.size) {
    const error = new Error(
      "LLM column semantics 수가 결정론적 column 수보다 많습니다.",
    );
    error.code = "COLUMN_SEMANTIC_COUNT_OVERFLOW";
    throw error;
  }
  for (const tableId of llmTables.keys()) {
    if (!baseTables.has(tableId)) {
      const error = new Error(`존재하지 않는 LLM tableId입니다: ${tableId}`);
      error.code = "LLM_TABLE_REFERENCE_INVALID";
      throw error;
    }
  }
  for (const [columnId, item] of llmColumns.entries()) {
    const base = baseColumns.get(columnId);
    if (!base || item.tableId !== base.table.tableId) {
      const error = new Error(
        `LLM column 참조가 유효하지 않습니다: ${columnId}`,
      );
      error.code = "LLM_COLUMN_REFERENCE_INVALID";
      throw error;
    }
  }
  return { baseTables, baseColumns, llmTables, llmColumns };
}

function resolveRole(baseColumn, llmColumn = {}) {
  const baseRole =
    normalizeRole(baseColumn.semanticRole || "unknown") || "unknown";
  const llmRole =
    normalizeRole(llmColumn.semanticRole || "unknown") || "unknown";
  const decision = normalizeText(llmColumn.decision || "UNKNOWN");
  const confidence = round(clamp(llmColumn.confidence));
  const protectedExplicit = isProtectedExplicit(baseColumn);
  const acceptedFields = [];
  const rejectedFields = [];
  const issues = [];
  let semanticRole = baseRole;
  let selectedSource = protectedExplicit
    ? "ORIGINAL_EXPLICIT"
    : "DETERMINISTIC";
  let resolutionDecision = protectedExplicit
    ? "PRESERVE_EXPLICIT"
    : "PRESERVE_DETERMINISTIC";
  let roleAccepted = false;
  let aliasAccepted = false;

  if (decision === "KEEP" || decision === "UNKNOWN") {
    resolutionDecision = protectedExplicit
      ? decision === "KEEP"
        ? "KEEP_EXPLICIT"
        : "PRESERVE_EXPLICIT"
      : decision === "KEEP"
        ? "KEEP_DETERMINISTIC"
        : "PRESERVE_DETERMINISTIC";
  } else if (confidence < ROLE_ACCEPT_THRESHOLD) {
    rejectedFields.push("semanticRole");
    issues.push(
      issue(
        "LLM_ROLE_LOW_CONFIDENCE",
        "INFO",
        "LLM 역할 제안의 신뢰도가 임계값보다 낮아 유지하지 않았습니다.",
        { baseRole, llmRole, confidence, threshold: ROLE_ACCEPT_THRESHOLD },
      ),
    );
  } else if (protectedExplicit) {
    if (sameRoleFamily(baseRole, llmRole)) {
      aliasAccepted = true;
      acceptedFields.push("roleAlias");
      resolutionDecision = "PRESERVE_EXPLICIT_WITH_LLM_ALIAS";
      selectedSource = "ORIGINAL_EXPLICIT";
    } else {
      rejectedFields.push("semanticRole");
      resolutionDecision = "PRESERVE_EXPLICIT";
      selectedSource = "ORIGINAL_EXPLICIT";
      issues.push(
        issue(
          "EXPLICIT_ROLE_OVERRIDE_REJECTED",
          "WARNING",
          "명시적 원본 역할과 충돌하는 LLM 역할 제안을 거부했습니다.",
          { baseRole, llmRole, confidence },
        ),
      );
    }
  } else if (isGenericRole(baseRole) && !isGenericRole(llmRole)) {
    semanticRole = llmRole;
    roleAccepted = true;
    acceptedFields.push("semanticRole");
    selectedSource = "LLM";
    resolutionDecision =
      decision === "REPLACE" ? "ACCEPT_LLM_REPLACE" : "ACCEPT_LLM_REFINE";
  } else if (isGenericRole(baseRole) && isGenericRole(llmRole)) {
    rejectedFields.push("semanticRole");
    issues.push(
      issue(
        "LLM_ROLE_NOT_MORE_SPECIFIC",
        "INFO",
        "LLM 역할이 결정론적 generic 역할보다 구체적이지 않아 유지하지 않았습니다.",
        { baseRole, llmRole, confidence },
      ),
    );
  } else if (sameRoleFamily(baseRole, llmRole)) {
    aliasAccepted = baseRole !== llmRole;
    if (aliasAccepted) acceptedFields.push("roleAlias");
    resolutionDecision = aliasAccepted
      ? "PRESERVE_DETERMINISTIC_WITH_LLM_ALIAS"
      : "KEEP_DETERMINISTIC";
  } else if (
    baseColumn.roleSource !== "explicit" &&
    decision === "REPLACE" &&
    confidence >= HEADER_REPLACE_THRESHOLD
  ) {
    semanticRole = llmRole;
    roleAccepted = true;
    acceptedFields.push("semanticRole");
    selectedSource = "LLM";
    resolutionDecision = "ACCEPT_LLM_HIGH_CONFIDENCE_REPLACE";
    issues.push(
      issue(
        "DETERMINISTIC_ROLE_REPLACED",
        "INFO",
        "비명시적 결정론 역할을 매우 높은 신뢰도의 LLM 역할로 교체했습니다.",
        { baseRole, llmRole, confidence },
      ),
    );
  } else {
    rejectedFields.push("semanticRole");
    issues.push(
      issue(
        "LLM_ROLE_CONFLICT_REJECTED",
        "WARNING",
        "비명시적 결정론 역할과 충돌하지만 교체 임계값에 미달해 LLM 역할을 거부했습니다.",
        {
          baseRole,
          llmRole,
          confidence,
          threshold: HEADER_REPLACE_THRESHOLD,
        },
      ),
    );
  }

  const aliases = unique([
    ...asArray(baseColumn.roleAliases),
    baseRole,
    roleAccepted || aliasAccepted ? llmRole : "",
  ]);
  return {
    semanticRole,
    roleAliases: aliases,
    roleAccepted,
    selectedSource,
    resolutionDecision,
    acceptedFields,
    rejectedFields,
    issues,
  };
}

function resolveMetric(baseColumn, llmColumn = {}, roleResult, semanticType) {
  const baseMetric = normalizeMetric(baseColumn.metricFamily || "");
  const llmMetric = normalizeMetric(llmColumn.metricFamily || "");
  const confidence = round(clamp(llmColumn.confidence));
  const metricAliases = unique([baseMetric]);
  const acceptedFields = [];
  const rejectedFields = [];
  const issues = [];
  let metricFamily = baseMetric;
  let selectedSource = baseMetric ? "DETERMINISTIC" : "NONE";

  if (!llmMetric) {
    return {
      metricFamily,
      metricAliases,
      selectedSource,
      acceptedFields,
      rejectedFields,
      issues,
    };
  }
  if (baseMetric) {
    if (sameMetricFamily(baseMetric, llmMetric)) {
      metricAliases.push(llmMetric);
      acceptedFields.push("metricAlias");
    } else {
      rejectedFields.push("metricFamily");
      issues.push(
        issue(
          "DETERMINISTIC_METRIC_PRESERVED",
          "INFO",
          "기존 metric family가 존재해 충돌하는 LLM metric을 보조 근거로만 기록했습니다.",
          { baseMetric, llmMetric, confidence },
        ),
      );
    }
  } else if (
    confidence >= ROLE_ACCEPT_THRESHOLD &&
    semanticType === "measure" &&
    ["REFINE", "REPLACE", "KEEP"].includes(llmColumn.decision)
  ) {
    metricFamily = llmMetric;
    metricAliases.push(llmMetric);
    selectedSource = "LLM";
    acceptedFields.push("metricFamily");
  } else {
    rejectedFields.push("metricFamily");
    issues.push(
      issue(
        "LLM_METRIC_NOT_ACCEPTED",
        "INFO",
        "LLM metric family가 신뢰도 또는 semantic type 조건을 충족하지 못했습니다.",
        { llmMetric, confidence, semanticType },
      ),
    );
  }

  return {
    metricFamily,
    metricAliases: unique(metricAliases),
    selectedSource,
    acceptedFields,
    rejectedFields,
    issues,
  };
}

function resolveColumn(baseColumn, llmColumn = {}, tableId) {
  const role = resolveRole(baseColumn, llmColumn);
  const semanticType = resolvedSemanticType({
    baseColumn,
    llmColumn,
    roleAccepted: role.roleAccepted,
  });
  const metric = resolveMetric(baseColumn, llmColumn, role, semanticType);
  const normalizedMeaning =
    llmColumn.confidence >= 0.7 && normalizeText(llmColumn.normalizedMeaning)
      ? normalizeText(llmColumn.normalizedMeaning)
      : normalizeText(
          baseColumn.normalizedHeader || baseColumn.sourceHeader || "",
        );
  const defaultAggregation = acceptedDefaultAggregation({
    llmColumn,
    semanticType,
  });
  const unitSemantic =
    llmColumn.confidence >= METADATA_ACCEPT_THRESHOLD
      ? normalizeText(llmColumn.unitSemantic || "NONE") || "NONE"
      : "NONE";
  const supportedOperations = unique([
    ...asArray(baseColumn.supportedOperations),
    ...operationCapabilities({
      role: role.semanticRole,
      semanticType: semanticType === "temporal" ? "dimension" : semanticType,
      dataType: baseColumn.dataType,
    }),
    defaultAggregation !== "none" ? defaultAggregation : "",
  ]);
  const capabilities = unique([
    ...asArray(baseColumn.capabilities),
    `column_role:${role.semanticRole}`,
    ...role.roleAliases.map((item) => `column_role:${item}`),
    `semantic_type:${semanticType}`,
    metric.metricFamily ? `metric_family:${metric.metricFamily}` : "",
    ...metric.metricAliases.map((item) => `metric_family:${item}`),
    ...supportedOperations.map((item) => `operation:${item}`),
  ]);
  const acceptedFields = unique([
    ...role.acceptedFields,
    ...metric.acceptedFields,
    llmColumn.confidence >= 0.7 ? "normalizedMeaning" : "",
    llmColumn.confidence >= METADATA_ACCEPT_THRESHOLD
      ? "defaultAggregation"
      : "",
    llmColumn.confidence >= METADATA_ACCEPT_THRESHOLD ? "unitSemantic" : "",
  ]);
  const rejectedFields = unique([
    ...role.rejectedFields,
    ...metric.rejectedFields,
  ]);
  const issues = [...role.issues, ...metric.issues];
  const selectedConfidence =
    role.selectedSource === "LLM"
      ? round(clamp(llmColumn.confidence))
      : round(clamp(baseColumn.roleConfidence));

  return {
    columnId: baseColumn.columnId,
    tableId,
    sourceHeader: normalizeText(baseColumn.sourceHeader || ""),
    normalizedHeader: normalizeText(baseColumn.normalizedHeader || ""),
    dataType: baseColumn.dataType,
    semanticRole: role.semanticRole,
    roleAliases: role.roleAliases,
    semanticType,
    metricFamily: metric.metricFamily,
    metricAliases: metric.metricAliases,
    normalizedMeaning,
    defaultAggregation,
    unitSemantic,
    supportedOperations,
    capabilities,
    confidence: selectedConfidence,
    selectedSource: role.selectedSource,
    resolutionDecision: role.resolutionDecision,
    deterministic: {
      semanticRole:
        normalizeRole(baseColumn.semanticRole || "unknown") || "unknown",
      semanticType:
        normalizeRole(baseColumn.semanticType || "unknown") || "unknown",
      metricFamily: normalizeMetric(baseColumn.metricFamily || ""),
      roleConfidence: round(clamp(baseColumn.roleConfidence)),
      roleSource: normalizeText(baseColumn.roleSource || ""),
      explicitSemanticRole: normalizeText(
        baseColumn.evidence?.explicitSemanticRole || "",
      ),
    },
    llm: {
      semanticRole:
        normalizeRole(llmColumn.semanticRole || "unknown") || "unknown",
      semanticType:
        normalizeRole(llmColumn.semanticType || "unknown") || "unknown",
      metricFamily: normalizeMetric(llmColumn.metricFamily || ""),
      decision: normalizeText(llmColumn.decision || "UNKNOWN"),
      confidence: round(clamp(llmColumn.confidence)),
      evidenceCodes: unique(llmColumn.evidenceCodes),
      description: normalizeText(llmColumn.description || ""),
    },
    acceptedFields,
    rejectedFields,
    issues,
  };
}

function relationKey(relation = {}) {
  return [
    relation.leftTableId,
    relation.rightTableId,
    relation.relationType,
  ].join("::");
}

function buildResolvedRelations(deterministicProfile, llmProfile, issues) {
  const tableIds = new Set(
    asArray(deterministicProfile.tables).map((item) => item.tableId),
  );
  const columnIds = new Set(
    asArray(deterministicProfile.tables).flatMap((table) =>
      asArray(table.columns).map((column) => column.columnId),
    ),
  );
  const relations = [];
  const seen = new Set();

  for (const table of asArray(deterministicProfile.tables)) {
    if (!table.sourceTableId || !tableIds.has(table.sourceTableId)) continue;
    const relation = {
      leftTableId: table.sourceTableId,
      rightTableId: table.tableId,
      relationType: "SOURCE_DERIVATION",
      cardinality: "ONE_TO_MANY",
      leftColumnIds: [],
      rightColumnIds: [],
      confidence: 1,
      selectedSource: "DETERMINISTIC_STRUCTURE",
      evidenceCodes: ["SOURCE_TABLE_LINK"],
      description: "queryJson의 sourceTableId로 확인된 파생 테이블 관계입니다.",
    };
    const key = relationKey(relation);
    seen.add(key);
    relations.push(relation);
  }

  for (const item of asArray(llmProfile.tableRelations)) {
    if (item.confidence < RELATION_ACCEPT_THRESHOLD) {
      issues.push(
        issue(
          "LLM_RELATION_LOW_CONFIDENCE",
          "INFO",
          "신뢰도가 낮은 LLM 테이블 관계를 병합에서 제외했습니다.",
          {
            leftTableId: item.leftTableId,
            rightTableId: item.rightTableId,
            relationType: item.relationType,
            confidence: item.confidence,
          },
        ),
      );
      continue;
    }
    if (
      !tableIds.has(item.leftTableId) ||
      !tableIds.has(item.rightTableId) ||
      asArray(item.leftColumnIds).some((id) => !columnIds.has(id)) ||
      asArray(item.rightColumnIds).some((id) => !columnIds.has(id))
    ) {
      issues.push(
        issue(
          "LLM_RELATION_REFERENCE_REJECTED",
          "WARNING",
          "유효하지 않은 ID를 포함한 LLM 테이블 관계를 제외했습니다.",
          { relationType: item.relationType },
        ),
      );
      continue;
    }
    const relation = {
      leftTableId: item.leftTableId,
      rightTableId: item.rightTableId,
      relationType: item.relationType,
      cardinality: item.cardinality,
      leftColumnIds: unique(item.leftColumnIds),
      rightColumnIds: unique(item.rightColumnIds),
      confidence: round(clamp(item.confidence)),
      selectedSource: "LLM",
      evidenceCodes: unique(item.evidenceCodes),
      description: normalizeText(item.description || ""),
    };
    const key = relationKey(relation);
    if (seen.has(key)) continue;
    seen.add(key);
    relations.push(relation);
  }
  return relations;
}

function buildResolvedSemanticProfile({
  deterministicProfile,
  llmProfile,
} = {}) {
  const indexes = assertMergeInputs(deterministicProfile, llmProfile);
  const issues = [];
  const tables = asArray(deterministicProfile.tables).map((baseTable) => {
    const llmTable = indexes.llmTables.get(baseTable.tableId) || {
      tablePurpose: "UNKNOWN",
      rowGrain: "",
      confidence: 0,
      evidenceCodes: [],
      description: "",
    };
    const tablePurposeAccepted =
      llmTable.confidence >= RELATION_ACCEPT_THRESHOLD;
    const columns = asArray(baseTable.columns).map((baseColumn) =>
      resolveColumn(
        baseColumn,
        indexes.llmColumns.get(baseColumn.columnId) || {
          semanticRole: "unknown",
          semanticType: "unknown",
          metricFamily: "NONE",
          normalizedMeaning: "",
          defaultAggregation: "none",
          unitSemantic: "NONE",
          decision: "UNKNOWN",
          confidence: 0,
          evidenceCodes: [],
          description: "",
        },
        baseTable.tableId,
      ),
    );
    for (const column of columns) {
      for (const columnIssue of column.issues) {
        issues.push({
          ...columnIssue,
          details: {
            ...columnIssue.details,
            tableId: baseTable.tableId,
            columnId: column.columnId,
          },
        });
      }
    }
    if (!tablePurposeAccepted) {
      issues.push(
        issue(
          "LLM_TABLE_PURPOSE_LOW_CONFIDENCE",
          "INFO",
          "테이블 목적 신뢰도가 낮아 UNKNOWN으로 유지했습니다.",
          { tableId: baseTable.tableId, confidence: llmTable.confidence },
        ),
      );
    }
    const availableRoles = uniqueSorted(
      columns.flatMap((column) => column.roleAliases),
    );
    const metricFamilies = uniqueSorted(
      columns.flatMap((column) => column.metricAliases),
    );
    const supportedOperations = uniqueSorted(
      columns.flatMap((column) => column.supportedOperations),
    );
    const capabilities = uniqueSorted([
      ...asArray(baseTable.capabilities),
      ...columns.flatMap((column) => column.capabilities),
      ...availableRoles.map((role) => `column_role:${role}`),
      ...metricFamilies.map((family) => `metric_family:${family}`),
      ...supportedOperations.map((operation) => `operation:${operation}`),
    ]);
    return {
      tableId: baseTable.tableId,
      sourceTableId: normalizeText(baseTable.sourceTableId || ""),
      sourceSheetName: normalizeText(baseTable.sourceSheetName || ""),
      flags: { ...baseTable.flags },
      shape: { ...baseTable.shape },
      tablePurpose: tablePurposeAccepted ? llmTable.tablePurpose : "UNKNOWN",
      rowGrain: tablePurposeAccepted
        ? normalizeText(llmTable.rowGrain || "")
        : "",
      tableSemanticConfidence: tablePurposeAccepted
        ? round(clamp(llmTable.confidence))
        : 0,
      tableSemanticSource: tablePurposeAccepted ? "LLM" : "UNKNOWN",
      tableSemanticEvidenceCodes: tablePurposeAccepted
        ? unique(llmTable.evidenceCodes)
        : [],
      availableRoles,
      metricFamilies,
      supportedOperations,
      capabilities,
      columns,
    };
  });

  const tableRelations = buildResolvedRelations(
    deterministicProfile,
    llmProfile,
    issues,
  );
  const availableRoles = uniqueSorted(
    tables.flatMap((table) => table.availableRoles),
  );
  const metricFamilies = uniqueSorted(
    tables.flatMap((table) => table.metricFamilies),
  );
  const supportedOperations = uniqueSorted(
    tables.flatMap((table) => table.supportedOperations),
  );
  const availableCapabilities = uniqueSorted([
    ...asArray(deterministicProfile.availableCapabilities),
    ...tables.flatMap((table) => table.capabilities),
    `business_domain:${llmProfile.classification.primaryDomain}`,
    `dataset_intent:${llmProfile.classification.datasetIntent}`,
    ...tableRelations.map(
      (relation) => `table_relation:${relation.relationType}`,
    ),
  ]);
  const allColumns = tables.flatMap((table) => table.columns);
  const blockingIssues = issues.filter(
    (item) => item.level === "BLOCKING",
  ).length;
  const warningIssues = issues.filter(
    (item) => item.level === "WARNING",
  ).length;
  const infoIssues = issues.filter((item) => item.level === "INFO").length;
  const profile = {
    version: RESOLVED_SEMANTIC_PROFILE_VERSION,
    schemaVersion: RESOLVED_SEMANTIC_SCHEMA_VERSION,
    policyVersion: RESOLVED_SEMANTIC_POLICY_VERSION,
    source: {
      caseId: normalizeText(deterministicProfile.source?.caseId || ""),
      fileName: normalizeText(deterministicProfile.source?.fileName || ""),
      deterministicProfileVersion: deterministicProfile.version,
      deterministicProfileSha256: deterministicProfile.profileSha256,
      llmProfileVersion: llmProfile.version,
      llmProfileSha256: llmProfile.profileSha256,
      llmModel: normalizeText(llmProfile.model?.model || ""),
      llmPromptVersion: normalizeText(llmProfile.promptVersion || ""),
    },
    policy: {
      roleAcceptThreshold: ROLE_ACCEPT_THRESHOLD,
      headerReplaceThreshold: HEADER_REPLACE_THRESHOLD,
      metadataAcceptThreshold: METADATA_ACCEPT_THRESHOLD,
      relationAcceptThreshold: RELATION_ACCEPT_THRESHOLD,
      explicitRoleOverrideAllowed: false,
      structuralFieldsMutableByLlm: false,
    },
    classification: {
      primaryDomain: llmProfile.classification.primaryDomain,
      secondaryDomains: unique(llmProfile.classification.secondaryDomains),
      datasetIntent: llmProfile.classification.datasetIntent,
      confidence: round(clamp(llmProfile.classification.confidence)),
      selectedSource: "LLM",
      description: normalizeText(llmProfile.classification.description || ""),
    },
    counts: {
      totalTables: tables.length,
      totalColumns: allColumns.length,
      llmRoleAcceptedColumns: allColumns.filter(
        (column) => column.selectedSource === "LLM",
      ).length,
      explicitRolePreservedColumns: allColumns.filter(
        (column) => column.selectedSource === "ORIGINAL_EXPLICIT",
      ).length,
      llmRoleRejectedColumns: allColumns.filter((column) =>
        column.rejectedFields.includes("semanticRole"),
      ).length,
      resolvedRelationCount: tableRelations.length,
      blockingIssues,
      warningIssues,
      infoIssues,
    },
    availableRoles,
    metricFamilies,
    supportedOperations,
    availableCapabilities,
    tables,
    tableRelations,
    ambiguities: asArray(llmProfile.ambiguities).map((item) => ({
      code: item.code,
      description: normalizeText(item.description || ""),
      tableIds: unique(item.tableIds),
      columnIds: unique(item.columnIds),
      selectedSource: "LLM",
    })),
    issues,
    requiresHumanReview:
      llmProfile.requiresHumanReview === true ||
      blockingIssues > 0 ||
      warningIssues > 0 ||
      llmProfile.classification.primaryDomain === "UNKNOWN",
  };
  profile.profileSha256 = sha256({ ...profile, profileSha256: undefined });
  return profile;
}

function validateResolvedSemanticProfile(profile = {}) {
  const errors = [];
  const warnings = [];
  if (!profile || typeof profile !== "object" || Array.isArray(profile)) {
    return {
      valid: false,
      errorCount: 1,
      warningCount: 0,
      errors: [{ code: "PROFILE_NOT_OBJECT" }],
      warnings,
    };
  }
  if (profile.version !== RESOLVED_SEMANTIC_PROFILE_VERSION) {
    errors.push({ code: "PROFILE_VERSION_INVALID" });
  }
  if (profile.schemaVersion !== RESOLVED_SEMANTIC_SCHEMA_VERSION) {
    errors.push({ code: "PROFILE_SCHEMA_VERSION_INVALID" });
  }
  if (profile.policyVersion !== RESOLVED_SEMANTIC_POLICY_VERSION) {
    errors.push({ code: "PROFILE_POLICY_VERSION_INVALID" });
  }
  if (
    !/^[a-f0-9]{64}$/.test(profile.source?.deterministicProfileSha256 || "")
  ) {
    errors.push({ code: "DETERMINISTIC_PROFILE_SHA_INVALID" });
  }
  if (!/^[a-f0-9]{64}$/.test(profile.source?.llmProfileSha256 || "")) {
    errors.push({ code: "LLM_PROFILE_SHA_INVALID" });
  }
  const tableIds = new Set();
  const columnIds = new Set();
  for (const table of asArray(profile.tables)) {
    if (!table.tableId || tableIds.has(table.tableId)) {
      errors.push({
        code: "TABLE_ID_INVALID_OR_DUPLICATED",
        tableId: table.tableId,
      });
    }
    tableIds.add(table.tableId);
    for (const column of asArray(table.columns)) {
      if (
        !column.columnId ||
        columnIds.has(column.columnId) ||
        column.tableId !== table.tableId
      ) {
        errors.push({
          code: "COLUMN_ID_INVALID_OR_DUPLICATED",
          columnId: column.columnId,
        });
      }
      columnIds.add(column.columnId);
      if (!column.semanticRole || !column.semanticType) {
        errors.push({
          code: "COLUMN_RESOLUTION_REQUIRED",
          columnId: column.columnId,
        });
      }
      if (
        column.selectedSource === "LLM" &&
        column.confidence < ROLE_ACCEPT_THRESHOLD
      ) {
        errors.push({
          code: "LLM_ROLE_ACCEPTED_BELOW_THRESHOLD",
          columnId: column.columnId,
        });
      }
    }
  }
  for (const relation of asArray(profile.tableRelations)) {
    if (
      !tableIds.has(relation.leftTableId) ||
      !tableIds.has(relation.rightTableId)
    ) {
      errors.push({ code: "RELATION_TABLE_REFERENCE_INVALID" });
    }
    if (asArray(relation.leftColumnIds).some((id) => !columnIds.has(id))) {
      errors.push({ code: "RELATION_LEFT_COLUMN_REFERENCE_INVALID" });
    }
    if (asArray(relation.rightColumnIds).some((id) => !columnIds.has(id))) {
      errors.push({ code: "RELATION_RIGHT_COLUMN_REFERENCE_INVALID" });
    }
  }
  if (profile.counts?.totalTables !== tableIds.size) {
    errors.push({ code: "TABLE_COUNT_MISMATCH" });
  }
  if (profile.counts?.totalColumns !== columnIds.size) {
    errors.push({ code: "COLUMN_COUNT_MISMATCH" });
  }
  if (!/^[a-f0-9]{64}$/.test(profile.profileSha256 || "")) {
    errors.push({ code: "PROFILE_SHA256_INVALID" });
  } else {
    const expected = sha256({ ...profile, profileSha256: undefined });
    if (expected !== profile.profileSha256) {
      errors.push({ code: "PROFILE_SHA256_MISMATCH" });
    }
  }
  if (profile.counts?.warningIssues > 0) {
    warnings.push({ code: "PROFILE_MERGE_WARNINGS_PRESENT" });
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
  RESOLVED_SEMANTIC_PROFILE_VERSION,
  RESOLVED_SEMANTIC_POLICY_VERSION,
  RESOLVED_SEMANTIC_SCHEMA_VERSION,
  ROLE_ACCEPT_THRESHOLD,
  HEADER_REPLACE_THRESHOLD,
  METADATA_ACCEPT_THRESHOLD,
  RELATION_ACCEPT_THRESHOLD,
  isGenericRole,
  sameRoleFamily,
  sameMetricFamily,
  assertMergeInputs,
  resolveColumn,
  buildResolvedSemanticProfile,
  validateResolvedSemanticProfile,
};
