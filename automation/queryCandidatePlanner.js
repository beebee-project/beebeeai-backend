"use strict";

const crypto = require("crypto");
const { normalizeText, sha256 } = require("./queryCandidateObservation");

const QUERY_CANDIDATE_PLANNER_INPUT_VERSION = "query_candidate_planner_input_v1";
const QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION =
  "query_candidate_planner_model_output_v1";
const QUERY_CANDIDATE_PLANNER_RESOLUTION_VERSION =
  "query_candidate_planner_resolution_v1";
const QUERY_CANDIDATE_PLANNER_ITEM_VERSION = "query_candidate_planner_item_v1";
const QUERY_CANDIDATE_PLANNER_POLICY_VERSION =
  "conditional_llm_candidate_planner_policy_v1";
const QUERY_CANDIDATE_PLANNER_CACHE_VERSION = "query_candidate_planner_cache_v1";
const DEFAULT_MODEL = "gpt-5.6-terra";
const DEFAULT_REASONING_EFFORT = "low";
const MAX_PROPOSALS = 3;
const MIN_PROPOSAL_CONFIDENCE = 0.72;
const COMPLEX_ROW_THRESHOLD = 8;

const MODEL_PRICING_USD_PER_MILLION = Object.freeze({
  "gpt-5.6-terra": Object.freeze({ input: 2, cachedInput: 0.2, output: 12 }),
});

const INVOCATION_STATUS = Object.freeze([
  "SKIPPED",
  "REQUIRED_NOT_RUN",
  "CALLED",
  "CACHE_HIT",
  "FAILED_SAFE",
]);

const PLANNER_DECISION = Object.freeze([
  "NOT_REQUIRED",
  "CALL_REQUIRED",
  "NO_ADDITION",
  "PROPOSED",
  "FAILED_SAFE",
]);

const PROPOSAL_DISPOSITION = Object.freeze([
  "ACCEPTED_FOR_REVALIDATION",
  "REJECTED",
]);

const ALLOWED_OPERATIONS = Object.freeze({
  countrows: Object.freeze({ operation: "count_rows", kinds: [] }),
  categorycount: Object.freeze({ operation: "category_count", kinds: ["group"] }),
  groupsum: Object.freeze({ operation: "group_sum", kinds: ["group", "measure"] }),
  groupavg: Object.freeze({ operation: "group_avg", kinds: ["group", "measure"] }),
  topbottom: Object.freeze({ operation: "top_bottom", kinds: ["group", "measure"] }),
  timesum: Object.freeze({ operation: "time_sum", kinds: ["period", "measure"] }),
  timeavg: Object.freeze({ operation: "time_avg", kinds: ["period", "measure"] }),
  timecount: Object.freeze({ operation: "time_count", kinds: ["period"] }),
  cumulativesum: Object.freeze({
    operation: "cumulative_sum",
    kinds: ["period", "measure"],
  }),
  crosssum: Object.freeze({
    operation: "cross_sum",
    kinds: ["dimension", "dimension", "measure"],
  }),
  crosscount: Object.freeze({
    operation: "cross_count",
    kinds: ["dimension", "dimension"],
  }),
});

const MODEL_EVIDENCE_CODES = Object.freeze([
  "NO_READY_RECOVERY",
  "LOW_READY_COVERAGE",
  "MISSING_OVERVIEW",
  "MISSING_GROUP_AGGREGATION",
  "MISSING_TIME_SERIES",
  "MISSING_RANKING",
  "MISSING_CROSS_TAB",
  "DEFERRED_CANDIDATE_RECOVERY",
  "REVIEW_CANDIDATE_RECOVERY",
]);

const MODEL_OUTPUT_SCHEMA = Object.freeze({
  type: "object",
  additionalProperties: false,
  required: ["version", "decision", "summary", "proposals"],
  properties: {
    version: { type: "string", const: QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION },
    decision: { type: "string", enum: ["NO_ADDITION", "PROPOSE"] },
    summary: { type: "string", maxLength: 400 },
    proposals: {
      type: "array",
      maxItems: MAX_PROPOSALS,
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "proposalKey",
          "title",
          "purpose",
          "operation",
          "sourceTableIds",
          "operandBindings",
          "outputType",
          "confidence",
          "evidenceCodes",
        ],
        properties: {
          proposalKey: { type: "string", maxLength: 80 },
          title: { type: "string", maxLength: 120 },
          purpose: { type: "string", maxLength: 240 },
          operation: {
            type: "string",
            enum: Object.values(ALLOWED_OPERATIONS).map((item) => item.operation),
          },
          sourceTableIds: {
            type: "array",
            minItems: 1,
            maxItems: 1,
            items: { type: "string" },
          },
          operandBindings: {
            type: "array",
            maxItems: 3,
            items: {
              type: "object",
              additionalProperties: false,
              required: ["kind", "columnId"],
              properties: {
                kind: {
                  type: "string",
                  enum: ["group", "measure", "period", "dimension"],
                },
                columnId: { type: "string" },
              },
            },
          },
          outputType: { type: "string", const: "summarysheet" },
          confidence: { type: "number", minimum: 0, maximum: 1 },
          evidenceCodes: {
            type: "array",
            maxItems: 6,
            items: { type: "string", enum: MODEL_EVIDENCE_CODES },
          },
        },
      },
    },
  },
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

function sortedUnique(values = []) {
  return unique(values).sort((a, b) => a.localeCompare(b, "ko"));
}

function normalizeLoose(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[^가-힣a-z0-9]+/gu, "");
}

function round(value, digits = 6) {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  return Number(number.toFixed(digits));
}

function clamp(value, minimum = 0, maximum = 1) {
  const number = Number(value);
  if (!Number.isFinite(number)) return minimum;
  return Math.min(maximum, Math.max(minimum, number));
}

function issue(path, code, message, details = {}) {
  return { path, code, message, details };
}

function operationDefinition(operation = "") {
  return ALLOWED_OPERATIONS[normalizeLoose(operation)] || null;
}

function operationCategory(operation = "") {
  const value = normalizeLoose(operation);
  if (["countrows", "categorycount", "timecount"].includes(value)) return "OVERVIEW";
  if (["groupsum", "groupavg"].includes(value)) return "GROUP_AGGREGATION";
  if (["timesum", "timeavg", "cumulativesum"].includes(value)) return "TIME_SERIES";
  if (value === "topbottom") return "RANKING";
  if (["crosssum", "crosscount"].includes(value)) return "CROSS_TAB";
  return "OTHER";
}

function tableIsPhysicalAndEligible(table = {}) {
  const flags = table.flags || {};
  return flags.analysisEligible === true && flags.virtual !== true;
}

function columnRoleTokens(column = {}) {
  return new Set(
    sortedUnique([
      column.semanticRole,
      column.semanticType,
      column.dataType,
      ...asArray(column.roleAliases),
    ]).map(normalizeLoose),
  );
}

function isPeriodColumn(column = {}) {
  const tokens = columnRoleTokens(column);
  return ["period", "date", "datetime", "time", "yearmonth", "month", "year"].some(
    (token) => tokens.has(token),
  );
}

function isMeasureColumn(column = {}) {
  const tokens = columnRoleTokens(column);
  if (["measure", "metric", "amount", "revenue", "sales", "budget", "cost", "expense", "score", "rating", "value", "quantity"].some((token) => tokens.has(token))) {
    return true;
  }
  return normalizeLoose(column.semanticType) === "measure" ||
    (normalizeLoose(column.dataType) === "number" && normalizeText(column.metricFamily));
}

function isGroupColumn(column = {}) {
  if (isPeriodColumn(column) || isMeasureColumn(column)) return false;
  const tokens = columnRoleTokens(column);
  return [
    "group",
    "dimension",
    "category",
    "status",
    "product",
    "customer",
    "entity",
    "organization",
    "department",
    "person",
  ].some((token) => tokens.has(token));
}

function sanitizedSemanticTables(semanticProfile = {}) {
  return asArray(semanticProfile.tables)
    .filter(tableIsPhysicalAndEligible)
    .slice(0, 12)
    .map((table) => ({
      tableId: normalizeText(table.tableId || ""),
      rowCount: Math.max(0, Number(table.shape?.rowCount || 0)),
      columnCount: Math.max(0, Number(table.shape?.columnCount || asArray(table.columns).length)),
      primary: table.flags?.primary === true,
      columns: asArray(table.columns)
        .slice(0, 80)
        .map((column) => ({
          columnId: normalizeText(column.columnId || ""),
          header: normalizeText(column.normalizedHeader || column.sourceHeader || "").slice(0, 120),
          dataType: normalizeText(column.dataType || "unknown"),
          semanticRole: normalizeText(column.semanticRole || "unknown"),
          semanticType: normalizeText(column.semanticType || "unknown"),
          metricFamily: normalizeText(column.metricFamily || ""),
          roleAliases: sortedUnique(column.roleAliases).slice(0, 8),
          uniqueRatio: Number.isFinite(Number(column.stats?.uniqueRatio))
            ? round(Number(column.stats.uniqueRatio))
            : null,
          nonEmptyRatio: Number.isFinite(Number(column.stats?.nonEmptyRatio))
            ? round(Number(column.stats.nonEmptyRatio))
            : null,
        })),
    }));
}

function deriveSemanticOpportunities(tables = []) {
  const categories = new Set();
  const details = [];
  for (const table of asArray(tables)) {
    const groups = asArray(table.columns).filter(isGroupColumn);
    const periods = asArray(table.columns).filter(isPeriodColumn);
    const measures = asArray(table.columns).filter(isMeasureColumn);
    const tableCategories = [];
    if (groups.length) tableCategories.push("OVERVIEW");
    if (groups.length && measures.length) {
      tableCategories.push("GROUP_AGGREGATION", "RANKING");
    }
    if (periods.length && measures.length) tableCategories.push("TIME_SERIES");
    if (groups.length >= 2 && measures.length) tableCategories.push("CROSS_TAB");
    for (const category of tableCategories) categories.add(category);
    details.push({
      tableId: normalizeText(table.tableId || ""),
      groupColumnIds: groups.map((item) => item.columnId),
      periodColumnIds: periods.map((item) => item.columnId),
      measureColumnIds: measures.map((item) => item.columnId),
      opportunityCategories: sortedUnique(tableCategories),
    });
  }
  return { categories: sortedUnique([...categories]), details };
}

function rankedReadyItems(rankingResolution = {}) {
  return asArray(rankingResolution.candidates)
    .filter((item) => item.rankingDisposition === "RANKED")
    .sort((left, right) => Number(left.rank || 0) - Number(right.rank || 0));
}

function unresolvedCandidateSummaries(candidateResolution = {}, feasibilityResolution = {}) {
  const result = [];
  for (const item of asArray(candidateResolution.candidates)) {
    if (item.result !== "STILL_DEFERRED") continue;
    result.push({
      candidateId: normalizeText(item.candidateId || ""),
      stage: "RESOLVER",
      status: "STILL_DEFERRED",
      operation: normalizeText(item.checks?.operandBinding?.operation || item.recipeId || ""),
      sourceTableIds: sortedUnique(item.checks?.sourceScope?.matchedRootTableIds),
      reasonCodes: sortedUnique(asArray(item.reasons).map((reason) => reason.code)).slice(0, 8),
    });
  }
  for (const item of asArray(feasibilityResolution.candidates)) {
    if (item.feasibilityStatus !== "REVIEW") continue;
    result.push({
      candidateId: normalizeText(item.candidateId || ""),
      stage: "FEASIBILITY",
      status: "REVIEW",
      operation: normalizeText(item.executionPlan?.operation || ""),
      sourceTableIds: sortedUnique(item.executionPlan?.sourceTableIds),
      reasonCodes: sortedUnique(asArray(item.reasons).map((reason) => reason.code)).slice(0, 8),
    });
  }
  return result.slice(0, 30);
}

function buildPlannerInput({
  semanticProfile = {},
  resolvedSemanticProfile = {},
  candidateResolution = {},
  candidateFeasibilityResolution = {},
  candidateRankingResolution = {},
} = {}) {
  const tables = sanitizedSemanticTables(semanticProfile);
  const opportunities = deriveSemanticOpportunities(tables);
  const ready = rankedReadyItems(candidateRankingResolution);
  const coveredCategories = sortedUnique(ready.map((item) => item.operationCategory));
  const missingCategories = opportunities.categories.filter(
    (category) => !coveredCategories.includes(category),
  );
  const unresolved = unresolvedCandidateSummaries(
    candidateResolution,
    candidateFeasibilityResolution,
  );
  const classification =
    resolvedSemanticProfile.classification ||
    resolvedSemanticProfile.businessDomain ||
    semanticProfile.classification ||
    {};
  const input = {
    version: QUERY_CANDIDATE_PLANNER_INPUT_VERSION,
    source: {
      caseId: normalizeText(
        semanticProfile.source?.caseId || candidateResolution.source?.caseId || "",
      ),
      semanticProfileVersion: normalizeText(semanticProfile.version || ""),
      semanticProfileSha256: normalizeText(
        semanticProfile.profileSha256 || semanticProfile.semanticProfileSha256 || sha256(semanticProfile),
      ),
      candidateResolutionSha256: normalizeText(candidateResolution.resolutionSha256 || ""),
      candidateFeasibilityResolutionSha256: normalizeText(
        candidateFeasibilityResolution.feasibilityResolutionSha256 || "",
      ),
      candidateRankingResolutionSha256: normalizeText(
        candidateRankingResolution.rankingResolutionSha256 || "",
      ),
    },
    privacy: {
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      originalFileIncluded: false,
      fileNameIncluded: false,
    },
    classification: {
      primaryDomain: normalizeText(
        classification.primaryDomain || resolvedSemanticProfile.primaryDomain || "UNKNOWN",
      ),
      datasetIntent: normalizeText(
        classification.datasetIntent || resolvedSemanticProfile.datasetIntent || "UNKNOWN",
      ),
      confidence: round(
        clamp(
          classification.confidence ??
            resolvedSemanticProfile.confidence ??
            resolvedSemanticProfile.businessDomain?.confidence ??
            0,
        ),
      ),
    },
    deterministicCoverage: {
      readyCount: ready.length,
      recommendedCount: asArray(candidateRankingResolution.recommendedCandidateIds).length,
      coveredCategories,
      opportunityCategories: opportunities.categories,
      missingCategories,
      unresolvedCount: unresolved.length,
      reviewCount: asArray(candidateFeasibilityResolution.candidates).filter(
        (item) => item.feasibilityStatus === "REVIEW",
      ).length,
      deferredCount: asArray(candidateResolution.candidates).filter(
        (item) => item.result === "STILL_DEFERRED",
      ).length,
    },
    tables,
    opportunityDetails: opportunities.details,
    rankedCandidates: ready.slice(0, 12).map((item) => ({
      candidateId: normalizeText(item.candidateId || ""),
      rank: Number(item.rank || 0),
      operation: normalizeText(item.operation || ""),
      operationCategory: normalizeText(item.operationCategory || ""),
      sourceTableIds: sortedUnique(item.sourceTableIds),
      recipeId: normalizeText(item.recipeId || ""),
      templateId: normalizeText(item.templateId || ""),
    })),
    unresolvedCandidates: unresolved,
    allowedOperations: Object.values(ALLOWED_OPERATIONS).map((definition) => ({
      operation: definition.operation,
      requiredOperandKinds: [...definition.kinds],
      outputType: "summarysheet",
      singlePhysicalTableOnly: true,
    })),
    limits: {
      maxProposals: MAX_PROPOSALS,
      minimumConfidence: MIN_PROPOSAL_CONFIDENCE,
    },
  };
  input.inputSha256 = sha256({ ...input, inputSha256: undefined });
  return input;
}

function evaluateConditionalPlannerTrigger(input = {}) {
  const eligibleTables = asArray(input.tables);
  const coverage = input.deterministicCoverage || {};
  const readyCount = Number(coverage.readyCount || 0);
  const maxRowCount = Math.max(0, ...eligibleTables.map((table) => Number(table.rowCount || 0)));
  const totalColumns = eligibleTables.reduce(
    (total, table) => total + asArray(table.columns).length,
    0,
  );
  const opportunityCount = asArray(coverage.opportunityCategories).length;
  const missingCount = asArray(coverage.missingCategories).length;
  const unresolvedCount = Number(coverage.unresolvedCount || 0);
  const complexEnough =
    maxRowCount >= COMPLEX_ROW_THRESHOLD && totalColumns >= 3 && opportunityCount > 0;

  if (!eligibleTables.length) {
    return {
      required: false,
      reasonCode: "NO_ANALYSIS_ELIGIBLE_TABLE",
      reason: "분석 가능한 물리 테이블이 없어 Planner를 호출하지 않습니다.",
      metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
    };
  }
  if (!opportunityCount) {
    return {
      required: false,
      reasonCode: "NO_SUPPORTED_SEMANTIC_OPPORTUNITY",
      reason: "지원 operation으로 변환할 의미 기회가 없어 Planner를 호출하지 않습니다.",
      metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
    };
  }
  if (readyCount >= 3) {
    return {
      required: false,
      reasonCode: "ADEQUATE_DETERMINISTIC_COVERAGE",
      reason: "READY 후보가 3개 이상이므로 결정론적 결과를 우선 사용합니다.",
      metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
    };
  }
  if (readyCount === 0) {
    return {
      required: true,
      reasonCode: "NO_READY_CANDIDATE",
      reason: "분석 가능한 데이터에 READY 후보가 없어 조건부 Planner 복구가 필요합니다.",
      metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
    };
  }
  if (!complexEnough) {
    return {
      required: false,
      reasonCode: "SIMPLE_DATASET_WITH_EXISTING_READY",
      reason: "소규모 데이터에 실행 가능한 후보가 이미 있어 추가 LLM 호출을 생략합니다.",
      metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
    };
  }
  if (missingCount >= 2) {
    return {
      required: true,
      reasonCode: "LOW_DETERMINISTIC_CATEGORY_COVERAGE",
      reason: "복합 데이터에서 READY 후보가 적고 의미 기회 범주가 충분히 덮이지 않았습니다.",
      metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
    };
  }
  if (unresolvedCount > 0 && missingCount > 0) {
    return {
      required: true,
      reasonCode: "UNRESOLVED_COVERAGE_GAP",
      reason: "미해결 후보와 의미 기회 공백이 함께 존재해 조건부 Planner 검토가 필요합니다.",
      metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
    };
  }
  return {
    required: false,
    reasonCode: "EXISTING_READY_SUFFICIENT",
    reason: "현재 READY 후보가 데이터 복잡도와 의미 기회를 충분히 대표합니다.",
    metrics: { readyCount, maxRowCount, totalColumns, opportunityCount, missingCount, unresolvedCount },
  };
}

function modelValidationIssue(code, path, message) {
  return { code, path, message };
}

function validateCandidatePlannerModelOutput(output = {}) {
  const errors = [];
  if (!output || typeof output !== "object" || Array.isArray(output)) {
    return {
      valid: false,
      errorCount: 1,
      errors: [modelValidationIssue("MODEL_OUTPUT_NOT_OBJECT", "", "모델 출력은 객체여야 합니다.")],
    };
  }
  if (output.version !== QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION) {
    errors.push(modelValidationIssue("MODEL_OUTPUT_VERSION_INVALID", "version", "모델 출력 version이 유효하지 않습니다."));
  }
  if (!["NO_ADDITION", "PROPOSE"].includes(output.decision)) {
    errors.push(modelValidationIssue("MODEL_DECISION_INVALID", "decision", "decision이 유효하지 않습니다."));
  }
  if (!Array.isArray(output.proposals)) {
    errors.push(modelValidationIssue("MODEL_PROPOSALS_NOT_ARRAY", "proposals", "proposals는 배열이어야 합니다."));
  }
  const proposals = asArray(output.proposals);
  if (proposals.length > MAX_PROPOSALS) {
    errors.push(modelValidationIssue("MODEL_PROPOSALS_LIMIT_EXCEEDED", "proposals", "제안 개수 제한을 초과했습니다."));
  }
  if (output.decision === "NO_ADDITION" && proposals.length) {
    errors.push(modelValidationIssue("NO_ADDITION_HAS_PROPOSALS", "proposals", "NO_ADDITION은 proposals가 비어 있어야 합니다."));
  }
  if (output.decision === "PROPOSE" && !proposals.length) {
    errors.push(modelValidationIssue("PROPOSE_HAS_NO_PROPOSALS", "proposals", "PROPOSE는 하나 이상의 제안이 필요합니다."));
  }
  const keys = new Set();
  proposals.forEach((proposal, index) => {
    const path = `proposals[${index}]`;
    const key = normalizeText(proposal.proposalKey || "");
    if (!key) errors.push(modelValidationIssue("PROPOSAL_KEY_REQUIRED", `${path}.proposalKey`, "proposalKey가 필요합니다."));
    if (keys.has(key)) errors.push(modelValidationIssue("PROPOSAL_KEY_DUPLICATED", `${path}.proposalKey`, "proposalKey가 중복됩니다."));
    keys.add(key);
    if (!operationDefinition(proposal.operation)) {
      errors.push(modelValidationIssue("PROPOSAL_OPERATION_INVALID", `${path}.operation`, "지원하지 않는 operation입니다."));
    }
    if (!Array.isArray(proposal.sourceTableIds) || proposal.sourceTableIds.length !== 1) {
      errors.push(modelValidationIssue("PROPOSAL_SOURCE_COUNT_INVALID", `${path}.sourceTableIds`, "sourceTableIds는 정확히 한 개여야 합니다."));
    }
    if (!Array.isArray(proposal.operandBindings)) {
      errors.push(modelValidationIssue("PROPOSAL_OPERANDS_NOT_ARRAY", `${path}.operandBindings`, "operandBindings는 배열이어야 합니다."));
    }
    if (normalizeText(proposal.outputType) !== "summarysheet") {
      errors.push(modelValidationIssue("PROPOSAL_OUTPUT_INVALID", `${path}.outputType`, "outputType은 summarysheet여야 합니다."));
    }
    if (!Number.isFinite(Number(proposal.confidence)) || Number(proposal.confidence) < 0 || Number(proposal.confidence) > 1) {
      errors.push(modelValidationIssue("PROPOSAL_CONFIDENCE_INVALID", `${path}.confidence`, "confidence는 0~1이어야 합니다."));
    }
    for (const code of asArray(proposal.evidenceCodes)) {
      if (!MODEL_EVIDENCE_CODES.includes(code)) {
        errors.push(modelValidationIssue("PROPOSAL_EVIDENCE_CODE_INVALID", `${path}.evidenceCodes`, "evidence code가 유효하지 않습니다."));
      }
    }
  });
  return { valid: errors.length === 0, errorCount: errors.length, errors };
}

function tableAndColumnMaps(input = {}) {
  const tables = new Map();
  const columns = new Map();
  for (const table of asArray(input.tables)) {
    const tableId = normalizeText(table.tableId || "");
    if (!tableId) continue;
    tables.set(tableId, table);
    for (const column of asArray(table.columns)) {
      const columnId = normalizeText(column.columnId || "");
      if (columnId) columns.set(columnId, { ...column, tableId });
    }
  }
  return { tables, columns };
}

function normalizedBindingList(bindings = []) {
  return asArray(bindings)
    .map((item) => ({
      kind: normalizeLoose(item.kind || ""),
      columnId: normalizeText(item.columnId || ""),
    }))
    .filter((item) => item.kind && item.columnId)
    .sort((left, right) =>
      `${left.kind}|${left.columnId}`.localeCompare(`${right.kind}|${right.columnId}`, "ko"),
    );
}

function proposalSignature({ operation, sourceTableIds, operandBindings } = {}) {
  const definition = operationDefinition(operation);
  return sha256({
    operation: definition?.operation || normalizeLoose(operation),
    sourceTableIds: sortedUnique(sourceTableIds),
    operandBindings: normalizedBindingList(operandBindings),
    outputType: "summarysheet",
  });
}

function existingCandidateSignatures({
  candidateFeasibilityResolution = {},
  candidateRankingResolution = {},
} = {}) {
  const signatures = new Set();
  for (const item of asArray(candidateFeasibilityResolution.candidates)) {
    const plan = item.executionPlan;
    if (!plan || !plan.operation) continue;
    const bindings = [
      ...asArray(plan.operandBindings).map((binding) => ({
        kind: binding.kind || binding.role,
        columnId: binding.columnId || asArray(binding.columnIds)[0],
      })),
      ...asArray(plan.requiredRoleBindings).flatMap((binding) =>
        asArray(binding.columnIds).map((columnId) => ({
          kind: binding.role,
          columnId,
        })),
      ),
    ];
    signatures.add(
      proposalSignature({
        operation: plan.operation,
        sourceTableIds: plan.sourceTableIds,
        operandBindings: bindings,
      }),
    );
  }
  for (const item of rankedReadyItems(candidateRankingResolution)) {
    if (!item.operation) continue;
    signatures.add(
      proposalSignature({
        operation: item.operation,
        sourceTableIds: item.sourceTableIds,
        operandBindings: [],
      }),
    );
  }
  return signatures;
}

function requiredKindCounts(kinds = []) {
  const counts = {};
  for (const kind of asArray(kinds)) counts[kind] = (counts[kind] || 0) + 1;
  return counts;
}

function validateAndNormalizeProposal(proposal = {}, index = 0, context = {}) {
  const rejectionCodes = [];
  const definition = operationDefinition(proposal.operation);
  if (!definition) rejectionCodes.push("UNSUPPORTED_OPERATION");
  const sourceTableIds = sortedUnique(proposal.sourceTableIds);
  if (sourceTableIds.length !== 1 || !context.maps.tables.has(sourceTableIds[0])) {
    rejectionCodes.push("SOURCE_TABLE_INVALID");
  }
  const bindings = normalizedBindingList(proposal.operandBindings);
  const actualCounts = requiredKindCounts(bindings.map((item) => item.kind));
  const expectedCounts = requiredKindCounts(definition?.kinds || []);
  const kindKeys = sortedUnique([...Object.keys(actualCounts), ...Object.keys(expectedCounts)]);
  if (kindKeys.some((kind) => Number(actualCounts[kind] || 0) !== Number(expectedCounts[kind] || 0))) {
    rejectionCodes.push("OPERAND_KIND_CONTRACT_MISMATCH");
  }
  const sourceTableId = sourceTableIds[0] || "";
  for (const binding of bindings) {
    const column = context.maps.columns.get(binding.columnId);
    if (!column || column.tableId !== sourceTableId) {
      rejectionCodes.push("OPERAND_COLUMN_OUT_OF_SCOPE");
    }
  }
  if (new Set(bindings.map((item) => item.columnId)).size !== bindings.length) {
    rejectionCodes.push("OPERAND_COLUMN_DUPLICATED");
  }
  if (normalizeText(proposal.outputType) !== "summarysheet") {
    rejectionCodes.push("OUTPUT_TYPE_UNSUPPORTED");
  }
  if (Number(proposal.confidence) < MIN_PROPOSAL_CONFIDENCE) {
    rejectionCodes.push("CONFIDENCE_BELOW_THRESHOLD");
  }
  const signature = proposalSignature({
    operation: definition?.operation || proposal.operation,
    sourceTableIds,
    operandBindings: bindings,
  });
  if (context.existingSignatures.has(signature) || context.acceptedSignatures.has(signature)) {
    rejectionCodes.push("DUPLICATE_EXISTING_CANDIDATE_SIGNATURE");
  }
  const candidateId = `llmplan_${normalizeLoose(definition?.operation || proposal.operation)}_${signature.slice(0, 12)}`;
  const item = {
    version: QUERY_CANDIDATE_PLANNER_ITEM_VERSION,
    proposalIndex: index,
    proposalKey: normalizeText(proposal.proposalKey || ""),
    candidateId,
    title: normalizeText(proposal.title || "").slice(0, 120),
    purpose: normalizeText(proposal.purpose || "").slice(0, 240),
    operation: definition?.operation || normalizeText(proposal.operation || ""),
    operationCategory: operationCategory(definition?.operation || proposal.operation),
    sourceTableIds,
    operandBindings: bindings,
    outputType: "summarysheet",
    confidence: round(clamp(proposal.confidence)),
    evidenceCodes: sortedUnique(proposal.evidenceCodes).filter((code) =>
      MODEL_EVIDENCE_CODES.includes(code),
    ),
    disposition: rejectionCodes.length
      ? "REJECTED"
      : "ACCEPTED_FOR_REVALIDATION",
    rejectionCodes: sortedUnique(rejectionCodes),
    requiresResolverReentry: rejectionCodes.length === 0,
    readyStatusAssigned: false,
    sourceCandidateMutation: false,
  };
  item.plannerItemSha256 = sha256({ ...item, plannerItemSha256: undefined });
  if (!rejectionCodes.length) context.acceptedSignatures.add(signature);
  return item;
}

function normalizeUsage(usage = {}) {
  const inputDetails = usage.input_tokens_details || usage.inputTokensDetails || {};
  return {
    inputTokens: Number(usage.input_tokens ?? usage.inputTokens) || 0,
    cachedInputTokens:
      Number(
        inputDetails.cached_tokens ??
          inputDetails.cachedTokens ??
          usage.cached_input_tokens ??
          usage.cachedInputTokens,
      ) || 0,
    outputTokens: Number(usage.output_tokens ?? usage.outputTokens) || 0,
    totalTokens: Number(usage.total_tokens ?? usage.totalTokens) || 0,
  };
}

function estimateCostUsd(usage = {}, model = DEFAULT_MODEL, pricingOverride) {
  const normalized = normalizeUsage(usage);
  const pricing = pricingOverride || MODEL_PRICING_USD_PER_MILLION[model] || null;
  if (!pricing) return null;
  const cached = Math.min(normalized.inputTokens, normalized.cachedInputTokens);
  const uncached = Math.max(0, normalized.inputTokens - cached);
  return round(
    (uncached * Number(pricing.input || 0) +
      cached * Number(pricing.cachedInput || 0) +
      normalized.outputTokens * Number(pricing.output || 0)) /
      1_000_000,
    8,
  );
}

function buildCandidatePlannerCacheKey({
  tenantId,
  input,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  cacheSecret,
  promptVersion = "query_candidate_planner_prompt_v1",
} = {}) {
  const tenant = normalizeText(tenantId);
  if (!tenant) throw new Error("tenantId가 필요합니다.");
  if (!cacheSecret) throw new Error("cacheSecret이 필요합니다.");
  const identity = {
    version: QUERY_CANDIDATE_PLANNER_CACHE_VERSION,
    tenantId: tenant,
    inputSha256: normalizeText(input?.inputSha256 || sha256(input || {})),
    model,
    reasoningEffort,
    promptVersion,
    policyVersion: QUERY_CANDIDATE_PLANNER_POLICY_VERSION,
  };
  return crypto
    .createHmac("sha256", Buffer.from(String(cacheSecret)))
    .update(JSON.stringify(identity))
    .digest("hex");
}

function buildPlannerResolutionBase({ input, trigger, source = {} } = {}) {
  return {
    version: QUERY_CANDIDATE_PLANNER_RESOLUTION_VERSION,
    itemVersion: QUERY_CANDIDATE_PLANNER_ITEM_VERSION,
    policy: {
      version: QUERY_CANDIDATE_PLANNER_POLICY_VERSION,
      conditionalInvocationOnly: true,
      deterministicCoverageCheckedFirst: true,
      adequateReadyCountSkipsLlm: 3,
      simpleDatasetWithReadySkipsLlm: true,
      maxProposals: MAX_PROPOSALS,
      minimumProposalConfidence: MIN_PROPOSAL_CONFIDENCE,
      supportedOperationsOnly: true,
      knownTableAndColumnIdsOnly: true,
      singlePhysicalTableOnly: true,
      duplicateSignatureRejected: true,
      acceptedProposalsRequireResolverReentry: true,
      readyStatusAssigned: false,
      sourceCandidatesAreNotRemovedOrMutated: true,
      productionRouteChanged: false,
      plaintextCacheAllowed: false,
    },
    source: {
      caseId: normalizeText(source.caseId || input?.source?.caseId || ""),
      candidateResolutionSha256: normalizeText(input?.source?.candidateResolutionSha256 || ""),
      candidateFeasibilityResolutionSha256: normalizeText(
        input?.source?.candidateFeasibilityResolutionSha256 || "",
      ),
      candidateRankingResolutionSha256: normalizeText(
        input?.source?.candidateRankingResolutionSha256 || "",
      ),
      semanticProfileSha256: normalizeText(input?.source?.semanticProfileSha256 || ""),
      inputSha256: normalizeText(input?.inputSha256 || ""),
    },
    privacy: {
      rawRowsSent: false,
      sampleValuesSent: false,
      originalFileSent: false,
      fileNameSent: false,
      includedTableCount: asArray(input?.tables).length,
      includedColumnCount: asArray(input?.tables).reduce(
        (total, table) => total + asArray(table.columns).length,
        0,
      ),
    },
    trigger,
  };
}

function finalizeResolution(document) {
  document.plannerResolutionSha256 = sha256({
    ...document,
    plannerResolutionSha256: undefined,
  });
  return document;
}

async function runConditionalCandidatePlanner({
  semanticProfile,
  resolvedSemanticProfile,
  candidateResolution,
  candidateFeasibilityResolution,
  candidateRankingResolution,
  provider,
  cache,
  tenantId,
  cacheSecret,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  pricing,
} = {}) {
  const input = buildPlannerInput({
    semanticProfile,
    resolvedSemanticProfile,
    candidateResolution,
    candidateFeasibilityResolution,
    candidateRankingResolution,
  });
  const trigger = evaluateConditionalPlannerTrigger(input);
  const base = buildPlannerResolutionBase({ input, trigger });
  if (!trigger.required) {
    return finalizeResolution({
      ...base,
      invocation: {
        status: "SKIPPED",
        providerCallCount: 0,
        cacheHit: false,
        model: "",
        reasoningEffort: "",
        responseId: "",
        failureCode: "",
      },
      decision: "NOT_REQUIRED",
      summary: trigger.reason,
      usage: { ...normalizeUsage({}), estimatedCostUsd: 0 },
      counts: { proposed: 0, accepted: 0, rejected: 0 },
      acceptedCandidateIds: [],
      rejectedProposalKeys: [],
      proposals: [],
      integrity: {
        sourceCandidatesPreserved: true,
        providerSkipped: true,
        acceptedProposalsRequireResolverReentry: true,
      },
    });
  }
  if (!provider || typeof provider.plan !== "function") {
    return finalizeResolution({
      ...base,
      invocation: {
        status: "REQUIRED_NOT_RUN",
        providerCallCount: 0,
        cacheHit: false,
        model,
        reasoningEffort,
        responseId: "",
        failureCode: "PROVIDER_NOT_CONFIGURED",
      },
      decision: "CALL_REQUIRED",
      summary: "조건부 Planner 호출 조건이 충족됐지만 provider가 설정되지 않았습니다.",
      usage: { ...normalizeUsage({}), estimatedCostUsd: 0 },
      counts: { proposed: 0, accepted: 0, rejected: 0 },
      acceptedCandidateIds: [],
      rejectedProposalKeys: [],
      proposals: [],
      integrity: {
        sourceCandidatesPreserved: true,
        providerSkipped: true,
        acceptedProposalsRequireResolverReentry: true,
      },
    });
  }

  let cacheKey = "";
  if (cache && tenantId && cacheSecret) {
    cacheKey = buildCandidatePlannerCacheKey({
      tenantId,
      input,
      model,
      reasoningEffort,
      cacheSecret,
    });
  }

  let providerResult = null;
  let invocationStatus = "CALLED";
  if (cacheKey && typeof cache.get === "function") {
    const cached = await cache.get(cacheKey);
    if (cached?.output) {
      const validation = validateCandidatePlannerModelOutput(cached.output);
      if (validation.valid) {
        providerResult = cached;
        invocationStatus = "CACHE_HIT";
      }
    }
  }

  if (!providerResult) {
    try {
      providerResult = await provider.plan({ input, model, reasoningEffort });
    } catch (error) {
      return finalizeResolution({
        ...base,
        invocation: {
          status: "FAILED_SAFE",
          providerCallCount: 1,
          cacheHit: false,
          model,
          reasoningEffort,
          responseId: "",
          failureCode: normalizeText(error?.code || "PLANNER_PROVIDER_FAILED"),
        },
        decision: "FAILED_SAFE",
        summary: "LLM Candidate Planner 호출이 실패해 기존 결정론적 결과만 유지합니다.",
        usage: { ...normalizeUsage({}), estimatedCostUsd: 0 },
        counts: { proposed: 0, accepted: 0, rejected: 0 },
        acceptedCandidateIds: [],
        rejectedProposalKeys: [],
        proposals: [],
        integrity: {
          sourceCandidatesPreserved: true,
          providerSkipped: false,
          acceptedProposalsRequireResolverReentry: true,
        },
      });
    }
  }

  const modelOutput = providerResult.output;
  const validation = validateCandidatePlannerModelOutput(modelOutput);
  if (!validation.valid) {
    return finalizeResolution({
      ...base,
      invocation: {
        status: "FAILED_SAFE",
        providerCallCount: invocationStatus === "CACHE_HIT" ? 0 : 1,
        cacheHit: invocationStatus === "CACHE_HIT",
        model: normalizeText(providerResult.model || model),
        reasoningEffort: normalizeText(providerResult.reasoningEffort || reasoningEffort),
        responseId: normalizeText(providerResult.responseId || ""),
        failureCode: "MODEL_OUTPUT_INVALID",
      },
      decision: "FAILED_SAFE",
      summary: "LLM Candidate Planner 출력이 strict contract를 통과하지 못해 폐기했습니다.",
      usage: {
        ...normalizeUsage(providerResult.usage || {}),
        estimatedCostUsd: estimateCostUsd(providerResult.usage || {}, providerResult.model || model, pricing),
      },
      counts: { proposed: 0, accepted: 0, rejected: 0 },
      acceptedCandidateIds: [],
      rejectedProposalKeys: [],
      proposals: [],
      modelValidation: validation,
      integrity: {
        sourceCandidatesPreserved: true,
        providerSkipped: invocationStatus === "CACHE_HIT",
        acceptedProposalsRequireResolverReentry: true,
      },
    });
  }

  if (cacheKey && invocationStatus !== "CACHE_HIT" && typeof cache.set === "function") {
    await cache.set(cacheKey, {
      output: modelOutput,
      model: normalizeText(providerResult.model || model),
      reasoningEffort: normalizeText(providerResult.reasoningEffort || reasoningEffort),
      responseId: normalizeText(providerResult.responseId || ""),
      usage: providerResult.usage || {},
    });
  }

  const context = {
    maps: tableAndColumnMaps(input),
    existingSignatures: existingCandidateSignatures({
      candidateFeasibilityResolution,
      candidateRankingResolution,
    }),
    acceptedSignatures: new Set(),
  };
  const proposals = asArray(modelOutput.proposals).map((proposal, index) =>
    validateAndNormalizeProposal(proposal, index, context),
  );
  const accepted = proposals.filter(
    (item) => item.disposition === "ACCEPTED_FOR_REVALIDATION",
  );
  const rejected = proposals.filter((item) => item.disposition === "REJECTED");
  const normalizedUsage = normalizeUsage(providerResult.usage || {});
  return finalizeResolution({
    ...base,
    invocation: {
      status: invocationStatus,
      providerCallCount: invocationStatus === "CACHE_HIT" ? 0 : 1,
      cacheHit: invocationStatus === "CACHE_HIT",
      model: normalizeText(providerResult.model || model),
      reasoningEffort: normalizeText(providerResult.reasoningEffort || reasoningEffort),
      responseId: normalizeText(providerResult.responseId || ""),
      failureCode: "",
    },
    decision:
      modelOutput.decision === "NO_ADDITION"
        ? "NO_ADDITION"
        : accepted.length
          ? "PROPOSED"
          : "NO_ADDITION",
    summary: normalizeText(modelOutput.summary || ""),
    usage: {
      ...normalizedUsage,
      estimatedCostUsd: estimateCostUsd(
        normalizedUsage,
        providerResult.model || model,
        pricing,
      ),
    },
    counts: {
      proposed: proposals.length,
      accepted: accepted.length,
      rejected: rejected.length,
    },
    acceptedCandidateIds: accepted.map((item) => item.candidateId),
    rejectedProposalKeys: rejected.map((item) => item.proposalKey),
    proposals,
    modelOutputSha256: sha256(modelOutput),
    integrity: {
      sourceCandidatesPreserved: true,
      providerSkipped: false,
      acceptedProposalsRequireResolverReentry: true,
      acceptedCandidateIdsUnique:
        accepted.length === new Set(accepted.map((item) => item.candidateId)).size,
      allAcceptedReferencesKnown: accepted.every((item) =>
        item.sourceTableIds.every((tableId) => context.maps.tables.has(tableId)) &&
        item.operandBindings.every((binding) => context.maps.columns.has(binding.columnId)),
      ),
    },
  });
}

function validateQueryCandidatePlannerResolution(document = {}) {
  const errors = [];
  const warnings = [];
  if (document.version !== QUERY_CANDIDATE_PLANNER_RESOLUTION_VERSION) {
    errors.push(issue("version", "INVALID_VERSION", "Planner resolution version이 유효하지 않습니다."));
  }
  if (document.itemVersion !== QUERY_CANDIDATE_PLANNER_ITEM_VERSION) {
    errors.push(issue("itemVersion", "INVALID_ITEM_VERSION", "Planner item version이 유효하지 않습니다."));
  }
  if (document.policy?.version !== QUERY_CANDIDATE_PLANNER_POLICY_VERSION) {
    errors.push(issue("policy.version", "INVALID_POLICY_VERSION", "Planner policy version이 유효하지 않습니다."));
  }
  if (!INVOCATION_STATUS.includes(document.invocation?.status)) {
    errors.push(issue("invocation.status", "INVALID_INVOCATION_STATUS", "invocation status가 유효하지 않습니다."));
  }
  if (!PLANNER_DECISION.includes(document.decision)) {
    errors.push(issue("decision", "INVALID_DECISION", "Planner decision이 유효하지 않습니다."));
  }
  if (document.privacy?.rawRowsSent !== false || document.privacy?.sampleValuesSent !== false || document.privacy?.originalFileSent !== false || document.privacy?.fileNameSent !== false) {
    errors.push(issue("privacy", "PRIVACY_BOUNDARY_VIOLATION", "원본 행·샘플값·파일·파일명은 전송되면 안 됩니다."));
  }
  const proposals = asArray(document.proposals);
  proposals.forEach((proposal, index) => {
    if (!PROPOSAL_DISPOSITION.includes(proposal.disposition)) {
      errors.push(issue(`proposals[${index}].disposition`, "INVALID_PROPOSAL_DISPOSITION", "proposal disposition이 유효하지 않습니다."));
    }
    if (proposal.readyStatusAssigned !== false) {
      errors.push(issue(`proposals[${index}].readyStatusAssigned`, "READY_ASSIGNED_TOO_EARLY", "Planner는 READY를 부여하면 안 됩니다."));
    }
    const expected = sha256({ ...proposal, plannerItemSha256: undefined });
    if (proposal.plannerItemSha256 !== expected) {
      errors.push(issue(`proposals[${index}].plannerItemSha256`, "SHA_MISMATCH", "Planner item SHA가 일치하지 않습니다."));
    }
  });
  const accepted = proposals.filter((item) => item.disposition === "ACCEPTED_FOR_REVALIDATION");
  const rejected = proposals.filter((item) => item.disposition === "REJECTED");
  if (Number(document.counts?.proposed || 0) !== proposals.length || Number(document.counts?.accepted || 0) !== accepted.length || Number(document.counts?.rejected || 0) !== rejected.length) {
    errors.push(issue("counts", "COUNT_MISMATCH", "Planner proposal count가 실제 항목과 다릅니다."));
  }
  if (document.invocation?.status === "SKIPPED" && Number(document.invocation.providerCallCount || 0) !== 0) {
    errors.push(issue("invocation.providerCallCount", "SKIPPED_PROVIDER_CALLED", "SKIPPED 상태에서는 provider 호출이 없어야 합니다."));
  }
  if (document.policy?.sourceCandidatesAreNotRemovedOrMutated !== true || document.integrity?.sourceCandidatesPreserved !== true) {
    errors.push(issue("integrity.sourceCandidatesPreserved", "SOURCE_CANDIDATE_MUTATION", "원본 후보 보존 정책이 필요합니다."));
  }
  if (!/^[a-f0-9]{64}$/.test(document.plannerResolutionSha256 || "")) {
    errors.push(issue("plannerResolutionSha256", "INVALID_SHA", "Planner resolution SHA가 유효하지 않습니다."));
  } else {
    const expected = sha256({ ...document, plannerResolutionSha256: undefined });
    if (expected !== document.plannerResolutionSha256) {
      errors.push(issue("plannerResolutionSha256", "SHA_MISMATCH", "Planner resolution SHA가 일치하지 않습니다."));
    }
  }
  if (document.invocation?.status === "FAILED_SAFE") {
    warnings.push(issue("invocation", "PLANNER_FAILED_SAFE", "Planner 실패로 결정론적 결과만 유지됩니다."));
  }
  return {
    version: "query_candidate_planner_validation_v1",
    valid: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings,
  };
}

function createMockCandidatePlannerProvider({ output, usage = {}, model = "mock-planner" } = {}) {
  let callCount = 0;
  return {
    get callCount() {
      return callCount;
    },
    async plan() {
      callCount += 1;
      return {
        provider: "MOCK",
        model,
        reasoningEffort: "none",
        responseId: `mock-${callCount}`,
        usage,
        output:
          output ||
          {
            version: QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION,
            decision: "NO_ADDITION",
            summary: "결정론적 후보 외에 안전하게 추가할 후보가 없습니다.",
            proposals: [],
          },
      };
    },
  };
}

module.exports = {
  QUERY_CANDIDATE_PLANNER_INPUT_VERSION,
  QUERY_CANDIDATE_PLANNER_MODEL_OUTPUT_VERSION,
  QUERY_CANDIDATE_PLANNER_RESOLUTION_VERSION,
  QUERY_CANDIDATE_PLANNER_ITEM_VERSION,
  QUERY_CANDIDATE_PLANNER_POLICY_VERSION,
  QUERY_CANDIDATE_PLANNER_CACHE_VERSION,
  DEFAULT_MODEL,
  DEFAULT_REASONING_EFFORT,
  MAX_PROPOSALS,
  MIN_PROPOSAL_CONFIDENCE,
  INVOCATION_STATUS,
  PLANNER_DECISION,
  PROPOSAL_DISPOSITION,
  ALLOWED_OPERATIONS,
  MODEL_EVIDENCE_CODES,
  MODEL_OUTPUT_SCHEMA,
  buildPlannerInput,
  deriveSemanticOpportunities,
  evaluateConditionalPlannerTrigger,
  validateCandidatePlannerModelOutput,
  proposalSignature,
  buildCandidatePlannerCacheKey,
  runConditionalCandidatePlanner,
  validateQueryCandidatePlannerResolution,
  normalizeUsage,
  estimateCostUsd,
  createMockCandidatePlannerProvider,
};
