const crypto = require("crypto");
const taxonomy = require("./querySemanticProfilerTaxonomy.json");
const { normalizeText, sha256 } = require("./queryCandidateObservation");
const {
  QUERY_SEMANTIC_PROFILER_PROMPT_VERSION,
} = require("./querySemanticProfilerPrompt");
const { mockClassifyBusinessDomain } = require("./queryBusinessDomainProfiler");

const QUERY_SEMANTIC_PROFILE_VERSION = "llm_semantic_profile_v1";
const QUERY_SEMANTIC_MODEL_OUTPUT_VERSION =
  "query_semantic_profiler_model_output_v1";
const QUERY_SEMANTIC_INPUT_VERSION = "query_semantic_profiler_input_v1";
const QUERY_SEMANTIC_SCHEMA_VERSION = "query_semantic_profiler_schema_v1";
const QUERY_SEMANTIC_CACHE_VERSION = "query_semantic_profiler_cache_v1";
const DEFAULT_MODEL = "gpt-5.6-terra";
const DEFAULT_REASONING_EFFORT = "low";
const MAX_TABLES = 20;
const MAX_COLUMNS = 120;
const MAX_HEADER_LENGTH = 120;

const MODEL_PRICING_USD_PER_MILLION = Object.freeze({
  "gpt-5.6-terra": Object.freeze({
    input: 2,
    cachedInput: 0.2,
    output: 12,
  }),
});

const MODEL_OUTPUT_SCHEMA = Object.freeze({
  type: "object",
  additionalProperties: false,
  required: [
    "version",
    "classification",
    "tableSemantics",
    "columnSemantics",
    "tableRelations",
    "ambiguities",
    "requiresHumanReview",
  ],
  properties: {
    version: {
      type: "string",
      const: QUERY_SEMANTIC_MODEL_OUTPUT_VERSION,
    },
    classification: {
      type: "object",
      additionalProperties: false,
      required: [
        "primaryDomain",
        "secondaryDomains",
        "datasetIntent",
        "confidence",
        "description",
      ],
      properties: {
        primaryDomain: { type: "string", enum: taxonomy.domains },
        secondaryDomains: {
          type: "array",
          maxItems: 3,
          items: {
            type: "string",
            enum: taxonomy.domains.filter((value) => value !== "UNKNOWN"),
          },
        },
        datasetIntent: { type: "string", enum: taxonomy.datasetIntents },
        confidence: { type: "number", minimum: 0, maximum: 1 },
        description: { type: "string" },
      },
    },
    tableSemantics: {
      type: "array",
      maxItems: MAX_TABLES,
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "tableId",
          "tablePurpose",
          "rowGrain",
          "confidence",
          "evidenceCodes",
          "description",
        ],
        properties: {
          tableId: { type: "string" },
          tablePurpose: { type: "string", enum: taxonomy.tablePurposes },
          rowGrain: { type: "string" },
          confidence: { type: "number", minimum: 0, maximum: 1 },
          evidenceCodes: {
            type: "array",
            maxItems: 6,
            items: { type: "string", enum: taxonomy.evidenceCodes },
          },
          description: { type: "string" },
        },
      },
    },
    columnSemantics: {
      type: "array",
      maxItems: MAX_COLUMNS,
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "tableId",
          "columnId",
          "normalizedMeaning",
          "semanticRole",
          "semanticType",
          "metricFamily",
          "defaultAggregation",
          "unitSemantic",
          "decision",
          "confidence",
          "evidenceCodes",
          "description",
        ],
        properties: {
          tableId: { type: "string" },
          columnId: { type: "string" },
          normalizedMeaning: { type: "string" },
          semanticRole: { type: "string", enum: taxonomy.semanticRoles },
          semanticType: { type: "string", enum: taxonomy.semanticTypes },
          metricFamily: { type: "string", enum: taxonomy.metricFamilies },
          defaultAggregation: {
            type: "string",
            enum: taxonomy.defaultAggregations,
          },
          unitSemantic: { type: "string", enum: taxonomy.unitSemantics },
          decision: { type: "string", enum: taxonomy.columnDecisions },
          confidence: { type: "number", minimum: 0, maximum: 1 },
          evidenceCodes: {
            type: "array",
            maxItems: 6,
            items: { type: "string", enum: taxonomy.evidenceCodes },
          },
          description: { type: "string" },
        },
      },
    },
    tableRelations: {
      type: "array",
      maxItems: 40,
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "leftTableId",
          "rightTableId",
          "relationType",
          "cardinality",
          "leftColumnIds",
          "rightColumnIds",
          "confidence",
          "evidenceCodes",
          "description",
        ],
        properties: {
          leftTableId: { type: "string" },
          rightTableId: { type: "string" },
          relationType: { type: "string", enum: taxonomy.relationTypes },
          cardinality: { type: "string", enum: taxonomy.cardinalities },
          leftColumnIds: {
            type: "array",
            maxItems: 12,
            items: { type: "string" },
          },
          rightColumnIds: {
            type: "array",
            maxItems: 12,
            items: { type: "string" },
          },
          confidence: { type: "number", minimum: 0, maximum: 1 },
          evidenceCodes: {
            type: "array",
            maxItems: 6,
            items: { type: "string", enum: taxonomy.evidenceCodes },
          },
          description: { type: "string" },
        },
      },
    },
    ambiguities: {
      type: "array",
      maxItems: 20,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["code", "description", "tableIds", "columnIds"],
        properties: {
          code: { type: "string", enum: taxonomy.ambiguityCodes },
          description: { type: "string" },
          tableIds: {
            type: "array",
            maxItems: 12,
            items: { type: "string" },
          },
          columnIds: {
            type: "array",
            maxItems: 24,
            items: { type: "string" },
          },
        },
      },
    },
    requiresHumanReview: { type: "boolean" },
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

function truncateText(value, maximum = MAX_HEADER_LENGTH) {
  const text = normalizeText(value);
  return text.length > maximum ? `${text.slice(0, maximum - 1)}…` : text;
}

function sanitizeColumn(column = {}) {
  const stats =
    column.stats && typeof column.stats === "object" ? column.stats : {};
  return {
    columnId: normalizeText(column.columnId || ""),
    header: truncateText(column.normalizedHeader || column.sourceHeader || ""),
    dataType: normalizeText(column.dataType || "unknown"),
    dataTypeConfidence: round(column.dataTypeConfidence),
    semanticRole: normalizeText(column.semanticRole || "unknown"),
    semanticType: normalizeText(column.semanticType || "unknown"),
    metricFamily: normalizeText(column.metricFamily || ""),
    roleConfidence: round(column.roleConfidence),
    supportedOperations: unique(column.supportedOperations).sort(),
    stats: {
      sampledValueCount: Number(stats.sampledValueCount) || 0,
      nonEmptyRatio: round(stats.nonEmptyRatio),
      uniqueRatio: round(stats.uniqueRatio),
    },
    issueCodes: unique(
      asArray(column.issues).map((issue) => issue?.code || ""),
    ).sort(),
  };
}

function sanitizeTable(table = {}, remainingColumns = MAX_COLUMNS) {
  const columns = asArray(table.columns)
    .slice(0, Math.max(0, remainingColumns))
    .map(sanitizeColumn);
  const quality =
    table.quality && typeof table.quality === "object" ? table.quality : {};
  return {
    tableId: normalizeText(table.tableId || ""),
    sourceTableId: normalizeText(table.sourceTableId || ""),
    sourceSheetName: truncateText(table.sourceSheetName || ""),
    flags: {
      primary: table.flags?.primary === true,
      analysisEligible: table.flags?.analysisEligible === true,
      templateEligible: table.flags?.templateEligible === true,
      virtual: table.flags?.virtual === true,
    },
    shape: {
      rowCount: Number(table.shape?.rowCount) || 0,
      columnCount: Number(table.shape?.columnCount) || columns.length,
      mergedHeader: table.shape?.mergedHeader === true,
      subtotalRowCount: Number(table.shape?.subtotalRowCount) || 0,
      totalRowCount: Number(table.shape?.totalRowCount) || 0,
    },
    quality: {
      nonEmptyRatio: round(quality.nonEmptyRatio),
    },
    availableRoles: unique(table.availableRoles).sort(),
    metricFamilies: unique(table.metricFamilies).sort(),
    supportedOperations: unique(table.supportedOperations).sort(),
    issueCodes: unique(
      asArray(table.issues).map((issue) => issue?.code || ""),
    ).sort(),
    columns,
  };
}

function buildSemanticProfilerInput({ semanticProfile } = {}) {
  if (!semanticProfile || typeof semanticProfile !== "object") {
    throw new TypeError("semanticProfile 객체가 필요합니다.");
  }
  const allTables = asArray(semanticProfile.tables).filter(
    (table) => table?.flags?.analysisEligible === true,
  );
  const selectedTables = [];
  let remainingColumns = MAX_COLUMNS;
  for (const table of allTables.slice(0, MAX_TABLES)) {
    const sanitized = sanitizeTable(table, remainingColumns);
    remainingColumns -= sanitized.columns.length;
    selectedTables.push(sanitized);
    if (remainingColumns <= 0) break;
  }
  const includedTableIds = new Set(
    selectedTables.map((table) => table.tableId).filter(Boolean),
  );
  for (const table of selectedTables) {
    if (table.sourceTableId && !includedTableIds.has(table.sourceTableId)) {
      table.sourceTableId = "";
    }
  }
  const sourceColumnCount = allTables.reduce(
    (sum, table) => sum + asArray(table.columns).length,
    0,
  );
  const includedColumnCount = selectedTables.reduce(
    (sum, table) => sum + table.columns.length,
    0,
  );
  const input = {
    version: QUERY_SEMANTIC_INPUT_VERSION,
    source: {
      semanticProfileVersion: normalizeText(semanticProfile.version || ""),
      semanticProfileSha256: normalizeText(
        semanticProfile.profileSha256 || sha256(semanticProfile),
      ),
    },
    counts: {
      sourceTableCount: allTables.length,
      includedTableCount: selectedTables.length,
      sourceColumnCount,
      includedColumnCount,
      tablesTruncated: allTables.length > selectedTables.length,
      columnsTruncated: sourceColumnCount > includedColumnCount,
    },
    deterministicSummary: {
      availableRoles: unique(semanticProfile.availableRoles).sort(),
      metricFamilies: unique(semanticProfile.metricFamilies).sort(),
      supportedOperations: unique(semanticProfile.supportedOperations).sort(),
      issueCodes: unique(
        asArray(semanticProfile.issues).map((issue) => issue?.code || ""),
      ).sort(),
    },
    tables: selectedTables,
    privacy: {
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      originalFileIncluded: false,
      fileNameIncluded: false,
    },
  };
  input.inputSha256 = sha256(input);
  return input;
}

function idMaps(input = {}) {
  const tableIds = new Set();
  const columnIds = new Set();
  const columnsByTable = new Map();
  const tableById = new Map();
  for (const table of asArray(input.tables)) {
    const tableId = normalizeText(table.tableId || "");
    if (!tableId) continue;
    tableIds.add(tableId);
    tableById.set(tableId, table);
    const set = new Set();
    for (const column of asArray(table.columns)) {
      const columnId = normalizeText(column.columnId || "");
      if (!columnId) continue;
      columnIds.add(columnId);
      set.add(columnId);
    }
    columnsByTable.set(tableId, set);
  }
  return { tableIds, columnIds, columnsByTable, tableById };
}

function duplicateValues(values = []) {
  const seen = new Set();
  const duplicates = new Set();
  for (const value of values) {
    if (seen.has(value)) duplicates.add(value);
    seen.add(value);
  }
  return [...duplicates];
}

function validateEnum(errors, code, value, allowed, detail = {}) {
  if (!allowed.includes(value)) errors.push({ code, actual: value, ...detail });
}

function validateReferences({ errors, item, index, kind, maps }) {
  for (const tableId of unique(item?.tableIds || item?.tableId)) {
    if (!maps.tableIds.has(tableId)) {
      errors.push({ code: `${kind}_TABLE_ID_UNKNOWN`, index, tableId });
    }
  }
  for (const columnId of unique(item?.columnIds)) {
    if (!maps.columnIds.has(columnId)) {
      errors.push({ code: `${kind}_COLUMN_ID_UNKNOWN`, index, columnId });
    }
  }
}

function validateSemanticProfilerModelOutput(output, input) {
  const errors = [];
  const warnings = [];
  if (!output || typeof output !== "object" || Array.isArray(output)) {
    return {
      valid: false,
      errorCount: 1,
      warningCount: 0,
      errors: [{ code: "MODEL_OUTPUT_NOT_OBJECT" }],
      warnings,
    };
  }
  if (output.version !== QUERY_SEMANTIC_MODEL_OUTPUT_VERSION) {
    errors.push({
      code: "MODEL_OUTPUT_VERSION_INVALID",
      actual: output.version,
    });
  }
  const classification = output.classification || {};
  validateEnum(
    errors,
    "PRIMARY_DOMAIN_INVALID",
    classification.primaryDomain,
    taxonomy.domains,
  );
  validateEnum(
    errors,
    "DATASET_INTENT_INVALID",
    classification.datasetIntent,
    taxonomy.datasetIntents,
  );
  const secondaryDomains = asArray(classification.secondaryDomains);
  if (secondaryDomains.length > 3) {
    errors.push({ code: "SECONDARY_DOMAIN_LIMIT_EXCEEDED" });
  }
  if (duplicateValues(secondaryDomains).length) {
    errors.push({ code: "SECONDARY_DOMAIN_DUPLICATED" });
  }
  for (const domain of secondaryDomains) {
    if (!taxonomy.domains.includes(domain) || domain === "UNKNOWN") {
      errors.push({ code: "SECONDARY_DOMAIN_INVALID", actual: domain });
    }
    if (domain === classification.primaryDomain) {
      errors.push({ code: "SECONDARY_DOMAIN_EQUALS_PRIMARY", actual: domain });
    }
  }
  if (!Number.isFinite(Number(classification.confidence))) {
    errors.push({ code: "CLASSIFICATION_CONFIDENCE_INVALID" });
  } else if (
    Number(classification.confidence) < 0 ||
    Number(classification.confidence) > 1
  ) {
    errors.push({ code: "CLASSIFICATION_CONFIDENCE_OUT_OF_RANGE" });
  }
  if (typeof output.requiresHumanReview !== "boolean") {
    errors.push({ code: "HUMAN_REVIEW_FLAG_INVALID" });
  }

  const maps = idMaps(input);
  const expectedTableIds = [...maps.tableIds].sort();
  const expectedColumnIds = [...maps.columnIds].sort();

  if (!Array.isArray(output.tableSemantics)) {
    errors.push({ code: "TABLE_SEMANTICS_NOT_ARRAY" });
  } else {
    const actualIds = output.tableSemantics.map((item) =>
      normalizeText(item?.tableId),
    );
    for (const duplicate of duplicateValues(actualIds)) {
      errors.push({ code: "TABLE_SEMANTICS_DUPLICATED", tableId: duplicate });
    }
    for (const tableId of expectedTableIds) {
      if (!actualIds.includes(tableId)) {
        errors.push({ code: "TABLE_SEMANTICS_MISSING", tableId });
      }
    }
    for (const tableId of actualIds) {
      if (!maps.tableIds.has(tableId)) {
        errors.push({ code: "TABLE_SEMANTICS_TABLE_ID_UNKNOWN", tableId });
      }
    }
    output.tableSemantics.forEach((item, index) => {
      validateEnum(
        errors,
        "TABLE_PURPOSE_INVALID",
        item?.tablePurpose,
        taxonomy.tablePurposes,
        { index },
      );
      for (const code of asArray(item?.evidenceCodes)) {
        validateEnum(
          errors,
          "TABLE_EVIDENCE_CODE_INVALID",
          code,
          taxonomy.evidenceCodes,
          { index },
        );
      }
    });
  }

  if (!Array.isArray(output.columnSemantics)) {
    errors.push({ code: "COLUMN_SEMANTICS_NOT_ARRAY" });
  } else {
    const actualIds = output.columnSemantics.map((item) =>
      normalizeText(item?.columnId),
    );
    for (const duplicate of duplicateValues(actualIds)) {
      errors.push({ code: "COLUMN_SEMANTICS_DUPLICATED", columnId: duplicate });
    }
    for (const columnId of expectedColumnIds) {
      if (!actualIds.includes(columnId)) {
        errors.push({ code: "COLUMN_SEMANTICS_MISSING", columnId });
      }
    }
    for (const columnId of actualIds) {
      if (!maps.columnIds.has(columnId)) {
        errors.push({ code: "COLUMN_SEMANTICS_COLUMN_ID_UNKNOWN", columnId });
      }
    }
    output.columnSemantics.forEach((item, index) => {
      const tableId = normalizeText(item?.tableId);
      const columnId = normalizeText(item?.columnId);
      if (!maps.tableIds.has(tableId)) {
        errors.push({
          code: "COLUMN_SEMANTICS_TABLE_ID_UNKNOWN",
          index,
          tableId,
        });
      } else if (!maps.columnsByTable.get(tableId)?.has(columnId)) {
        errors.push({
          code: "COLUMN_SEMANTICS_COLUMN_TABLE_MISMATCH",
          index,
          tableId,
          columnId,
        });
      }
      validateEnum(
        errors,
        "COLUMN_ROLE_INVALID",
        item?.semanticRole,
        taxonomy.semanticRoles,
        { index },
      );
      validateEnum(
        errors,
        "COLUMN_TYPE_INVALID",
        item?.semanticType,
        taxonomy.semanticTypes,
        { index },
      );
      validateEnum(
        errors,
        "COLUMN_METRIC_FAMILY_INVALID",
        item?.metricFamily,
        taxonomy.metricFamilies,
        { index },
      );
      validateEnum(
        errors,
        "COLUMN_AGGREGATION_INVALID",
        item?.defaultAggregation,
        taxonomy.defaultAggregations,
        { index },
      );
      validateEnum(
        errors,
        "COLUMN_UNIT_INVALID",
        item?.unitSemantic,
        taxonomy.unitSemantics,
        { index },
      );
      validateEnum(
        errors,
        "COLUMN_DECISION_INVALID",
        item?.decision,
        taxonomy.columnDecisions,
        { index },
      );
      for (const code of asArray(item?.evidenceCodes)) {
        validateEnum(
          errors,
          "COLUMN_EVIDENCE_CODE_INVALID",
          code,
          taxonomy.evidenceCodes,
          { index },
        );
      }
      if (item?.semanticType === "measure" && item?.metricFamily === "NONE") {
        warnings.push({
          code: "MEASURE_WITHOUT_METRIC_FAMILY",
          index,
          columnId,
        });
      }
      if (
        item?.semanticType !== "measure" &&
        item?.defaultAggregation === "sum"
      ) {
        warnings.push({ code: "NON_MEASURE_SUM_AGGREGATION", index, columnId });
      }
    });
  }

  if (!Array.isArray(output.tableRelations)) {
    errors.push({ code: "TABLE_RELATIONS_NOT_ARRAY" });
  } else {
    const relationKeys = [];
    output.tableRelations.forEach((item, index) => {
      const leftTableId = normalizeText(item?.leftTableId);
      const rightTableId = normalizeText(item?.rightTableId);
      if (!maps.tableIds.has(leftTableId)) {
        errors.push({
          code: "RELATION_LEFT_TABLE_UNKNOWN",
          index,
          leftTableId,
        });
      }
      if (!maps.tableIds.has(rightTableId)) {
        errors.push({
          code: "RELATION_RIGHT_TABLE_UNKNOWN",
          index,
          rightTableId,
        });
      }
      if (leftTableId && leftTableId === rightTableId) {
        errors.push({
          code: "RELATION_SELF_REFERENCE",
          index,
          tableId: leftTableId,
        });
      }
      validateEnum(
        errors,
        "RELATION_TYPE_INVALID",
        item?.relationType,
        taxonomy.relationTypes,
        { index },
      );
      validateEnum(
        errors,
        "RELATION_CARDINALITY_INVALID",
        item?.cardinality,
        taxonomy.cardinalities,
        { index },
      );
      const leftColumns = unique(item?.leftColumnIds);
      const rightColumns = unique(item?.rightColumnIds);
      if (leftColumns.length !== rightColumns.length) {
        errors.push({ code: "RELATION_KEY_PAIR_LENGTH_MISMATCH", index });
      }
      for (const columnId of leftColumns) {
        if (!maps.columnsByTable.get(leftTableId)?.has(columnId)) {
          errors.push({
            code: "RELATION_LEFT_COLUMN_INVALID",
            index,
            leftTableId,
            columnId,
          });
        }
      }
      for (const columnId of rightColumns) {
        if (!maps.columnsByTable.get(rightTableId)?.has(columnId)) {
          errors.push({
            code: "RELATION_RIGHT_COLUMN_INVALID",
            index,
            rightTableId,
            columnId,
          });
        }
      }
      for (const code of asArray(item?.evidenceCodes)) {
        validateEnum(
          errors,
          "RELATION_EVIDENCE_CODE_INVALID",
          code,
          taxonomy.evidenceCodes,
          { index },
        );
      }
      relationKeys.push(
        [
          leftTableId,
          rightTableId,
          item?.relationType,
          leftColumns.join(","),
          rightColumns.join(","),
        ].join("|"),
      );
    });
    for (const duplicate of duplicateValues(relationKeys)) {
      errors.push({
        code: "TABLE_RELATION_DUPLICATED",
        relationKey: duplicate,
      });
    }
  }

  if (!Array.isArray(output.ambiguities)) {
    errors.push({ code: "AMBIGUITIES_NOT_ARRAY" });
  } else {
    output.ambiguities.forEach((item, index) => {
      validateEnum(
        errors,
        "AMBIGUITY_CODE_INVALID",
        item?.code,
        taxonomy.ambiguityCodes,
        { index },
      );
      validateReferences({ errors, item, index, kind: "AMBIGUITY", maps });
    });
  }

  if (
    (input.counts?.tablesTruncated || input.counts?.columnsTruncated) &&
    !asArray(output.ambiguities).some(
      (item) => item?.code === "INPUT_TRUNCATED",
    )
  ) {
    warnings.push({ code: "TRUNCATED_INPUT_WITHOUT_AMBIGUITY" });
  }
  if (
    Number(classification.confidence) < 0.7 &&
    output.requiresHumanReview !== true
  ) {
    warnings.push({ code: "LOW_CONFIDENCE_WITHOUT_HUMAN_REVIEW" });
  }

  return {
    valid: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings,
  };
}

function normalizeUsage(usage = {}) {
  const inputDetails =
    usage.input_tokens_details || usage.inputTokensDetails || {};
  const outputDetails =
    usage.output_tokens_details || usage.outputTokensDetails || {};
  return {
    inputTokens: Number(usage.input_tokens ?? usage.inputTokens) || 0,
    cachedInputTokens:
      Number(
        inputDetails.cached_tokens ??
          inputDetails.cachedTokens ??
          usage.cached_input_tokens ??
          usage.cachedInputTokens,
      ) || 0,
    cacheWriteTokens:
      Number(
        inputDetails.cache_write_tokens ??
          inputDetails.cacheWriteTokens ??
          usage.cache_write_tokens ??
          usage.cacheWriteTokens,
      ) || 0,
    outputTokens: Number(usage.output_tokens ?? usage.outputTokens) || 0,
    reasoningTokens:
      Number(outputDetails.reasoning_tokens ?? outputDetails.reasoningTokens) ||
      0,
    totalTokens: Number(usage.total_tokens ?? usage.totalTokens) || 0,
  };
}

function estimateCostUsd(usage = {}, model = DEFAULT_MODEL, pricingOverride) {
  const normalized = normalizeUsage(usage);
  const pricing =
    pricingOverride || MODEL_PRICING_USD_PER_MILLION[model] || null;
  if (!pricing) return null;
  const cached = Math.min(normalized.inputTokens, normalized.cachedInputTokens);
  const cacheWrite = Math.min(
    Math.max(0, normalized.inputTokens - cached),
    normalized.cacheWriteTokens,
  );
  const uncached = Math.max(0, normalized.inputTokens - cached - cacheWrite);
  const cost =
    (uncached * Number(pricing.input || 0) +
      cached * Number(pricing.cachedInput || 0) +
      cacheWrite * Number(pricing.input || 0) * 1.25 +
      normalized.outputTokens * Number(pricing.output || 0)) /
    1_000_000;
  return round(cost, 8);
}

function normalizedEvidenceCodes(values) {
  return unique(values).filter((value) =>
    taxonomy.evidenceCodes.includes(value),
  );
}

function buildSemanticProfile({
  semanticProfile,
  input,
  modelOutput,
  provider = "OPENAI_RESPONSES",
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  responseId = "",
  usage = {},
  pricing,
} = {}) {
  const profilerInput =
    input || buildSemanticProfilerInput({ semanticProfile });
  const validation = validateSemanticProfilerModelOutput(
    modelOutput,
    profilerInput,
  );
  if (!validation.valid) {
    const error = new Error(
      "통합 Semantic Profiler 모델 출력 검증에 실패했습니다.",
    );
    error.code = "SEMANTIC_PROFILER_MODEL_OUTPUT_INVALID";
    error.validation = validation;
    throw error;
  }
  const normalizedUsage = normalizeUsage(usage);
  const profile = {
    version: QUERY_SEMANTIC_PROFILE_VERSION,
    schemaVersion: QUERY_SEMANTIC_SCHEMA_VERSION,
    taxonomyVersion: taxonomy.version,
    promptVersion: QUERY_SEMANTIC_PROFILER_PROMPT_VERSION,
    source: {
      caseId: normalizeText(semanticProfile?.source?.caseId || ""),
      fileName: normalizeText(semanticProfile?.source?.fileName || ""),
      semanticProfileVersion: normalizeText(
        semanticProfile?.version ||
          profilerInput.source?.semanticProfileVersion ||
          "",
      ),
      semanticProfileSha256: normalizeText(
        semanticProfile?.profileSha256 ||
          profilerInput.source?.semanticProfileSha256 ||
          "",
      ),
      inputSha256: profilerInput.inputSha256,
    },
    model: {
      provider: normalizeText(provider),
      model: normalizeText(model),
      reasoningEffort: normalizeText(reasoningEffort),
      responseId: normalizeText(responseId),
    },
    privacy: {
      rawRowsSent: false,
      sampleValuesSent: false,
      originalFileSent: false,
      fileNameSent: false,
      includedTableCount: profilerInput.counts.includedTableCount,
      includedColumnCount: profilerInput.counts.includedColumnCount,
    },
    classification: {
      primaryDomain: modelOutput.classification.primaryDomain,
      secondaryDomains: unique(modelOutput.classification.secondaryDomains),
      datasetIntent: modelOutput.classification.datasetIntent,
      confidence: round(clamp(modelOutput.classification.confidence)),
      description: normalizeText(modelOutput.classification.description || ""),
    },
    tableSemantics: asArray(modelOutput.tableSemantics).map((item) => ({
      tableId: normalizeText(item.tableId || ""),
      tablePurpose: item.tablePurpose,
      rowGrain: normalizeText(item.rowGrain || ""),
      confidence: round(clamp(item.confidence)),
      evidenceCodes: normalizedEvidenceCodes(item.evidenceCodes),
      description: normalizeText(item.description || ""),
    })),
    columnSemantics: asArray(modelOutput.columnSemantics).map((item) => ({
      tableId: normalizeText(item.tableId || ""),
      columnId: normalizeText(item.columnId || ""),
      normalizedMeaning: normalizeText(item.normalizedMeaning || ""),
      semanticRole: item.semanticRole,
      semanticType: item.semanticType,
      metricFamily: item.metricFamily,
      defaultAggregation: item.defaultAggregation,
      unitSemantic: item.unitSemantic,
      decision: item.decision,
      confidence: round(clamp(item.confidence)),
      evidenceCodes: normalizedEvidenceCodes(item.evidenceCodes),
      description: normalizeText(item.description || ""),
    })),
    tableRelations: asArray(modelOutput.tableRelations).map((item) => ({
      leftTableId: normalizeText(item.leftTableId || ""),
      rightTableId: normalizeText(item.rightTableId || ""),
      relationType: item.relationType,
      cardinality: item.cardinality,
      leftColumnIds: unique(item.leftColumnIds),
      rightColumnIds: unique(item.rightColumnIds),
      confidence: round(clamp(item.confidence)),
      evidenceCodes: normalizedEvidenceCodes(item.evidenceCodes),
      description: normalizeText(item.description || ""),
    })),
    ambiguities: asArray(modelOutput.ambiguities).map((item) => ({
      code: item.code,
      description: normalizeText(item.description || ""),
      tableIds: unique(item.tableIds),
      columnIds: unique(item.columnIds),
    })),
    requiresHumanReview: modelOutput.requiresHumanReview === true,
    usage: {
      ...normalizedUsage,
      estimatedCostUsd: estimateCostUsd(normalizedUsage, model, pricing),
    },
    integrity: {
      modelOutputSha256: sha256(modelOutput),
      validationWarningCount: validation.warningCount,
      inputTableCount: profilerInput.counts.includedTableCount,
      inputColumnCount: profilerInput.counts.includedColumnCount,
      tableSemanticCount: asArray(modelOutput.tableSemantics).length,
      columnSemanticCount: asArray(modelOutput.columnSemantics).length,
      relationCount: asArray(modelOutput.tableRelations).length,
    },
  };
  profile.profileSha256 = sha256({ ...profile, profileSha256: undefined });
  return profile;
}

function validateSemanticProfile(profile) {
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
  if (profile.version !== QUERY_SEMANTIC_PROFILE_VERSION) {
    errors.push({ code: "PROFILE_VERSION_INVALID" });
  }
  if (profile.schemaVersion !== QUERY_SEMANTIC_SCHEMA_VERSION) {
    errors.push({ code: "PROFILE_SCHEMA_VERSION_INVALID" });
  }
  if (profile.taxonomyVersion !== taxonomy.version) {
    errors.push({ code: "PROFILE_TAXONOMY_VERSION_INVALID" });
  }
  validateEnum(
    errors,
    "PROFILE_PRIMARY_DOMAIN_INVALID",
    profile.classification?.primaryDomain,
    taxonomy.domains,
  );
  validateEnum(
    errors,
    "PROFILE_DATASET_INTENT_INVALID",
    profile.classification?.datasetIntent,
    taxonomy.datasetIntents,
  );
  if (profile.privacy?.rawRowsSent !== false) {
    errors.push({ code: "PROFILE_RAW_ROWS_PRIVACY_INVALID" });
  }
  if (profile.privacy?.sampleValuesSent !== false) {
    errors.push({ code: "PROFILE_SAMPLE_VALUES_PRIVACY_INVALID" });
  }
  if (profile.privacy?.originalFileSent !== false) {
    errors.push({ code: "PROFILE_ORIGINAL_FILE_PRIVACY_INVALID" });
  }
  if (profile.privacy?.fileNameSent !== false) {
    errors.push({ code: "PROFILE_FILE_NAME_PRIVACY_INVALID" });
  }
  if (
    profile.integrity?.tableSemanticCount !==
    profile.privacy?.includedTableCount
  ) {
    errors.push({ code: "PROFILE_TABLE_SEMANTIC_COUNT_MISMATCH" });
  }
  if (
    profile.integrity?.columnSemanticCount !==
    profile.privacy?.includedColumnCount
  ) {
    errors.push({ code: "PROFILE_COLUMN_SEMANTIC_COUNT_MISMATCH" });
  }
  if (!/^[a-f0-9]{64}$/.test(profile.profileSha256 || "")) {
    errors.push({ code: "PROFILE_SHA256_INVALID" });
  } else {
    const expected = sha256({ ...profile, profileSha256: undefined });
    if (expected !== profile.profileSha256) {
      errors.push({ code: "PROFILE_SHA256_MISMATCH" });
    }
  }
  if (profile.integrity?.validationWarningCount > 0) {
    warnings.push({ code: "PROFILE_MODEL_OUTPUT_WARNINGS_PRESENT" });
  }
  return {
    valid: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings,
  };
}

function buildSemanticProfilerCacheKey({
  tenantId,
  semanticProfile,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  cacheSecret,
} = {}) {
  const tenant = normalizeText(tenantId);
  if (!tenant) throw new Error("tenantId가 필요합니다.");
  if (!cacheSecret) throw new Error("cacheSecret이 필요합니다.");
  const semanticProfileSha256 = normalizeText(
    semanticProfile?.profileSha256 || sha256(semanticProfile || {}),
  );
  const identity = {
    version: QUERY_SEMANTIC_CACHE_VERSION,
    tenantId: tenant,
    semanticProfileSha256,
    model,
    reasoningEffort,
    promptVersion: QUERY_SEMANTIC_PROFILER_PROMPT_VERSION,
    schemaVersion: QUERY_SEMANTIC_SCHEMA_VERSION,
    taxonomyVersion: taxonomy.version,
  };
  return crypto
    .createHmac("sha256", Buffer.from(String(cacheSecret)))
    .update(JSON.stringify(identity))
    .digest("hex");
}

async function profileSemantics({
  semanticProfile,
  provider,
  cache,
  tenantId,
  cacheSecret,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  pricing,
} = {}) {
  const input = buildSemanticProfilerInput({ semanticProfile });
  const cacheKey = buildSemanticProfilerCacheKey({
    tenantId,
    semanticProfile,
    model,
    reasoningEffort,
    cacheSecret,
  });
  if (cache && typeof cache.get === "function") {
    const cached = await cache.get(cacheKey);
    if (cached) {
      const validation = validateSemanticProfile(cached);
      if (validation.valid) {
        return {
          profile: cached,
          execution: {
            cacheHit: true,
            providerCalled: false,
            cacheKey,
          },
        };
      }
    }
  }
  if (!provider || typeof provider.profile !== "function") {
    throw new TypeError("캐시 미적중 시 provider.profile 함수가 필요합니다.");
  }
  const response = await provider.profile({ input, model, reasoningEffort });
  const profile = buildSemanticProfile({
    semanticProfile,
    input,
    modelOutput: response.output,
    provider: response.provider || "OPENAI_RESPONSES",
    model: response.model || model,
    reasoningEffort: response.reasoningEffort || reasoningEffort,
    responseId: response.responseId || "",
    usage: response.usage || {},
    pricing,
  });
  if (cache && typeof cache.set === "function") {
    await cache.set(cacheKey, profile);
  }
  return {
    profile,
    execution: {
      cacheHit: false,
      providerCalled: true,
      cacheKey,
    },
  };
}

function businessDomainMockInputFromSemanticProfilerInput(input = {}) {
  return {
    availableRoles: unique(input.deterministicSummary?.availableRoles).sort(),
    metricFamilies: unique(input.deterministicSummary?.metricFamilies).sort(),
    supportedOperations: unique(
      input.deterministicSummary?.supportedOperations,
    ).sort(),
    issueCodes: unique(input.deterministicSummary?.issueCodes).sort(),
    tables: asArray(input.tables).map((table) => ({
      tableId: table.tableId,
      sourceTableId: table.sourceTableId || "",
      sourceSheetName: table.sourceSheetName || "",
      flags: table.flags || {},
      shape: table.shape || {},
      availableRoles: unique(table.availableRoles).sort(),
      metricFamilies: unique(table.metricFamilies).sort(),
      supportedOperations: unique(table.supportedOperations).sort(),
      columns: asArray(table.columns).map((column) => ({
        columnId: column.columnId,
        header: column.header || "",
        semanticRole: column.semanticRole || "",
        semanticType: column.semanticType || "",
        metricFamily: column.metricFamily || "",
      })),
    })),
  };
}

function inferDomain(input = {}) {
  // Patch 6A is the deterministic mock authority for business-domain
  // classification. The integrated profiler must not maintain a second,
  // independently ordered set of domain heuristics because priority changes
  // can make the two mock paths disagree (for example event-applicant vs
  // course-evaluation fixtures). Production OpenAI classification is not
  // affected by this delegation; it is used only by the offline mock path.
  const domainOutput = mockClassifyBusinessDomain(
    businessDomainMockInputFromSemanticProfilerInput(input),
  );
  const description = normalizeText(
    asArray(domainOutput.evidence)[0]?.description ||
      "기존 업무 영역 mock 분류 결과를 통합 Semantic Profiler가 재사용합니다.",
  );
  return [
    domainOutput.primaryDomain,
    domainOutput.datasetIntent,
    Number(domainOutput.confidence) || 0,
    description,
  ];
}

function headerMatches(header, patterns) {
  return patterns.some((pattern) => pattern.test(header));
}

function inferColumnSemantic(column = {}) {
  const header = normalizeText(column.header || "").toLowerCase();
  const currentRole = normalizeText(
    column.semanticRole || "unknown",
  ).toLowerCase();
  const currentType = normalizeText(
    column.semanticType || "unknown",
  ).toLowerCase();
  const currentMetric = normalizeText(column.metricFamily || "");
  let semanticRole = taxonomy.semanticRoles.includes(currentRole)
    ? currentRole
    : "unknown";
  let semanticType = taxonomy.semanticTypes.includes(currentType)
    ? currentType
    : "unknown";
  let metricFamily = taxonomy.metricFamilies.includes(currentMetric)
    ? currentMetric
    : "NONE";
  let defaultAggregation = "none";
  let unitSemantic = "NONE";
  let confidence = Math.max(0.65, Number(column.roleConfidence) || 0);
  let evidenceCodes = ["DETERMINISTIC_ROLE", "HEADER_SEMANTICS"];

  const rules = [
    [[/이름|성명|담당자|직원|사원|교수|학생/], "person", "dimension"],
    [[/기관|소속|부서|조직|회사/], "organization", "dimension"],
    [[/고객|거래처|수요처/], "customer", "dimension"],
    [[/공급|협력사|벤더|업체/], "vendor", "dimension"],
    [[/상품|제품|품목|물품/], "product", "dimension"],
    [[/상태|여부|구분|진행/], "status", "dimension"],
    [[/날짜|일자|일시|등록일|신청일|거래일/], "date", "temporal"],
    [[/월|분기|연도|기간|년도/], "period", "temporal"],
    [[/지역|주소|장소/], "location", "dimension"],
    [[/과제|프로젝트|사업명/], "project", "dimension"],
    [[/행사|세미나|워크숍/], "event", "dimension"],
    [[/강의|강좌|교과|과목/], "course", "dimension"],
    [[/질문|문항/], "question", "dimension"],
    [[/응답|답변|의견/], "response", "text"],
    [[/매출|수익|매출액/], "revenue", "measure"],
    [[/예산/], "budget", "measure"],
    [[/집행|실적|실제/], "actual", "measure"],
    [[/비용|원가|지출|경비/], "cost", "measure"],
    [[/수량|개수|건수|인원|재고량/], "quantity", "measure"],
    [[/점수|평점/], "score", "measure"],
    [[/비율|율|달성률|참석률/], "rate", "measure"],
    [[/퍼센트|백분율|%/], "percentage", "measure"],
    [[/순위|랭킹/], "rank", "measure"],
    [[/목표/], "target", "measure"],
    [[/차이|증감|편차/], "variance", "measure"],
    [[/코드|번호|id|식별/], "identifier", "identifier"],
  ];
  for (const [patterns, role, type] of rules) {
    if (headerMatches(header, patterns)) {
      semanticRole = role;
      semanticType = type;
      confidence = Math.max(confidence, 0.9);
      break;
    }
  }
  const metricByRole = {
    revenue: "revenue",
    amount: "amount",
    budget: "budget",
    actual: "actual",
    cost: "cost",
    quantity: "quantity",
    count: "count",
    score: "score",
    rating: "rating",
    rate: "rate",
    percentage: "percentage",
    duration: "duration",
    target: "target",
    variance: "variance",
  };
  if (metricByRole[semanticRole]) metricFamily = metricByRole[semanticRole];
  if (semanticType === "measure") {
    defaultAggregation = ["rate", "percentage", "score", "rating"].includes(
      metricFamily,
    )
      ? "average"
      : "sum";
  } else if (semanticRole === "identifier") {
    defaultAggregation = "countDistinct";
  }
  if (
    ["revenue", "amount", "budget", "actual", "cost", "expense"].includes(
      metricFamily,
    )
  ) {
    unitSemantic = "currency";
  } else if (["quantity", "count", "inventory"].includes(metricFamily)) {
    unitSemantic = "count";
  } else if (["rate", "percentage", "achievement"].includes(metricFamily)) {
    unitSemantic = "percent";
  } else if (["score", "rating"].includes(metricFamily)) {
    unitSemantic = "score";
  } else if (semanticType === "temporal") {
    unitSemantic = "date";
  }
  const normalizedCurrentMetric = currentMetric || "NONE";
  const roleChanged = semanticRole !== currentRole && currentRole !== "unknown";
  const metricChanged = metricFamily !== normalizedCurrentMetric;
  let decision = "KEEP";
  if (
    currentRole === "unknown" ||
    currentRole === "dimension" ||
    currentRole === "measure"
  ) {
    decision = semanticRole === "unknown" ? "UNKNOWN" : "REFINE";
  } else if (roleChanged) {
    decision = "REPLACE";
  } else if (metricChanged) {
    decision = "REFINE";
  }
  if (semanticRole === "unknown") {
    semanticType = semanticType === "unknown" ? "unknown" : semanticType;
    metricFamily = metricFamily || "NONE";
    confidence = Math.min(confidence, 0.55);
    evidenceCodes = ["INSUFFICIENT_EVIDENCE"];
  }
  return {
    normalizedMeaning: normalizeText(
      column.header || semanticRole || "unknown",
    ),
    semanticRole,
    semanticType,
    metricFamily: metricFamily || "NONE",
    defaultAggregation,
    unitSemantic,
    decision,
    confidence: round(clamp(confidence)),
    evidenceCodes: unique(evidenceCodes),
    description:
      semanticRole === "unknown"
        ? "헤더와 결정론적 역할만으로 의미를 확정하기 어렵습니다."
        : `열 헤더와 결정론적 타입을 근거로 ${semanticRole} 역할로 해석합니다.`,
  };
}

function inferTablePurpose(table, domain) {
  if (table.flags?.virtual) {
    return table.sourceTableId ? "DERIVED_LONG" : "CROSS_TAB";
  }
  if (table.shape?.subtotalRowCount > 0 || table.shape?.totalRowCount > 0) {
    return "SUMMARY";
  }
  if (domain === "SALES_REVENUE") return "TRANSACTION";
  if (domain === "EVENT_ATTENDANCE" || domain === "HR_PEOPLE") return "ROSTER";
  if (domain === "SURVEY_FEEDBACK") return "SURVEY_RESPONSE";
  if (asArray(table.metricFamilies).length) return "FACT";
  return "DIMENSION";
}

function mockProfileSemantics(input = {}) {
  const [primaryDomain, datasetIntent, domainConfidence, domainDescription] =
    inferDomain(input);
  const columnSemantics = [];
  const tableSemantics = [];
  const ambiguities = [];
  for (const table of asArray(input.tables)) {
    const purpose = inferTablePurpose(table, primaryDomain);
    tableSemantics.push({
      tableId: table.tableId,
      tablePurpose: purpose,
      rowGrain: table.flags?.virtual
        ? "파생 long-form 행"
        : "원본 테이블의 한 행 단위",
      confidence: table.flags?.virtual || table.flags?.primary ? 0.92 : 0.82,
      evidenceCodes: unique([
        "TABLE_STRUCTURE",
        table.sourceTableId ? "SOURCE_TABLE_LINK" : "ROLE_COMBINATION",
      ]),
      description: table.flags?.virtual
        ? "sourceTableId와 virtual 플래그를 가진 파생 테이블입니다."
        : "테이블 구조와 열 역할 조합을 기준으로 목적을 분류했습니다.",
    });
    for (const column of asArray(table.columns)) {
      const inferred = inferColumnSemantic(column);
      columnSemantics.push({
        tableId: table.tableId,
        columnId: column.columnId,
        ...inferred,
      });
      if (inferred.semanticRole === "unknown") {
        ambiguities.push({
          code: "COLUMN_MEANING_AMBIGUOUS",
          description: "열 의미를 확정할 근거가 부족합니다.",
          tableIds: [table.tableId],
          columnIds: [column.columnId],
        });
      }
    }
  }
  const tableRelations = [];
  for (const table of asArray(input.tables)) {
    if (!table.sourceTableId) continue;
    const source = asArray(input.tables).find(
      (candidate) => candidate.tableId === table.sourceTableId,
    );
    if (!source) continue;
    tableRelations.push({
      leftTableId: source.tableId,
      rightTableId: table.tableId,
      relationType: "SOURCE_DERIVATION",
      cardinality: "ONE_TO_MANY",
      leftColumnIds: [],
      rightColumnIds: [],
      confidence: 0.99,
      evidenceCodes: ["SOURCE_TABLE_LINK", "TABLE_STRUCTURE"],
      description: "명시적 sourceTableId를 가진 파생 테이블 관계입니다.",
    });
  }
  if (input.counts?.tablesTruncated || input.counts?.columnsTruncated) {
    ambiguities.push({
      code: "INPUT_TRUNCATED",
      description: "일부 테이블 또는 열이 입력 한도 때문에 제외됐습니다.",
      tableIds: [],
      columnIds: [],
    });
  }
  const unknown = primaryDomain === "UNKNOWN";
  return {
    version: QUERY_SEMANTIC_MODEL_OUTPUT_VERSION,
    classification: {
      primaryDomain,
      secondaryDomains: [],
      datasetIntent,
      confidence: domainConfidence,
      description: domainDescription,
    },
    tableSemantics,
    columnSemantics,
    tableRelations,
    ambiguities,
    requiresHumanReview:
      unknown ||
      ambiguities.some((item) => item.code !== "INPUT_TRUNCATED") ||
      domainConfidence < 0.7,
  };
}

function createMockSemanticProfilerProvider() {
  return {
    async profile({ input, model, reasoningEffort }) {
      return {
        provider: "MOCK_TEST_ONLY",
        model: model || DEFAULT_MODEL,
        reasoningEffort: reasoningEffort || DEFAULT_REASONING_EFFORT,
        responseId: `mock_${input.inputSha256.slice(0, 16)}`,
        output: mockProfileSemantics(input),
        usage: {
          input_tokens: 0,
          output_tokens: 0,
          total_tokens: 0,
          output_tokens_details: { reasoning_tokens: 0 },
        },
      };
    },
  };
}

module.exports = {
  QUERY_SEMANTIC_PROFILE_VERSION,
  QUERY_SEMANTIC_MODEL_OUTPUT_VERSION,
  QUERY_SEMANTIC_INPUT_VERSION,
  QUERY_SEMANTIC_SCHEMA_VERSION,
  QUERY_SEMANTIC_CACHE_VERSION,
  DEFAULT_MODEL,
  DEFAULT_REASONING_EFFORT,
  MAX_TABLES,
  MAX_COLUMNS,
  MODEL_OUTPUT_SCHEMA,
  MODEL_PRICING_USD_PER_MILLION,
  taxonomy,
  buildSemanticProfilerInput,
  validateSemanticProfilerModelOutput,
  normalizeUsage,
  estimateCostUsd,
  buildSemanticProfile,
  validateSemanticProfile,
  buildSemanticProfilerCacheKey,
  profileSemantics,
  businessDomainMockInputFromSemanticProfilerInput,
  inferColumnSemantic,
  mockProfileSemantics,
  createMockSemanticProfilerProvider,
};
