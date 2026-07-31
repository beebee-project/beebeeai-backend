const crypto = require("crypto");
const taxonomy = require("./queryBusinessDomainTaxonomy.json");
const { normalizeText, sha256 } = require("./queryCandidateObservation");
const {
  QUERY_BUSINESS_DOMAIN_PROMPT_VERSION,
} = require("./queryBusinessDomainPrompt");

const QUERY_BUSINESS_DOMAIN_PROFILE_VERSION = "llm_business_domain_profile_v1";
const QUERY_BUSINESS_DOMAIN_MODEL_OUTPUT_VERSION =
  "query_business_domain_model_output_v1";
const QUERY_BUSINESS_DOMAIN_INPUT_VERSION = "query_business_domain_input_v1";
const QUERY_BUSINESS_DOMAIN_SCHEMA_VERSION = "query_business_domain_schema_v1";
const QUERY_BUSINESS_DOMAIN_CACHE_VERSION = "query_business_domain_cache_v1";
const DEFAULT_MODEL = "gpt-5.6-terra";
const DEFAULT_REASONING_EFFORT = "low";
const MAX_TABLES = 20;
const MAX_COLUMNS = 240;
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
    "primaryDomain",
    "secondaryDomains",
    "datasetIntent",
    "confidence",
    "evidence",
    "ambiguities",
    "requiresHumanReview",
  ],
  properties: {
    version: {
      type: "string",
      const: QUERY_BUSINESS_DOMAIN_MODEL_OUTPUT_VERSION,
    },
    primaryDomain: {
      type: "string",
      enum: taxonomy.domains,
    },
    secondaryDomains: {
      type: "array",
      maxItems: 3,
      items: {
        type: "string",
        enum: taxonomy.domains.filter((value) => value !== "UNKNOWN"),
      },
    },
    datasetIntent: {
      type: "string",
      enum: taxonomy.datasetIntents,
    },
    confidence: {
      type: "number",
      minimum: 0,
      maximum: 1,
    },
    evidence: {
      type: "array",
      maxItems: 10,
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "tableId",
          "columnIds",
          "signalCode",
          "description",
          "strength",
        ],
        properties: {
          tableId: { type: "string" },
          columnIds: {
            type: "array",
            maxItems: 12,
            items: { type: "string" },
          },
          signalCode: {
            type: "string",
            enum: taxonomy.evidenceSignals,
          },
          description: { type: "string" },
          strength: {
            type: "number",
            minimum: 0,
            maximum: 1,
          },
        },
      },
    },
    ambiguities: {
      type: "array",
      maxItems: 10,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["code", "description", "tableIds", "columnIds"],
        properties: {
          code: {
            type: "string",
            enum: taxonomy.ambiguityCodes,
          },
          description: { type: "string" },
          tableIds: {
            type: "array",
            maxItems: 12,
            items: { type: "string" },
          },
          columnIds: {
            type: "array",
            maxItems: 20,
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
  return {
    columnId: normalizeText(column.columnId || ""),
    header: truncateText(column.normalizedHeader || column.sourceHeader || ""),
    dataType: normalizeText(column.dataType || "unknown"),
    semanticRole: normalizeText(column.semanticRole || "unknown"),
    semanticType: normalizeText(column.semanticType || "unknown"),
    metricFamily: normalizeText(column.metricFamily || ""),
    roleConfidence: round(column.roleConfidence),
    supportedOperations: unique(column.supportedOperations).sort(),
  };
}

function sanitizeTable(table = {}, remainingColumns = MAX_COLUMNS) {
  const columns = asArray(table.columns)
    .slice(0, Math.max(0, remainingColumns))
    .map(sanitizeColumn);
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
    availableRoles: unique(table.availableRoles).sort(),
    metricFamilies: unique(table.metricFamilies).sort(),
    supportedOperations: unique(table.supportedOperations).sort(),
    columns,
  };
}

function buildBusinessDomainInput({ semanticProfile } = {}) {
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
  const sourceTableCount = allTables.length;
  const sourceColumnCount = allTables.reduce(
    (sum, table) => sum + asArray(table.columns).length,
    0,
  );
  const input = {
    version: QUERY_BUSINESS_DOMAIN_INPUT_VERSION,
    source: {
      semanticProfileVersion: normalizeText(semanticProfile.version || ""),
      semanticProfileSha256: normalizeText(
        semanticProfile.profileSha256 || sha256(semanticProfile),
      ),
    },
    counts: {
      sourceTableCount,
      includedTableCount: selectedTables.length,
      sourceColumnCount,
      includedColumnCount: selectedTables.reduce(
        (sum, table) => sum + table.columns.length,
        0,
      ),
      tablesTruncated: sourceTableCount > selectedTables.length,
      columnsTruncated:
        sourceColumnCount >
        selectedTables.reduce((sum, table) => sum + table.columns.length, 0),
    },
    availableRoles: unique(semanticProfile.availableRoles).sort(),
    metricFamilies: unique(semanticProfile.metricFamilies).sort(),
    supportedOperations: unique(semanticProfile.supportedOperations).sort(),
    issueCodes: unique(
      asArray(semanticProfile.issues).map((issue) => issue?.code || ""),
    ).sort(),
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

function idSets(input = {}) {
  const tableIds = new Set();
  const columnIds = new Set();
  const columnsByTable = new Map();
  for (const table of asArray(input.tables)) {
    const tableId = normalizeText(table.tableId || "");
    if (tableId) tableIds.add(tableId);
    const tableColumns = new Set();
    for (const column of asArray(table.columns)) {
      const columnId = normalizeText(column.columnId || "");
      if (!columnId) continue;
      columnIds.add(columnId);
      tableColumns.add(columnId);
    }
    if (tableId) columnsByTable.set(tableId, tableColumns);
  }
  return { tableIds, columnIds, columnsByTable };
}

function validateBusinessDomainModelOutput(output, input) {
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
  if (output.version !== QUERY_BUSINESS_DOMAIN_MODEL_OUTPUT_VERSION) {
    errors.push({
      code: "MODEL_OUTPUT_VERSION_INVALID",
      actual: output.version,
    });
  }
  if (!taxonomy.domains.includes(output.primaryDomain)) {
    errors.push({
      code: "PRIMARY_DOMAIN_INVALID",
      actual: output.primaryDomain,
    });
  }
  const secondaryDomains = asArray(output.secondaryDomains);
  if (secondaryDomains.length > 3) {
    errors.push({ code: "SECONDARY_DOMAIN_LIMIT_EXCEEDED" });
  }
  if (new Set(secondaryDomains).size !== secondaryDomains.length) {
    errors.push({ code: "SECONDARY_DOMAIN_DUPLICATED" });
  }
  for (const domain of secondaryDomains) {
    if (!taxonomy.domains.includes(domain) || domain === "UNKNOWN") {
      errors.push({ code: "SECONDARY_DOMAIN_INVALID", actual: domain });
    }
    if (domain === output.primaryDomain) {
      errors.push({ code: "SECONDARY_DOMAIN_EQUALS_PRIMARY", actual: domain });
    }
  }
  if (!taxonomy.datasetIntents.includes(output.datasetIntent)) {
    errors.push({
      code: "DATASET_INTENT_INVALID",
      actual: output.datasetIntent,
    });
  }
  if (!Number.isFinite(Number(output.confidence))) {
    errors.push({ code: "CONFIDENCE_INVALID" });
  } else if (Number(output.confidence) < 0 || Number(output.confidence) > 1) {
    errors.push({ code: "CONFIDENCE_OUT_OF_RANGE" });
  }
  if (typeof output.requiresHumanReview !== "boolean") {
    errors.push({ code: "HUMAN_REVIEW_FLAG_INVALID" });
  }
  if (Number(output.confidence) < 0.7 && output.requiresHumanReview !== true) {
    warnings.push({ code: "LOW_CONFIDENCE_WITHOUT_HUMAN_REVIEW" });
  }
  if (
    output.primaryDomain === "UNKNOWN" &&
    output.requiresHumanReview !== true
  ) {
    warnings.push({ code: "UNKNOWN_DOMAIN_WITHOUT_HUMAN_REVIEW" });
  }

  const { tableIds, columnIds, columnsByTable } = idSets(input);
  const validateReferences = (item, index, kind) => {
    for (const tableId of unique(item?.tableIds || item?.tableId)) {
      if (!tableIds.has(tableId)) {
        errors.push({
          code: `${kind}_TABLE_ID_UNKNOWN`,
          index,
          tableId,
        });
      }
    }
    for (const columnId of unique(item?.columnIds)) {
      if (!columnIds.has(columnId)) {
        errors.push({
          code: `${kind}_COLUMN_ID_UNKNOWN`,
          index,
          columnId,
        });
      }
      if (kind === "EVIDENCE" && item?.tableId) {
        const tableColumns = columnsByTable.get(item.tableId);
        if (tableColumns && !tableColumns.has(columnId)) {
          errors.push({
            code: "EVIDENCE_COLUMN_TABLE_MISMATCH",
            index,
            tableId: item.tableId,
            columnId,
          });
        }
      }
    }
  };

  if (!Array.isArray(output.evidence)) {
    errors.push({ code: "EVIDENCE_NOT_ARRAY" });
  } else {
    output.evidence.forEach((item, index) => {
      if (!taxonomy.evidenceSignals.includes(item?.signalCode)) {
        errors.push({ code: "EVIDENCE_SIGNAL_INVALID", index });
      }
      if (!Number.isFinite(Number(item?.strength))) {
        errors.push({ code: "EVIDENCE_STRENGTH_INVALID", index });
      }
      validateReferences(item, index, "EVIDENCE");
    });
  }
  if (!Array.isArray(output.ambiguities)) {
    errors.push({ code: "AMBIGUITIES_NOT_ARRAY" });
  } else {
    output.ambiguities.forEach((item, index) => {
      if (!taxonomy.ambiguityCodes.includes(item?.code)) {
        errors.push({ code: "AMBIGUITY_CODE_INVALID", index });
      }
      validateReferences(item, index, "AMBIGUITY");
    });
  }
  if (
    output.primaryDomain !== "UNKNOWN" &&
    Array.isArray(output.evidence) &&
    output.evidence.length === 0
  ) {
    errors.push({ code: "PRIMARY_DOMAIN_EVIDENCE_REQUIRED" });
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

function buildBusinessDomainProfile({
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
  const domainInput = input || buildBusinessDomainInput({ semanticProfile });
  const validation = validateBusinessDomainModelOutput(
    modelOutput,
    domainInput,
  );
  if (!validation.valid) {
    const error = new Error("업무 영역 모델 출력 검증에 실패했습니다.");
    error.code = "BUSINESS_DOMAIN_MODEL_OUTPUT_INVALID";
    error.validation = validation;
    throw error;
  }
  const normalizedUsage = normalizeUsage(usage);
  const profile = {
    version: QUERY_BUSINESS_DOMAIN_PROFILE_VERSION,
    schemaVersion: QUERY_BUSINESS_DOMAIN_SCHEMA_VERSION,
    taxonomyVersion: taxonomy.version,
    promptVersion: QUERY_BUSINESS_DOMAIN_PROMPT_VERSION,
    source: {
      caseId: normalizeText(semanticProfile?.source?.caseId || ""),
      fileName: normalizeText(semanticProfile?.source?.fileName || ""),
      semanticProfileVersion: normalizeText(
        semanticProfile?.version ||
          domainInput.source?.semanticProfileVersion ||
          "",
      ),
      semanticProfileSha256: normalizeText(
        semanticProfile?.profileSha256 ||
          domainInput.source?.semanticProfileSha256 ||
          "",
      ),
      inputSha256: domainInput.inputSha256,
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
      includedTableCount: domainInput.counts.includedTableCount,
      includedColumnCount: domainInput.counts.includedColumnCount,
    },
    classification: {
      primaryDomain: modelOutput.primaryDomain,
      secondaryDomains: unique(modelOutput.secondaryDomains),
      datasetIntent: modelOutput.datasetIntent,
      confidence: round(clamp(modelOutput.confidence)),
      evidence: asArray(modelOutput.evidence).map((item) => ({
        tableId: normalizeText(item.tableId || ""),
        columnIds: unique(item.columnIds),
        signalCode: item.signalCode,
        description: normalizeText(item.description || ""),
        strength: round(clamp(item.strength)),
      })),
      ambiguities: asArray(modelOutput.ambiguities).map((item) => ({
        code: item.code,
        description: normalizeText(item.description || ""),
        tableIds: unique(item.tableIds),
        columnIds: unique(item.columnIds),
      })),
      requiresHumanReview: modelOutput.requiresHumanReview === true,
    },
    usage: {
      ...normalizedUsage,
      estimatedCostUsd: estimateCostUsd(normalizedUsage, model, pricing),
    },
    integrity: {
      modelOutputSha256: sha256(modelOutput),
      validationWarningCount: validation.warningCount,
      inputTableCount: domainInput.counts.includedTableCount,
      inputColumnCount: domainInput.counts.includedColumnCount,
    },
  };
  profile.profileSha256 = sha256({ ...profile, profileSha256: undefined });
  return profile;
}

function validateBusinessDomainProfile(profile) {
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
  if (profile.version !== QUERY_BUSINESS_DOMAIN_PROFILE_VERSION) {
    errors.push({ code: "PROFILE_VERSION_INVALID" });
  }
  if (profile.schemaVersion !== QUERY_BUSINESS_DOMAIN_SCHEMA_VERSION) {
    errors.push({ code: "PROFILE_SCHEMA_VERSION_INVALID" });
  }
  if (profile.taxonomyVersion !== taxonomy.version) {
    errors.push({ code: "PROFILE_TAXONOMY_VERSION_INVALID" });
  }
  if (!taxonomy.domains.includes(profile.classification?.primaryDomain)) {
    errors.push({ code: "PROFILE_PRIMARY_DOMAIN_INVALID" });
  }
  if (
    !taxonomy.datasetIntents.includes(profile.classification?.datasetIntent)
  ) {
    errors.push({ code: "PROFILE_DATASET_INTENT_INVALID" });
  }
  if (profile.privacy?.rawRowsSent !== false) {
    errors.push({ code: "PROFILE_RAW_ROWS_PRIVACY_INVALID" });
  }
  if (profile.privacy?.sampleValuesSent !== false) {
    errors.push({ code: "PROFILE_SAMPLE_VALUES_PRIVACY_INVALID" });
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

function buildBusinessDomainCacheKey({
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
    version: QUERY_BUSINESS_DOMAIN_CACHE_VERSION,
    tenantId: tenant,
    semanticProfileSha256,
    model,
    reasoningEffort,
    promptVersion: QUERY_BUSINESS_DOMAIN_PROMPT_VERSION,
    schemaVersion: QUERY_BUSINESS_DOMAIN_SCHEMA_VERSION,
    taxonomyVersion: taxonomy.version,
  };
  return crypto
    .createHmac("sha256", Buffer.from(String(cacheSecret)))
    .update(JSON.stringify(identity))
    .digest("hex");
}

async function profileBusinessDomain({
  semanticProfile,
  provider,
  cache,
  tenantId,
  cacheSecret,
  model = DEFAULT_MODEL,
  reasoningEffort = DEFAULT_REASONING_EFFORT,
  pricing,
} = {}) {
  if (!provider || typeof provider.profile !== "function") {
    throw new TypeError("provider.profile 함수가 필요합니다.");
  }
  const input = buildBusinessDomainInput({ semanticProfile });
  const cacheKey = buildBusinessDomainCacheKey({
    tenantId,
    semanticProfile,
    model,
    reasoningEffort,
    cacheSecret,
  });
  if (cache && typeof cache.get === "function") {
    const cached = await cache.get(cacheKey);
    if (cached) {
      const validation = validateBusinessDomainProfile(cached);
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
  const response = await provider.profile({
    input,
    model,
    reasoningEffort,
  });
  const profile = buildBusinessDomainProfile({
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

function mockClassifyBusinessDomain(input = {}) {
  const roleText = unique([
    ...asArray(input.availableRoles),
    ...asArray(input.tables).flatMap((table) => table.availableRoles || []),
  ])
    .join(" ")
    .toLowerCase();
  const metricText = unique([
    ...asArray(input.metricFamilies),
    ...asArray(input.tables).flatMap((table) => table.metricFamilies || []),
  ])
    .join(" ")
    .toLowerCase();
  const headerText = asArray(input.tables)
    .flatMap((table) => asArray(table.columns).map((column) => column.header))
    .join(" ")
    .toLowerCase();
  const sheetText = asArray(input.tables)
    .map((table) => table.sourceSheetName)
    .join(" ")
    .toLowerCase();
  const contextText = `${roleText} ${headerText} ${sheetText}`;
  const firstTable = asArray(input.tables)[0] || {};
  const firstColumnIds = asArray(firstTable.columns)
    .slice(0, 4)
    .map((column) => column.columnId)
    .filter(Boolean);
  let primaryDomain = "OPERATIONS_GENERAL";
  let datasetIntent = "GENERAL_TABULAR_RECORDS";
  let confidence = 0.72;
  let signalCode = "ROLE_COMBINATION";
  let description = "일반 업무 테이블의 구조와 역할 조합이 확인됩니다.";

  if (
    /sales|revenue/.test(metricText) ||
    /매출|판매|거래처|품목/.test(headerText)
  ) {
    primaryDomain = "SALES_REVENUE";
    datasetIntent = "TRANSACTION_ANALYSIS";
    confidence = 0.96;
    signalCode = "METRIC_FAMILY";
    description = "매출 metric family와 거래·품목·거래처 열 조합이 확인됩니다.";
  } else if (
    /budget|cost|amount/.test(metricText) ||
    /예산|집행|지출|비용/.test(headerText)
  ) {
    primaryDomain = "FINANCE_BUDGET";
    datasetIntent = "BUDGET_EXECUTION";
    confidence = 0.91;
    signalCode = "HEADER_SEMANTICS";
    description = "예산·집행·비용 관련 열 의미가 확인됩니다.";
  } else if (/attendance|참석|출석|신청자/.test(contextText)) {
    primaryDomain = "EVENT_ATTENDANCE";
    datasetIntent = /신청/.test(`${headerText} ${sheetText}`)
      ? "APPLICATION_TRACKING"
      : "ATTENDANCE_TRACKING";
    confidence = 0.9;
    signalCode = "HEADER_SEMANTICS";
    description = "참석자·출석·신청 관련 구조가 확인됩니다.";
  } else if (/평가|강의|교과|만족도|점수/.test(headerText)) {
    primaryDomain = "EDUCATION_EVALUATION";
    datasetIntent = "PERFORMANCE_EVALUATION";
    confidence = 0.88;
    signalCode = "HEADER_SEMANTICS";
    description = "교육·강좌·평가 관련 열 의미가 확인됩니다.";
  } else if (/설문|응답|만족도|의견/.test(headerText)) {
    primaryDomain = "SURVEY_FEEDBACK";
    datasetIntent = "SURVEY_ANALYSIS";
    confidence = 0.86;
    signalCode = "HEADER_SEMANTICS";
    description = "설문·응답·의견 관련 열 의미가 확인됩니다.";
  } else if (/재고|입고|출고|창고|물류/.test(headerText)) {
    primaryDomain = "INVENTORY_LOGISTICS";
    datasetIntent = "INVENTORY_TRACKING";
    confidence = 0.88;
    signalCode = "HEADER_SEMANTICS";
    description = "재고·입출고·물류 관련 열 의미가 확인됩니다.";
  } else if (
    /person|employee/.test(roleText) ||
    /사번|직원|인사|근무/.test(headerText)
  ) {
    primaryDomain = "HR_PEOPLE";
    datasetIntent = "ROSTER_MANAGEMENT";
    confidence = 0.8;
    signalCode = "ROLE_COMBINATION";
    description = "사람·조직 중심의 명부 구조가 확인됩니다.";
  } else if (/연구|과제|사업|성과|교수|학생연구/.test(headerText)) {
    primaryDomain = "PROJECT_RESEARCH_ADMIN";
    datasetIntent = "STATUS_TRACKING";
    confidence = 0.82;
    signalCode = "HEADER_SEMANTICS";
    description = "연구·과제·사업 관리 관련 열 의미가 확인됩니다.";
  } else if (!asArray(input.tables).length || !firstColumnIds.length) {
    primaryDomain = "UNKNOWN";
    datasetIntent = "UNKNOWN";
    confidence = 0.2;
    signalCode = "INSUFFICIENT_EVIDENCE";
    description = "업무 영역을 판단할 수 있는 테이블과 열 정보가 부족합니다.";
  }

  const unknown = primaryDomain === "UNKNOWN";
  return {
    version: QUERY_BUSINESS_DOMAIN_MODEL_OUTPUT_VERSION,
    primaryDomain,
    secondaryDomains: [],
    datasetIntent,
    confidence,
    evidence: [
      {
        tableId: normalizeText(firstTable.tableId || ""),
        columnIds: firstColumnIds,
        signalCode,
        description,
        strength: confidence,
      },
    ].filter((item) => item.tableId),
    ambiguities: unknown
      ? [
          {
            code: "INSUFFICIENT_CONTEXT",
            description: "업무 영역을 판정할 구조 정보가 부족합니다.",
            tableIds: [],
            columnIds: [],
          },
        ]
      : [],
    requiresHumanReview: unknown || confidence < 0.7,
  };
}

function createMockBusinessDomainProvider() {
  return {
    async profile({ input, model, reasoningEffort }) {
      return {
        provider: "MOCK_TEST_ONLY",
        model: model || DEFAULT_MODEL,
        reasoningEffort: reasoningEffort || DEFAULT_REASONING_EFFORT,
        responseId: `mock_${input.inputSha256.slice(0, 16)}`,
        output: mockClassifyBusinessDomain(input),
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
  QUERY_BUSINESS_DOMAIN_PROFILE_VERSION,
  QUERY_BUSINESS_DOMAIN_MODEL_OUTPUT_VERSION,
  QUERY_BUSINESS_DOMAIN_INPUT_VERSION,
  QUERY_BUSINESS_DOMAIN_SCHEMA_VERSION,
  QUERY_BUSINESS_DOMAIN_CACHE_VERSION,
  DEFAULT_MODEL,
  DEFAULT_REASONING_EFFORT,
  MODEL_OUTPUT_SCHEMA,
  MODEL_PRICING_USD_PER_MILLION,
  taxonomy,
  buildBusinessDomainInput,
  validateBusinessDomainModelOutput,
  normalizeUsage,
  estimateCostUsd,
  buildBusinessDomainProfile,
  validateBusinessDomainProfile,
  buildBusinessDomainCacheKey,
  profileBusinessDomain,
  mockClassifyBusinessDomain,
  createMockBusinessDomainProvider,
};
