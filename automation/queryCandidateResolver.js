"use strict";

const {
  normalizeText,
  sha256,
} = require("./queryCandidateObservation");
const {
  assessCandidate,
} = require("./queryCandidateRetriever");

const QUERY_CANDIDATE_RESOLUTION_VERSION = "query_candidate_resolution_v1";
const QUERY_CANDIDATE_RESOLUTION_ITEM_VERSION = "query_candidate_resolution_item_v1";
const PREVIOUS_QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION =
  "deterministic_candidate_resolution_policy_v1_1";
const QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION =
  "deterministic_candidate_resolution_policy_v1_2";

const RESOLUTION_RESULT = Object.freeze([
  "RESOLVED",
  "STILL_DEFERRED",
  "EXCLUDED",
]);
const CHECK_STATUS = Object.freeze(["PASS", "FAIL", "UNKNOWN", "NOT_APPLICABLE"]);
const REASON_LEVEL = Object.freeze(["INFO", "WARNING", "BLOCKING"]);


const DOMAIN_EVIDENCE_TOKEN_MAP = Object.freeze({
  SALES_REVENUE: [
    "sales", "revenue", "sales_amount", "product_sales",
    "매출", "판매금액", "매출액", "판매수량",
  ],
  FINANCE_BUDGET: [
    "budget", "expense", "expenditure", "labor_cost", "budget_execution",
    "예산", "집행", "지출", "비용", "인건비",
  ],
  EVENT_ATTENDANCE: [
    "attendance", "attendee", "applicant", "application_status", "event_participant",
    "출석", "참석", "신청자", "참가자", "참석여부", "참가상태",
  ],
  EDUCATION_EVALUATION: [
    "course", "lecture", "instructor", "education_evaluation", "evaluation_score",
    "강좌", "강의", "강사", "수업", "교육평가", "평가점수",
  ],
  HR_PEOPLE: [
    "employee", "staff", "personnel", "human_resource",
    "직원", "인사", "근로자",
  ],
  INVENTORY_LOGISTICS: [
    "inventory", "warehouse", "shipment", "stock_quantity",
    "재고", "창고", "입출고", "배송",
  ],
  PROJECT_RESEARCH_ADMIN: [
    "project", "research", "grant", "research_project",
    "프로젝트", "연구", "과제", "지원사업",
  ],
  SURVEY_FEEDBACK: [
    "survey", "feedback", "satisfaction", "questionnaire", "recommendation_score",
    "설문", "피드백", "만족도", "의견", "추천점수",
  ],
  CUSTOMER_VENDOR: [
    "vendor", "supplier", "customer_inquiry", "client_inquiry",
    "거래처", "공급업체", "고객문의", "문의유형",
  ],
});


const NAMED_TEMPLATE_AUXILIARY_DOMAINS = new Set([
  "SURVEY_FEEDBACK",
]);

const STRUCTURAL_GENERIC_RECIPE_IDS = new Set([
  "single_source_dashboard",
  "multi_source_schema_union",
  "time_sum",
  "time_avg",
  "group_sum",
  "group_avg",
  "group_summary",
  "composition_ratio",
  "cumulative_sum",
  "top_bottom",
  "cross_sum",
  "cross_count",
  "category_count",
  "count_rows",
].map(normalizeLoose));


const RECIPE_OPERAND_SPECS = Object.freeze({
  timesum: { operation: "time_sum", operands: ["period", "measure"] },
  timeavg: { operation: "time_avg", operands: ["period", "measure"] },
  cumulativesum: { operation: "cumulative_sum", operands: ["period", "measure"] },
  groupsum: { operation: "group_sum", operands: ["group", "measure"] },
  groupavg: { operation: "group_avg", operands: ["group", "measure"] },
  groupsummary: { operation: "group_summary", operands: ["group", "measure"] },
  compositionratio: { operation: "composition_ratio", operands: ["group", "measure"] },
  topbottom: { operation: "top_bottom", operands: ["group", "measure"] },
  categorycount: { operation: "category_count", operands: ["group"] },
  crosssum: { operation: "cross_sum", operands: ["dimension", "dimension", "measure"] },
  crosscount: { operation: "cross_count", operands: ["dimension", "dimension"] },
});

const GENERIC_PERIOD_OPERANDS = new Set([
  "기간", "일자", "날짜", "연월", "월", "년월", "date", "period", "month", "yearmonth",
].map(normalizeLoose));

const DOMAIN_TOKEN_MAP = Object.freeze({
  SALES_REVENUE: [
    "sales", "revenue", "selling", "product_sales", "customer_sales",
    "매출", "판매", "매상",
  ],
  FINANCE_BUDGET: [
    "budget", "finance", "expense", "expenditure", "cost", "spending",
    "예산", "집행", "지출", "비용", "재정",
  ],
  EVENT_ATTENDANCE: [
    "attendance", "attendee", "participant", "applicant", "application",
    "event", "workshop", "출석", "참석", "신청", "행사", "명단",
  ],
  EDUCATION_EVALUATION: [
    "course", "education", "evaluation", "performance_evaluation", "class",
    "lecture", "교육", "강좌", "강의", "수업", "평가", "교과",
  ],
  HR_PEOPLE: [
    "employee", "staff", "personnel", "human_resource", "hr", "인사", "직원", "근로자",
  ],
  INVENTORY_LOGISTICS: [
    "inventory", "stock", "warehouse", "logistics", "shipment", "재고", "물류", "입출고",
  ],
  PROJECT_RESEARCH_ADMIN: [
    "project", "research", "grant", "task", "프로젝트", "과제", "연구", "사업",
  ],
  SURVEY_FEEDBACK: [
    "survey", "feedback", "satisfaction", "questionnaire", "설문", "만족도", "의견",
  ],
  CUSTOMER_VENDOR: [
    "customer", "client", "vendor", "supplier", "고객", "거래처", "공급업체", "업체",
  ],
});

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const result = [];
  const seen = new Set();
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

function candidateMap(items = []) {
  const map = new Map();
  for (const item of asArray(items)) {
    const candidateId = normalizeText(item?.candidateId || "");
    if (candidateId && !map.has(candidateId)) map.set(candidateId, item);
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

function eligiblePhysicalRoots(profile = {}) {
  return unique(eligibleTables(profile).map(physicalRootId));
}

function tablesForRoot(profile = {}, rootId = "") {
  const key = normalizeLoose(rootId);
  return eligibleTables(profile).filter(
    (table) => normalizeLoose(physicalRootId(table)) === key,
  );
}

function synthesizeCandidate(retrievalItem = {}, sourceTableIds = undefined) {
  return {
    version: normalizeText(retrievalItem.provenance?.candidateItemVersion || ""),
    candidateId: normalizeText(retrievalItem.candidateId || ""),
    recipeId: normalizeText(retrievalItem.recipeId || ""),
    templateId: normalizeText(retrievalItem.templateId || ""),
    candidateType: normalizeText(retrievalItem.candidateType || "UNKNOWN"),
    observedClass: normalizeText(retrievalItem.observedClass || "UNKNOWN"),
    visibility: normalizeText(retrievalItem.visibility || "VISIBLE"),
    rank: Number.isInteger(retrievalItem.originalRank)
      ? retrievalItem.originalRank
      : null,
    score: Number.isFinite(Number(retrievalItem.originalScore))
      ? Number(retrievalItem.originalScore)
      : null,
    sourceTableIds:
      sourceTableIds === undefined
        ? unique(retrievalItem.sourceTableIds)
        : unique(sourceTableIds),
    status: normalizeText(retrievalItem.provenance?.candidateStatus || "UNASSESSED"),
  };
}

function hasBlockingReasons(assessment = {}) {
  return asArray(assessment.reasons).some((item) => item.level === "BLOCKING");
}

function rootAssessment(retrievalItem, capability, profile, rootId) {
  const candidate = synthesizeCandidate(retrievalItem, [rootId]);
  const assessment = assessCandidate(candidate, capability, profile);
  return {
    rootId,
    assessment,
    viable:
      assessment.checks?.sourceScope?.status === "PASS" &&
      !hasBlockingReasons(assessment),
  };
}

function enhancedSourceScope(retrievalItem = {}, capability = {}, profile = {}) {
  const explicitIds = unique(retrievalItem.sourceTableIds);
  if (explicitIds.length) {
    const assessment = assessCandidate(
      synthesizeCandidate(retrievalItem),
      capability,
      profile,
    );
    return {
      status: assessment.checks?.sourceScope?.status || "UNKNOWN",
      mode: assessment.checks?.sourceScope?.mode || "EXPLICIT_SOURCE_UNRESOLVED",
      selectedRootIds: unique(assessment.matchedPhysicalTableIds),
      requestedSourceTableIds: explicitIds,
      assessment,
      rootAssessments: [],
      reasonCode:
        assessment.checks?.sourceScope?.status === "PASS"
          ? "EXPLICIT_SOURCE_RESOLVED"
          : "EXPLICIT_SOURCE_UNRESOLVED",
    };
  }

  const roots = eligiblePhysicalRoots(profile);
  if (!roots.length) {
    const assessment = assessCandidate(
      synthesizeCandidate(retrievalItem),
      capability,
      profile,
    );
    return {
      status: "FAIL",
      mode: "NO_ELIGIBLE_SOURCE",
      selectedRootIds: [],
      requestedSourceTableIds: [],
      assessment,
      rootAssessments: [],
      reasonCode: "NO_ANALYSIS_ELIGIBLE_TABLE",
    };
  }

  if (roots.length === 1) {
    const only = rootAssessment(retrievalItem, capability, profile, roots[0]);
    return {
      status: "PASS",
      mode: "SINGLE_PHYSICAL_SOURCE",
      selectedRootIds: [roots[0]],
      requestedSourceTableIds: [],
      assessment: only.assessment,
      rootAssessments: [only],
      reasonCode: "SINGLE_PHYSICAL_SOURCE_SELECTED",
    };
  }

  const assessed = roots.map((rootId) =>
    rootAssessment(retrievalItem, capability, profile, rootId),
  );
  const viable = assessed.filter((item) => item.viable);
  if (viable.length === 1) {
    return {
      status: "PASS",
      mode: "SEMANTIC_UNIQUE_PHYSICAL_SOURCE",
      selectedRootIds: [viable[0].rootId],
      requestedSourceTableIds: [],
      assessment: viable[0].assessment,
      rootAssessments: assessed,
      reasonCode: "SEMANTIC_SOURCE_UNIQUELY_RESOLVED",
    };
  }

  if (viable.length > 1) {
    return {
      status: "UNKNOWN",
      mode: "MULTIPLE_SEMANTIC_SOURCE_MATCHES",
      selectedRootIds: viable.map((item) => item.rootId),
      requestedSourceTableIds: [],
      assessment: viable
        .slice()
        .sort(
          (a, b) =>
            Number(b.assessment.retrievalScore || 0) -
            Number(a.assessment.retrievalScore || 0),
        )[0].assessment,
      rootAssessments: assessed,
      reasonCode: "MULTIPLE_ELIGIBLE_TABLES_SEMANTICALLY_MATCH",
    };
  }

  const best = assessed
    .slice()
    .sort(
      (a, b) =>
        Number(b.assessment.retrievalScore || 0) -
        Number(a.assessment.retrievalScore || 0),
    )[0];
  return {
    status: "UNKNOWN",
    mode: "NO_SEMANTIC_SOURCE_MATCH",
    selectedRootIds: [],
    requestedSourceTableIds: [],
    assessment: best.assessment,
    rootAssessments: assessed,
    reasonCode: "NO_PHYSICAL_SOURCE_SATISFIES_REQUIREMENTS",
  };
}

function domainsFromText(value = "", tokenMap = DOMAIN_TOKEN_MAP) {
  const parts = asArray(value)
    .flatMap((item) => String(item || "").split(/[^가-힣A-Za-z0-9]+/gu))
    .map(normalizeLoose)
    .filter(Boolean);
  const matches = [];
  for (const [domain, tokens] of Object.entries(tokenMap)) {
    if (tokens.some((token) => {
      const normalizedToken = normalizeLoose(token);
      return normalizedToken && parts.some((part) => part.includes(normalizedToken));
    })) {
      matches.push(domain);
    }
  }
  return unique(matches);
}

function candidateDomainSignals(retrievalItem = {}, capability = {}) {
  const strongDomains = domainsFromText([
    retrievalItem.candidateId,
    retrievalItem.templateId,
    capability.templateId,
    ...asArray(capability.matchedTemplateIds),
  ]);
  const weakDomains = domainsFromText([
    retrievalItem.recipeId,
    capability.recipeId,
    capability.bindingKey,
    ...asArray(capability.recipeIds),
    ...asArray(capability.metricFamilies),
    ...asArray(capability.supportedMetricIds),
  ]);
  const bindingStatus = normalizeText(
    capability.bindingStatus || retrievalItem.bindingStatus || "UNBOUND",
  ).toUpperCase();
  const expectedDomains = strongDomains.length
    ? strongDomains
    : ["BOUND", "PARTIAL"].includes(bindingStatus)
      ? weakDomains
      : [];
  const templateId = normalizeText(
    retrievalItem.templateId || capability.templateId || "",
  );
  const coreStrongDomains = strongDomains.filter(
    (domain) => !NAMED_TEMPLATE_AUXILIARY_DOMAINS.has(domain),
  );
  const primaryAnchorDomains = templateId && strongDomains.length
    ? coreStrongDomains.length
      ? coreStrongDomains
      : strongDomains
    : [];
  return {
    expectedDomains,
    strongDomains,
    weakDomains,
    templateId,
    primaryAnchorDomains,
    primaryAnchorRequired: Boolean(templateId && primaryAnchorDomains.length),
    signalStrength: strongDomains.length
      ? "STRONG_IDENTITY"
      : expectedDomains.length
        ? "BOUND_CAPABILITY"
        : "NONE",
  };
}

function semanticEvidenceDomains(profile = {}) {
  const parts = [];
  for (const table of asArray(profile.tables)) {
    parts.push(table.purpose, table.normalizedPurpose);
    for (const column of asArray(table.columns)) {
      parts.push(
        column.sourceHeader,
        column.normalizedHeader,
        column.normalizedMeaning,
        column.semanticRole,
        column.semanticType,
        column.metricFamily,
        column.unitSemantic,
        ...asArray(column.roleAliases),
        ...asArray(column.metricAliases),
      );
    }
  }
  return domainsFromText(parts, DOMAIN_EVIDENCE_TOKEN_MAP);
}

function isStructuralGenericCandidate(retrievalItem = {}, capability = {}) {
  const templateId = normalizeText(
    retrievalItem.templateId || capability.templateId || "",
  );
  if (templateId) return false;
  const candidateId = normalizeLoose(retrievalItem.candidateId || "");
  const recipeIds = unique([
    retrievalItem.recipeId,
    capability.recipeId,
    ...asArray(capability.recipeIds),
  ]).map(normalizeLoose);
  if (candidateId.startsWith("multisource")) return true;
  return recipeIds.some((recipeId) =>
    STRUCTURAL_GENERIC_RECIPE_IDS.has(recipeId),
  );
}


function identifierSegments(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .split(/_+/u)
    .map((item) => normalizeLoose(item))
    .filter(Boolean);
}

function parsedRecipeOperandSpec(retrievalItem = {}, capability = {}) {
  const identifiers = unique([
    retrievalItem.candidateId,
    retrievalItem.recipeId,
    capability.recipeId,
    ...asArray(capability.recipeIds),
  ]);

  for (const identifier of identifiers) {
    const segments = identifierSegments(identifier);
    for (let index = 0; index < segments.length; index += 1) {
      const one = segments[index];
      const two = index + 1 < segments.length
        ? `${segments[index]}${segments[index + 1]}`
        : "";
      const key = RECIPE_OPERAND_SPECS[one]
        ? one
        : RECIPE_OPERAND_SPECS[two]
          ? two
          : "";
      if (!key) continue;
      const spec = RECIPE_OPERAND_SPECS[key];
      const operandStart = key === two ? index + 2 : index + 1;
      const values = segments.slice(operandStart);
      if (values.length < spec.operands.length) continue;
      return {
        status: "REQUIRED",
        operation: spec.operation,
        identifier,
        operands: spec.operands.map((kind, operandIndex) => ({
          kind,
          expectedToken: values[operandIndex],
        })),
      };
    }
  }

  return {
    status: "NOT_APPLICABLE",
    operation: "",
    identifier: "",
    operands: [],
  };
}

function selectedScopeTables(profile = {}, scope = {}) {
  const matchedIds = new Set(unique(scope.assessment?.matchedTableIds).map(normalizeLoose));
  if (matchedIds.size) {
    return eligibleTables(profile).filter((table) =>
      matchedIds.has(normalizeLoose(table.tableId)),
    );
  }
  const rootIds = unique(scope.selectedRootIds);
  if (rootIds.length) {
    const result = [];
    const seen = new Set();
    for (const rootId of rootIds) {
      for (const table of tablesForRoot(profile, rootId)) {
        const key = normalizeLoose(table.tableId);
        if (!key || seen.has(key)) continue;
        seen.add(key);
        result.push(table);
      }
    }
    return result;
  }
  return [];
}

function columnHeaderTokens(column = {}) {
  return unique([
    column.sourceHeader,
    column.normalizedHeader,
    column.normalizedMeaning,
  ]).map(normalizeLoose).filter(Boolean);
}

function measureStem(value = "") {
  const normalized = normalizeLoose(value);
  return normalized.replace(/(?:금액|점수|수량|지표값|값|액)$/u, "");
}

function tokenMatchesHeader(expectedToken = "", headerToken = "", kind = "") {
  const expected = normalizeLoose(expectedToken);
  const header = normalizeLoose(headerToken);
  if (!expected || !header) return false;
  if (expected === header) return true;
  if (kind === "measure") {
    const expectedStem = measureStem(expected);
    const headerStem = measureStem(header);
    if (
      expectedStem &&
      headerStem &&
      expectedStem.length >= 2 &&
      expectedStem === headerStem
    ) {
      return true;
    }
  }
  if (Math.min(expected.length, header.length) < 3) return false;
  return expected.includes(header) || header.includes(expected);
}

function columnOperandEvidence(column = {}, kind = "", expectedToken = "", mode = "EXACT_HEADER") {
  return {
    columnId: normalizeText(column.columnId || ""),
    columnName: normalizeText(column.sourceHeader || column.normalizedHeader || ""),
    expectedToken: normalizeText(expectedToken || ""),
    operandKind: kind,
    matchMode: mode,
    semanticRole: normalizeText(column.semanticRole || ""),
    dataType: normalizeText(column.dataType || ""),
    semanticType: normalizeText(column.semanticType || ""),
    metricFamily: normalizeText(column.metricFamily || ""),
  };
}

function semanticPeriodColumns(columns = []) {
  return asArray(columns).filter((column) => {
    const roles = unique([
      column.semanticRole,
      ...asArray(column.roleAliases),
    ]).map((item) => normalizeLoose(item));
    const dataType = normalizeLoose(column.dataType || "");
    return roles.some((role) => ["period", "date", "time", "month", "yearmonth"].includes(role)) ||
      ["date", "datetime", "period"].includes(dataType);
  });
}

function matchOperandColumns(columns = [], operand = {}) {
  const expectedToken = normalizeLoose(operand.expectedToken || "");
  const exact = asArray(columns).filter((column) =>
    columnHeaderTokens(column).some((header) =>
      tokenMatchesHeader(expectedToken, header, operand.kind),
    ),
  );
  if (exact.length) {
    return {
      status: "PASS",
      kind: operand.kind,
      expectedToken,
      matchMode: "EXACT_OR_COMPATIBLE_HEADER",
      matched: exact.map((column) =>
        columnOperandEvidence(column, operand.kind, expectedToken, "EXACT_OR_COMPATIBLE_HEADER"),
      ),
    };
  }

  if (operand.kind === "period" && GENERIC_PERIOD_OPERANDS.has(expectedToken)) {
    const semantic = semanticPeriodColumns(columns);
    if (semantic.length === 1) {
      return {
        status: "PASS",
        kind: operand.kind,
        expectedToken,
        matchMode: "UNIQUE_SEMANTIC_PERIOD",
        matched: semantic.map((column) =>
          columnOperandEvidence(column, operand.kind, expectedToken, "UNIQUE_SEMANTIC_PERIOD"),
        ),
      };
    }
    if (semantic.length > 1) {
      return {
        status: "UNKNOWN",
        kind: operand.kind,
        expectedToken,
        matchMode: "MULTIPLE_SEMANTIC_PERIOD_COLUMNS",
        matched: semantic.map((column) =>
          columnOperandEvidence(column, operand.kind, expectedToken, "MULTIPLE_SEMANTIC_PERIOD_COLUMNS"),
        ),
      };
    }
  }

  return {
    status: "FAIL",
    kind: operand.kind,
    expectedToken,
    matchMode: "OPERAND_COLUMN_NOT_FOUND",
    matched: [],
  };
}

function recipeOperandBindingCheck(
  retrievalItem = {},
  capability = {},
  profile = {},
  scope = {},
) {
  const parsed = parsedRecipeOperandSpec(retrievalItem, capability);
  if (parsed.status === "NOT_APPLICABLE") {
    return {
      status: "NOT_APPLICABLE",
      operation: "",
      identifier: "",
      operands: [],
      matchedColumnIds: [],
      reasonCode: "NO_EXPLICIT_RECIPE_OPERANDS",
    };
  }

  const tables = selectedScopeTables(profile, scope);
  if (!tables.length) {
    return {
      status: "UNKNOWN",
      operation: parsed.operation,
      identifier: parsed.identifier,
      operands: parsed.operands.map((operand) => ({
        ...operand,
        status: "UNKNOWN",
        matchMode: "SOURCE_SCOPE_NOT_RESOLVED",
        matched: [],
      })),
      matchedColumnIds: [],
      reasonCode: "OPERAND_SOURCE_SCOPE_NOT_RESOLVED",
    };
  }

  const columns = tables.flatMap((table) => asArray(table.columns));
  const operands = parsed.operands.map((operand) =>
    matchOperandColumns(columns, operand),
  );
  const failed = operands.filter((operand) => operand.status === "FAIL");
  const ambiguous = operands.filter((operand) => operand.status === "UNKNOWN");
  const status = failed.length
    ? "FAIL"
    : ambiguous.length
      ? "UNKNOWN"
      : "PASS";
  return {
    status,
    operation: parsed.operation,
    identifier: parsed.identifier,
    operands,
    matchedColumnIds: unique(
      operands.flatMap((operand) =>
        asArray(operand.matched).map((match) => match.columnId),
      ),
    ),
    reasonCode: status === "PASS"
      ? "RECIPE_OPERANDS_BOUND"
      : status === "UNKNOWN"
        ? "RECIPE_OPERAND_BINDING_AMBIGUOUS"
        : "RECIPE_OPERAND_BINDING_NOT_CONFIRMED",
  };
}

function mergeEvidence(...groups) {
  const result = [];
  const seen = new Set();
  for (const item of groups.flatMap((group) => asArray(group))) {
    const columnId = normalizeText(item?.columnId || "");
    const key = columnId || JSON.stringify(item || {});
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(item);
  }
  return result;
}

function domainAlignmentCheck(retrievalItem = {}, capability = {}, profile = {}) {
  const signals = candidateDomainSignals(retrievalItem, capability);
  const expectedDomains = signals.expectedDomains;
  const classification = profile.classification || {};
  const primaryDomain = normalizeText(classification.primaryDomain || "UNKNOWN");
  const declaredDomains = unique([
    primaryDomain,
    ...asArray(classification.secondaryDomains),
  ]).filter((domain) => domain && domain !== "UNKNOWN");
  const evidenceDomains = semanticEvidenceDomains(profile);
  const actualDomains = unique([...declaredDomains, ...evidenceDomains]);
  const confidence = Number(classification.confidence || 0);
  const matchedDomains = expectedDomains.filter((domain) =>
    actualDomains.includes(domain),
  );
  const primaryAnchorDomains = signals.primaryAnchorDomains || [];
  const primaryAnchorRequired = signals.primaryAnchorRequired === true;
  const primaryAnchorMatched =
    !primaryAnchorRequired || primaryAnchorDomains.includes(primaryDomain);

  const base = {
    expectedDomains,
    actualDomains,
    declaredDomains,
    evidenceDomains,
    matchedDomains,
    primaryDomain,
    primaryAnchorDomains,
    primaryAnchorRequired,
    primaryAnchorMatched,
    signalStrength: signals.signalStrength,
    confidence,
  };

  if (!expectedDomains.length) {
    return {
      ...base,
      status: "NOT_APPLICABLE",
      expectedDomains: [],
      matchedDomains: [],
      reasonCode: "NO_DOMAIN_SIGNAL_IN_CANDIDATE",
    };
  }
  if (!declaredDomains.length || confidence < 0.8) {
    return {
      ...base,
      status: "UNKNOWN",
      reasonCode: "DATASET_DOMAIN_NOT_CONFIDENT",
    };
  }
  if (!primaryAnchorMatched) {
    return {
      ...base,
      status: "FAIL",
      reasonCode: "NAMED_TEMPLATE_PRIMARY_DOMAIN_CONFLICT",
    };
  }
  if (matchedDomains.length) {
    const matchedByDeclared = matchedDomains.some((domain) =>
      declaredDomains.includes(domain),
    );
    return {
      ...base,
      status: "PASS",
      reasonCode: matchedByDeclared
        ? "CANDIDATE_DOMAIN_MATCHED"
        : "CANDIDATE_DOMAIN_MATCHED_BY_SEMANTIC_EVIDENCE",
    };
  }
  if (signals.signalStrength !== "STRONG_IDENTITY") {
    return {
      ...base,
      status: "UNKNOWN",
      matchedDomains: [],
      reasonCode: "WEAK_DOMAIN_SIGNAL_NOT_CONCLUSIVE",
    };
  }
  return {
    ...base,
    status: "FAIL",
    matchedDomains: [],
    reasonCode: "CANDIDATE_DOMAIN_CONFLICT",
  };
}

function executorSupportCheck(retrievalItem = {}, capability = {}) {
  const recipePresent = Boolean(normalizeText(
    retrievalItem.recipeId || capability.recipeId || asArray(capability.recipeIds)[0] || "",
  ));
  const declaredStatus = normalizeText(
    capability.executorSupport?.status || "UNKNOWN",
  ).toUpperCase();
  const outputTypes = unique(capability.executorSupport?.outputTypes);

  if (!recipePresent) {
    return {
      status: "UNKNOWN",
      declaredStatus,
      recipePresent: false,
      outputTypes,
      mode: "RECIPE_NOT_BOUND",
    };
  }
  if (declaredStatus === "DECLARED") {
    return {
      status: "PASS",
      declaredStatus,
      recipePresent: true,
      outputTypes,
      mode: "DECLARED_EXECUTOR",
    };
  }
  if (declaredStatus === "GENERIC" && outputTypes.length) {
    return {
      status: "PASS",
      declaredStatus,
      recipePresent: true,
      outputTypes,
      mode: "GENERIC_EXECUTOR_REQUIRES_FEASIBILITY_GATE",
    };
  }
  return {
    status: "UNKNOWN",
    declaredStatus,
    recipePresent: true,
    outputTypes,
    mode: "EXECUTOR_SUPPORT_UNRESOLVED",
  };
}

function resolutionReason(code, level, message, details = {}) {
  return { code, level, message, details };
}

function sourceCheckFrom(scope = {}) {
  const assessment = scope.assessment || {};
  return {
    status: scope.status,
    mode: scope.mode,
    selectedRootIds: unique(scope.selectedRootIds),
    requestedSourceTableIds: unique(scope.requestedSourceTableIds),
    matchedTableIds: unique(assessment.matchedTableIds),
    matchedPhysicalTableIds: unique(assessment.matchedPhysicalTableIds),
    reasonCode: normalizeText(scope.reasonCode || ""),
    rootCandidates: asArray(scope.rootAssessments).map((item) => ({
      physicalRootId: normalizeText(item.rootId || ""),
      viable: item.viable === true,
      score: Number(item.assessment?.retrievalScore || 0),
      blockingReasonCodes: unique(
        asArray(item.assessment?.reasons)
          .filter((reason) => reason.level === "BLOCKING")
          .map((reason) => reason.code),
      ),
    })),
  };
}

function scoreResolution({ assessment = {}, domain = {}, executor = {}, scope = {} }) {
  let score = Number(assessment.retrievalScore || 0);
  if (domain.status === "PASS") score += 5;
  if (domain.status === "FAIL") score -= 25;
  if (executor.status === "PASS") score += 3;
  if (scope.mode === "SEMANTIC_UNIQUE_PHYSICAL_SOURCE") score += 4;
  if (scope.status === "UNKNOWN") score -= 15;
  return Number(Math.max(0, Math.min(100, score)).toFixed(6));
}

function carriedItem(retrievalItem = {}, result = "RESOLVED") {
  const reasonCode =
    result === "RESOLVED"
      ? "PRIOR_RETRIEVED_CARRIED"
      : "PRIOR_EXCLUDED_CARRIED";
  const reasonLevel = result === "RESOLVED" ? "INFO" : "BLOCKING";
  const reasonMessage =
    result === "RESOLVED"
      ? "패치 5에서 이미 결정론적으로 통과한 후보를 유지했습니다."
      : "패치 5에서 명백한 필수조건 누락으로 제외된 후보를 보수적으로 유지했습니다.";
  const item = {
    version: QUERY_CANDIDATE_RESOLUTION_ITEM_VERSION,
    candidateId: normalizeText(retrievalItem.candidateId || ""),
    recipeId: normalizeText(retrievalItem.recipeId || ""),
    templateId: normalizeText(retrievalItem.templateId || ""),
    candidateType: normalizeText(retrievalItem.candidateType || "UNKNOWN"),
    result,
    previousRetrievalResult: normalizeText(retrievalItem.result || ""),
    bindingStatus: normalizeText(retrievalItem.bindingStatus || "UNBOUND"),
    bindingSource: normalizeText(retrievalItem.bindingSource || "NONE"),
    originalRank: Number.isInteger(retrievalItem.originalRank)
      ? retrievalItem.originalRank
      : null,
    originalScore: Number.isFinite(Number(retrievalItem.originalScore))
      ? Number(retrievalItem.originalScore)
      : null,
    resolutionScore: Number(retrievalItem.retrievalScore || 0),
    sourceTableIds: unique(retrievalItem.sourceTableIds),
    matchedTableIds: unique(retrievalItem.matchedTableIds),
    matchedPhysicalTableIds: unique(retrievalItem.matchedPhysicalTableIds),
    matchedColumnIds: unique(retrievalItem.matchedColumnIds),
    checks: {
      sourceScope: retrievalItem.checks?.sourceScope || {
        status: "UNKNOWN",
        mode: "NOT_REASSESSED",
      },
      domainAlignment: {
        status: "NOT_APPLICABLE",
        expectedDomains: [],
        actualDomains: [],
        confidence: 0,
        reasonCode: "TERMINAL_PRIOR_RESULT",
      },
      executorSupport: retrievalItem.checks?.executorSupport || {
        status: "UNKNOWN",
        declaredStatus: "UNKNOWN",
        recipePresent: Boolean(retrievalItem.recipeId),
        outputTypes: [],
        mode: "NOT_REASSESSED",
      },
      operandBinding: retrievalItem.checks?.operandBinding || {
        status: "NOT_APPLICABLE",
        operation: "",
        identifier: "",
        operands: [],
        matchedColumnIds: [],
        reasonCode: "TERMINAL_PRIOR_RESULT",
      },
      requiredRoles: asArray(retrievalItem.checks?.requiredRoles),
      requiredCapabilities: asArray(retrievalItem.checks?.requiredCapabilities),
      metricFamily: retrievalItem.checks?.metricFamily || {
        status: "NOT_APPLICABLE",
        required: [],
        available: [],
      },
      constraints: asArray(retrievalItem.checks?.constraints),
    },
    reasons: [
      resolutionReason(reasonCode, reasonLevel, reasonMessage, {}),
      ...asArray(retrievalItem.reasons),
    ],
    missingRequirements: asArray(retrievalItem.missingRequirements),
    evidence: asArray(retrievalItem.evidence),
    provenance: {
      retrievalItemVersion: normalizeText(retrievalItem.version || ""),
      retrievalItemSha256: normalizeText(retrievalItem.retrievalItemSha256 || ""),
      candidateStatus: normalizeText(
        retrievalItem.provenance?.candidateStatus || "UNASSESSED",
      ),
      terminalPriorResult: true,
      semanticReassessmentPerformed: false,
    },
  };
  item.resolutionItemSha256 = sha256({ ...item, resolutionItemSha256: undefined });
  return item;
}

function resolveDeferredCandidate(
  retrievalItem = {},
  capability = {},
  resolvedProfile = {},
) {
  const reasons = [];
  const bindingStatus = normalizeText(
    capability.bindingStatus || retrievalItem.bindingStatus || "UNBOUND",
  ).toUpperCase();

  if (bindingStatus === "UNBOUND") {
    reasons.push(resolutionReason(
      "CAPABILITY_BINDING_UNBOUND",
      "WARNING",
      "후보 capability가 manifest에 연결되지 않아 의미 검증을 완료할 수 없습니다.",
      {},
    ));
    return buildResolvedItem({
      retrievalItem,
      capability,
      resolvedProfile,
      result: "STILL_DEFERRED",
      reasons,
      scope: null,
      assessment: null,
      domain: domainAlignmentCheck(retrievalItem, capability, resolvedProfile),
      executor: executorSupportCheck(retrievalItem, capability),
      operandBinding: null,
    });
  }

  const scope = enhancedSourceScope(retrievalItem, capability, resolvedProfile);
  const assessment = scope.assessment;
  const domain = domainAlignmentCheck(retrievalItem, capability, resolvedProfile);
  const executor = executorSupportCheck(retrievalItem, capability);
  const operandBinding = recipeOperandBindingCheck(
    retrievalItem,
    capability,
    resolvedProfile,
    scope,
  );

  if (scope.status === "FAIL") {
    reasons.push(resolutionReason(
      "NO_ANALYSIS_ELIGIBLE_TABLE",
      "BLOCKING",
      "분석 가능한 테이블이 없어 후보를 생성할 수 없습니다.",
      {},
    ));
  } else if (scope.status === "UNKNOWN") {
    reasons.push(resolutionReason(
      scope.reasonCode || "SOURCE_SCOPE_STILL_AMBIGUOUS",
      "WARNING",
      "병합된 의미를 적용해도 후보가 사용할 원본 테이블을 하나로 확정할 수 없습니다.",
      { selectedRootIds: scope.selectedRootIds },
    ));
  } else if (scope.mode === "SEMANTIC_UNIQUE_PHYSICAL_SOURCE") {
    reasons.push(resolutionReason(
      "SOURCE_SCOPE_RESOLVED_BY_SEMANTICS",
      "INFO",
      "필수 역할과 capability 충족도를 이용해 원본 테이블을 단일하게 확정했습니다.",
      { selectedRootIds: scope.selectedRootIds },
    ));
  }

  if (domain.status === "FAIL") {
    const namedPrimaryConflict =
      domain.reasonCode === "NAMED_TEMPLATE_PRIMARY_DOMAIN_CONFLICT";
    reasons.push(resolutionReason(
      domain.reasonCode || "CANDIDATE_DOMAIN_CONFLICT",
      "BLOCKING",
      namedPrimaryConflict
        ? "명명형 template 후보의 핵심 업무영역이 데이터의 primary domain과 일치하지 않습니다."
        : "후보의 업무영역 신호가 데이터의 확정 업무영역과 충돌합니다.",
      {
        expectedDomains: domain.expectedDomains,
        actualDomains: domain.actualDomains,
        primaryDomain: domain.primaryDomain,
        primaryAnchorDomains: domain.primaryAnchorDomains,
        confidence: domain.confidence,
      },
    ));
  } else if (domain.status === "UNKNOWN") {
    reasons.push(resolutionReason(
      "DATASET_DOMAIN_NOT_CONFIDENT",
      "WARNING",
      "업무영역 confidence가 낮아 후보의 domain 적합성을 확정하지 않았습니다.",
      { expectedDomains: domain.expectedDomains, actualDomains: domain.actualDomains },
    ));
  }

  if (executor.status === "UNKNOWN") {
    reasons.push(resolutionReason(
      executor.mode,
      "WARNING",
      "recipe 또는 executor 연결을 결정론적으로 확인하지 못했습니다.",
      { declaredStatus: executor.declaredStatus, recipePresent: executor.recipePresent },
    ));
  } else if (executor.mode === "GENERIC_EXECUTOR_REQUIRES_FEASIBILITY_GATE") {
    reasons.push(resolutionReason(
      "GENERIC_EXECUTOR_REQUIRES_FEASIBILITY_GATE",
      "INFO",
      "generic executor 연결은 확인됐으며 실제 실행 가능성은 후속 Feasibility Gate에서 검증합니다.",
      { outputTypes: executor.outputTypes },
    ));
  }

  if (operandBinding.status === "FAIL") {
    reasons.push(resolutionReason(
      "RECIPE_OPERAND_BINDING_NOT_CONFIRMED",
      bindingStatus === "INFERRED" ? "WARNING" : "BLOCKING",
      "recipe 식별자가 요구하는 그룹·기간·측정값 열을 실제 source table에서 모두 확인하지 못했습니다.",
      {
        operation: operandBinding.operation,
        identifier: operandBinding.identifier,
        missingOperands: operandBinding.operands
          .filter((operand) => operand.status === "FAIL")
          .map((operand) => ({
            kind: operand.kind,
            expectedToken: operand.expectedToken,
          })),
      },
    ));
  } else if (operandBinding.status === "UNKNOWN") {
    reasons.push(resolutionReason(
      "RECIPE_OPERAND_BINDING_AMBIGUOUS",
      "WARNING",
      "recipe operand에 대응하는 열이 여러 개이거나 source scope가 불명확해 정확한 열 결속을 확정하지 못했습니다.",
      {
        operation: operandBinding.operation,
        identifier: operandBinding.identifier,
        ambiguousOperands: operandBinding.operands
          .filter((operand) => operand.status === "UNKNOWN")
          .map((operand) => ({
            kind: operand.kind,
            expectedToken: operand.expectedToken,
            matchedColumnIds: asArray(operand.matched).map((match) => match.columnId),
          })),
      },
    ));
  } else if (operandBinding.status === "PASS") {
    reasons.push(resolutionReason(
      "RECIPE_OPERANDS_BOUND",
      "INFO",
      "recipe 식별자의 그룹·기간·측정값 operand를 실제 source table 열에 결속했습니다.",
      {
        operation: operandBinding.operation,
        matchedColumnIds: operandBinding.matchedColumnIds,
      },
    ));
  }

  const structuralGeneric = isStructuralGenericCandidate(
    retrievalItem,
    capability,
  );
  const inferredIdentityConfirmed =
    bindingStatus !== "INFERRED" ||
    domain.status === "PASS" ||
    structuralGeneric;
  if (!inferredIdentityConfirmed) {
    reasons.push(resolutionReason(
      "INFERRED_TEMPLATE_IDENTITY_NOT_CONFIRMED",
      "WARNING",
      "명명형 INFERRED 후보의 업무 의미를 후보 ID·template ID 또는 데이터 의미 근거로 확인하지 못했습니다.",
      {
        candidateId: normalizeText(retrievalItem.candidateId || ""),
        templateId: normalizeText(retrievalItem.templateId || ""),
        domainStatus: domain.status,
        domainReasonCode: domain.reasonCode,
      },
    ));
  }

  const assessmentBlocking = asArray(assessment?.reasons).filter(
    (item) => item.level === "BLOCKING",
  );
  for (const item of assessmentBlocking) {
    reasons.push(resolutionReason(
      item.code,
      bindingStatus === "INFERRED" ? "WARNING" : "BLOCKING",
      item.message,
      item.details || {},
    ));
  }

  const operandConclusive = ["PASS", "NOT_APPLICABLE"].includes(
    operandBinding.status,
  );
  let result = "STILL_DEFERRED";
  if (domain.status === "FAIL" && domain.confidence >= 0.8) {
    result = "EXCLUDED";
  } else if (scope.status === "FAIL") {
    result = "EXCLUDED";
  } else if (scope.status === "PASS" && executor.status === "PASS") {
    if (
      !assessmentBlocking.length &&
      inferredIdentityConfirmed &&
      operandConclusive
    ) {
      result = "RESOLVED";
    } else if (
      ["BOUND", "PARTIAL"].includes(bindingStatus) &&
      (assessmentBlocking.length || operandBinding.status === "FAIL")
    ) {
      result = "EXCLUDED";
    }
  }

  if (result === "RESOLVED") {
    reasons.push(resolutionReason(
      "SEMANTIC_REQUIREMENTS_RESOLVED",
      "INFO",
      "병합된 의미 profile로 필수 역할·operation·metric·source 조건을 확인했습니다.",
      {},
    ));
  } else if (result === "STILL_DEFERRED" && bindingStatus === "INFERRED") {
    reasons.push(resolutionReason(
      "INFERRED_REQUIREMENTS_NOT_CONCLUSIVE",
      "WARNING",
      "식별자 기반 추론 요구조건만으로는 후보를 안전하게 확정하거나 제외할 수 없습니다.",
      {},
    ));
  }

  return buildResolvedItem({
    retrievalItem,
    capability,
    resolvedProfile,
    result,
    reasons,
    scope,
    assessment,
    domain,
    executor,
    operandBinding,
  });
}

function buildResolvedItem({
  retrievalItem,
  capability,
  resolvedProfile,
  result,
  reasons,
  scope,
  assessment,
  domain,
  executor,
  operandBinding,
}) {
  const sourceCheck = scope
    ? sourceCheckFrom(scope)
    : {
      status: "UNKNOWN",
      mode: "NOT_ASSESSED",
      selectedRootIds: [],
      requestedSourceTableIds: unique(retrievalItem.sourceTableIds),
      matchedTableIds: [],
      matchedPhysicalTableIds: [],
      reasonCode: "NOT_ASSESSED",
      rootCandidates: [],
    };
  const assessed = assessment || {};
  const normalizedOperandBinding = operandBinding || {
    status: "NOT_APPLICABLE",
    operation: "",
    identifier: "",
    operands: [],
    matchedColumnIds: [],
    reasonCode: "NOT_ASSESSED",
  };
  const operandEvidence = asArray(normalizedOperandBinding.operands)
    .flatMap((operand) => asArray(operand.matched));
  const evidence = mergeEvidence(assessed.evidence, operandEvidence);
  const blockingReasons = asArray(reasons).filter((item) => item.level === "BLOCKING");
  const item = {
    version: QUERY_CANDIDATE_RESOLUTION_ITEM_VERSION,
    candidateId: normalizeText(retrievalItem.candidateId || capability.candidateId || ""),
    recipeId: normalizeText(retrievalItem.recipeId || capability.recipeId || ""),
    templateId: normalizeText(retrievalItem.templateId || capability.templateId || ""),
    candidateType: normalizeText(retrievalItem.candidateType || capability.candidateType || "UNKNOWN"),
    result,
    previousRetrievalResult: normalizeText(retrievalItem.result || "DEFERRED"),
    bindingStatus: normalizeText(capability.bindingStatus || retrievalItem.bindingStatus || "UNBOUND"),
    bindingSource: normalizeText(capability.bindingSource || retrievalItem.bindingSource || "NONE"),
    originalRank: Number.isInteger(retrievalItem.originalRank)
      ? retrievalItem.originalRank
      : null,
    originalScore: Number.isFinite(Number(retrievalItem.originalScore))
      ? Number(retrievalItem.originalScore)
      : null,
    resolutionScore: scoreResolution({
      assessment: assessed,
      domain,
      executor,
      scope: scope || {},
    }),
    sourceTableIds: unique(retrievalItem.sourceTableIds),
    matchedTableIds: unique(assessed.matchedTableIds),
    matchedPhysicalTableIds: unique(assessed.matchedPhysicalTableIds),
    matchedColumnIds: unique([
      ...asArray(assessed.matchedColumnIds),
      ...asArray(normalizedOperandBinding.matchedColumnIds),
    ]),
    checks: {
      sourceScope: sourceCheck,
      domainAlignment: domain,
      executorSupport: executor,
      operandBinding: normalizedOperandBinding,
      requiredRoles: asArray(assessed.checks?.requiredRoles),
      requiredCapabilities: asArray(assessed.checks?.requiredCapabilities),
      metricFamily: assessed.checks?.metricFamily || {
        status: "NOT_APPLICABLE",
        required: [],
        available: [],
      },
      constraints: asArray(assessed.checks?.constraints),
    },
    reasons: asArray(reasons),
    missingRequirements: blockingReasons.map((item) => ({
      code: item.code,
      message: item.message,
      details: item.details || {},
    })),
    evidence,
    provenance: {
      retrievalItemVersion: normalizeText(retrievalItem.version || ""),
      retrievalItemSha256: normalizeText(retrievalItem.retrievalItemSha256 || ""),
      capabilityItemVersion: normalizeText(capability.version || ""),
      capabilitySha256: normalizeText(capability.capabilitySha256 || ""),
      resolvedSemanticProfileVersion: normalizeText(resolvedProfile.version || ""),
      resolvedSemanticProfileSha256: normalizeText(resolvedProfile.profileSha256 || ""),
      candidateStatus: normalizeText(
        retrievalItem.provenance?.candidateStatus || "UNASSESSED",
      ),
      terminalPriorResult: false,
      semanticReassessmentPerformed: true,
    },
  };
  item.resolutionItemSha256 = sha256({ ...item, resolutionItemSha256: undefined });
  return item;
}

function resolveCandidate(retrievalItem = {}, capability = {}, profile = {}) {
  if (retrievalItem.result === "RETRIEVED") {
    return carriedItem(retrievalItem, "RESOLVED");
  }
  if (retrievalItem.result === "EXCLUDED") {
    return carriedItem(retrievalItem, "EXCLUDED");
  }
  return resolveDeferredCandidate(retrievalItem, capability, profile);
}

function buildQueryCandidateResolution({
  retrieval = {},
  capabilityManifest = {},
  resolvedSemanticProfile = {},
} = {}) {
  const retrievalCandidates = asArray(retrieval.candidates);
  const capabilityCandidates = asArray(capabilityManifest.candidates);
  const capabilityById = candidateMap(capabilityCandidates);
  const retrievalIds = unique(retrievalCandidates.map((item) => item.candidateId));
  const capabilityIds = unique(capabilityCandidates.map((item) => item.candidateId));
  const retrievalIdSet = new Set(retrievalIds);
  const capabilityIdSet = new Set(capabilityIds);

  const candidates = retrievalCandidates.map((retrievalItem) => {
    const capability = capabilityById.get(normalizeText(retrievalItem.candidateId)) || {
      version: "",
      candidateId: retrievalItem.candidateId,
      recipeId: retrievalItem.recipeId,
      templateId: retrievalItem.templateId,
      candidateType: retrievalItem.candidateType,
      bindingStatus: "UNBOUND",
      bindingSource: "NONE",
      requiredColumnRoles: [],
      requiredCapabilities: [],
      metricFamilies: [],
      executorSupport: { status: "UNKNOWN", outputTypes: [], reasons: [] },
      constraints: {},
    };
    return resolveCandidate(retrievalItem, capability, resolvedSemanticProfile);
  });

  const resolution = {
    version: QUERY_CANDIDATE_RESOLUTION_VERSION,
    itemVersion: QUERY_CANDIDATE_RESOLUTION_ITEM_VERSION,
    policy: {
      version: QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION,
      priorRetrievedCarriedAsResolved: true,
      priorExcludedAreTerminal: true,
      onlyDeferredCandidatesAreSemanticallyReassessed: true,
      unboundCandidatesRemainDeferred: true,
      inferredFailureIsConservative: true,
      genericExecutorMayResolveSemantics: true,
      strongCandidateIdentityOverridesDynamicRecipeTokens: true,
      semanticSecondaryDomainEvidenceAllowed: true,
      inferredNamedTemplatesRequireIdentityEvidence: true,
      inferredAggregateRecipeOperandsRequireColumnBinding: true,
      operandSpecificColumnEvidenceRequired: true,
      finalReadyStatusAssigned: false,
      candidateStatusMutation: false,
    },
    source: {
      caseId: normalizeText(
        retrieval.source?.caseId || resolvedSemanticProfile.source?.caseId || "",
      ),
      fileName: normalizeText(
        retrieval.source?.fileName || resolvedSemanticProfile.source?.fileName || "",
      ),
      retrievalVersion: normalizeText(retrieval.version || ""),
      retrievalSha256: normalizeText(retrieval.retrievalSha256 || ""),
      capabilityManifestVersion: normalizeText(capabilityManifest.version || ""),
      capabilityManifestSha256: normalizeText(capabilityManifest.manifestSha256 || ""),
      resolvedSemanticProfileVersion: normalizeText(resolvedSemanticProfile.version || ""),
      resolvedSemanticProfileSha256: normalizeText(resolvedSemanticProfile.profileSha256 || ""),
      deterministicSemanticProfileSha256: normalizeText(
        resolvedSemanticProfile.source?.deterministicProfileSha256 || "",
      ),
      primaryDomain: normalizeText(
        resolvedSemanticProfile.classification?.primaryDomain || "UNKNOWN",
      ),
      datasetIntent: normalizeText(
        resolvedSemanticProfile.classification?.datasetIntent || "UNKNOWN",
      ),
      domainConfidence: Number(
        resolvedSemanticProfile.classification?.confidence || 0,
      ),
    },
    integrity: {
      retrievalCandidateCount: retrievalCandidates.length,
      capabilityCandidateCount: capabilityCandidates.length,
      missingCapabilityCandidateIds: retrievalIds.filter(
        (id) => !capabilityIdSet.has(id),
      ),
      orphanCapabilityCandidateIds: capabilityIds.filter(
        (id) => !retrievalIdSet.has(id),
      ),
      candidateCountMatch:
        retrievalCandidates.length === capabilityCandidates.length &&
        retrievalIds.every((id) => capabilityIdSet.has(id)),
      retrievalCapabilityHashMatch:
        normalizeText(retrieval.source?.capabilityManifestSha256 || "") ===
        normalizeText(capabilityManifest.manifestSha256 || ""),
      retrievalDeterministicProfileHashMatch:
        normalizeText(retrieval.source?.semanticProfileSha256 || "") ===
        normalizeText(
          resolvedSemanticProfile.source?.deterministicProfileSha256 || "",
        ),
    },
    counts: {
      total: candidates.length,
      resolved: candidates.filter((item) => item.result === "RESOLVED").length,
      stillDeferred: candidates.filter((item) => item.result === "STILL_DEFERRED").length,
      excluded: candidates.filter((item) => item.result === "EXCLUDED").length,
      carriedResolved: candidates.filter(
        (item) => item.result === "RESOLVED" && item.provenance.terminalPriorResult,
      ).length,
      semanticResolved: candidates.filter(
        (item) =>
          item.result === "RESOLVED" &&
          item.previousRetrievalResult === "DEFERRED",
      ).length,
      carriedExcluded: candidates.filter(
        (item) => item.result === "EXCLUDED" && item.provenance.terminalPriorResult,
      ).length,
      newlyExcluded: candidates.filter(
        (item) =>
          item.result === "EXCLUDED" &&
          item.previousRetrievalResult === "DEFERRED",
      ).length,
      inferredResolved: candidates.filter(
        (item) => item.result === "RESOLVED" && item.bindingStatus === "INFERRED",
      ).length,
      unboundStillDeferred: candidates.filter(
        (item) =>
          item.result === "STILL_DEFERRED" && item.bindingStatus === "UNBOUND",
      ).length,
      sourceAmbiguous: candidates.filter(
        (item) => item.checks?.sourceScope?.status === "UNKNOWN",
      ).length,
    },
    candidates,
  };
  resolution.resolutionSha256 = sha256({
    ...resolution,
    resolutionSha256: undefined,
  });
  return resolution;
}

function validationIssue(path, code, message) {
  return { path, code, message };
}

function validateResolutionItem(item = {}, index = 0) {
  const path = `candidates[${index}]`;
  const errors = [];
  const warnings = [];
  if (item.version !== QUERY_CANDIDATE_RESOLUTION_ITEM_VERSION) {
    errors.push(validationIssue(`${path}.version`, "invalid_version", "resolution item version이 유효하지 않습니다."));
  }
  if (!normalizeText(item.candidateId)) {
    errors.push(validationIssue(`${path}.candidateId`, "required", "candidateId가 필요합니다."));
  }
  if (!RESOLUTION_RESULT.includes(item.result)) {
    errors.push(validationIssue(`${path}.result`, "invalid_enum", "resolution result가 유효하지 않습니다."));
  }
  if (!Number.isFinite(Number(item.resolutionScore)) || item.resolutionScore < 0 || item.resolutionScore > 100) {
    errors.push(validationIssue(`${path}.resolutionScore`, "invalid_range", "resolutionScore는 0~100이어야 합니다."));
  }
  for (const check of [
    item.checks?.sourceScope,
    item.checks?.domainAlignment,
    item.checks?.executorSupport,
    item.checks?.operandBinding,
  ]) {
    if (!CHECK_STATUS.includes(check?.status)) {
      errors.push(validationIssue(`${path}.checks`, "invalid_check_status", "check status가 유효하지 않습니다."));
    }
  }
  for (const reason of asArray(item.reasons)) {
    if (!REASON_LEVEL.includes(reason.level)) {
      errors.push(validationIssue(`${path}.reasons`, "invalid_reason_level", "reason level이 유효하지 않습니다."));
    }
  }
  if (!['UNASSESSED', ''].includes(normalizeText(item.provenance?.candidateStatus || ''))) {
    warnings.push(validationIssue(`${path}.provenance.candidateStatus`, "candidate_status_mutated", "Resolver는 candidate status를 변경하지 않아야 합니다."));
  }
  if (item.result === "STILL_DEFERRED") {
    warnings.push(validationIssue(path, "candidate_still_deferred", "후속 manifest 보강 또는 조건부 Planner 검토가 필요합니다."));
  }
  const expectedSha = sha256({ ...item, resolutionItemSha256: undefined });
  if (item.resolutionItemSha256 !== expectedSha) {
    errors.push(validationIssue(`${path}.resolutionItemSha256`, "sha_mismatch", "resolution item SHA-256이 일치하지 않습니다."));
  }
  return { errors, warnings };
}

function validateQueryCandidateResolution(resolution = {}) {
  const errors = [];
  const warnings = [];
  if (resolution.version !== QUERY_CANDIDATE_RESOLUTION_VERSION) {
    errors.push(validationIssue("version", "invalid_version", "resolution version이 유효하지 않습니다."));
  }
  if (resolution.itemVersion !== QUERY_CANDIDATE_RESOLUTION_ITEM_VERSION) {
    errors.push(validationIssue("itemVersion", "invalid_version", "resolution item version이 유효하지 않습니다."));
  }
  if (resolution.policy?.version !== QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION) {
    errors.push(validationIssue("policy.version", "invalid_version", "resolution policy version이 유효하지 않습니다."));
  }
  const candidates = asArray(resolution.candidates);
  if (!Array.isArray(resolution.candidates)) {
    errors.push(validationIssue("candidates", "invalid_type", "candidates는 배열이어야 합니다."));
  }
  const ids = new Set();
  candidates.forEach((item, index) => {
    const validation = validateResolutionItem(item, index);
    errors.push(...validation.errors);
    warnings.push(...validation.warnings);
    if (ids.has(item.candidateId)) {
      errors.push(validationIssue(`candidates[${index}].candidateId`, "duplicate", "candidateId가 중복됩니다."));
    }
    ids.add(item.candidateId);
  });

  const expectedCounts = {
    total: candidates.length,
    resolved: candidates.filter((item) => item.result === "RESOLVED").length,
    stillDeferred: candidates.filter((item) => item.result === "STILL_DEFERRED").length,
    excluded: candidates.filter((item) => item.result === "EXCLUDED").length,
    carriedResolved: candidates.filter(
      (item) => item.result === "RESOLVED" && item.provenance?.terminalPriorResult,
    ).length,
    semanticResolved: candidates.filter(
      (item) => item.result === "RESOLVED" && item.previousRetrievalResult === "DEFERRED",
    ).length,
    carriedExcluded: candidates.filter(
      (item) => item.result === "EXCLUDED" && item.provenance?.terminalPriorResult,
    ).length,
    newlyExcluded: candidates.filter(
      (item) => item.result === "EXCLUDED" && item.previousRetrievalResult === "DEFERRED",
    ).length,
    inferredResolved: candidates.filter(
      (item) => item.result === "RESOLVED" && item.bindingStatus === "INFERRED",
    ).length,
    unboundStillDeferred: candidates.filter(
      (item) => item.result === "STILL_DEFERRED" && item.bindingStatus === "UNBOUND",
    ).length,
    sourceAmbiguous: candidates.filter(
      (item) => item.checks?.sourceScope?.status === "UNKNOWN",
    ).length,
  };
  for (const [key, expected] of Object.entries(expectedCounts)) {
    if (Number(resolution.counts?.[key] || 0) !== expected) {
      errors.push(validationIssue(`counts.${key}`, "count_mismatch", `${key} count가 실제 후보 수와 다릅니다.`));
    }
  }

  if (!resolution.integrity?.retrievalCapabilityHashMatch) {
    errors.push(validationIssue("integrity.retrievalCapabilityHashMatch", "source_hash_mismatch", "retrieval과 capability manifest hash가 일치하지 않습니다."));
  }
  if (!resolution.integrity?.retrievalDeterministicProfileHashMatch) {
    errors.push(validationIssue("integrity.retrievalDeterministicProfileHashMatch", "source_hash_mismatch", "retrieval의 deterministic profile과 resolved profile source hash가 일치하지 않습니다."));
  }
  if (asArray(resolution.integrity?.missingCapabilityCandidateIds).length) {
    warnings.push(validationIssue("integrity.missingCapabilityCandidateIds", "capability_candidates_missing", "일부 후보 capability가 없어 UNBOUND로 처리됐습니다."));
  }
  if (asArray(resolution.integrity?.orphanCapabilityCandidateIds).length) {
    warnings.push(validationIssue("integrity.orphanCapabilityCandidateIds", "capability_candidates_orphaned", "retrieval에 없는 capability 후보가 있습니다."));
  }

  const expectedSha = sha256({ ...resolution, resolutionSha256: undefined });
  if (resolution.resolutionSha256 !== expectedSha) {
    errors.push(validationIssue("resolutionSha256", "sha_mismatch", "resolution SHA-256이 일치하지 않습니다."));
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
  QUERY_CANDIDATE_RESOLUTION_VERSION,
  QUERY_CANDIDATE_RESOLUTION_ITEM_VERSION,
  QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION,
  PREVIOUS_QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION,
  RESOLUTION_RESULT,
  buildQueryCandidateResolution,
  validateQueryCandidateResolution,
  resolveCandidate,
  resolveDeferredCandidate,
  enhancedSourceScope,
  domainAlignmentCheck,
  semanticEvidenceDomains,
  isStructuralGenericCandidate,
  parsedRecipeOperandSpec,
  recipeOperandBindingCheck,
  executorSupportCheck,
};
