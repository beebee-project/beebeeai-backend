const fs = require("fs");
const path = require("path");
const { normalizeText, sha256 } = require("./queryCandidateObservation");

const QUERY_CANDIDATE_CAPABILITY_MANIFEST_VERSION =
  "query_candidate_capability_manifest_v1";
const QUERY_CANDIDATE_CAPABILITY_ITEM_VERSION =
  "query_candidate_capability_item_v1";
const QUERY_CANDIDATE_CAPABILITY_OVERLAY_VERSION =
  "query_candidate_capability_overlay_v1";

const BINDING_STATUS = Object.freeze([
  "BOUND",
  "PARTIAL",
  "INFERRED",
  "UNBOUND",
]);
const BINDING_SOURCE = Object.freeze([
  "CONTRACT_CATALOG",
  "OVERLAY",
  "EXPLICIT_CANDIDATE",
  "IDENTIFIER_INFERENCE",
  "NONE",
]);
const EXECUTOR_SUPPORT_STATUS = Object.freeze([
  "DECLARED",
  "GENERIC",
  "UNKNOWN",
]);

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const rows = [];
  for (const value of asArray(values)) {
    const text = normalizeText(value);
    if (!text || seen.has(text)) continue;
    seen.add(text);
    rows.push(text);
  }
  return rows;
}

function normalizeKey(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/^template[_:\-.]*/, "")
    .replace(/[\s\\/.:\-]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeLoose(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[^가-힣a-z0-9]+/g, "");
}

function readJsonIfExists(filePath, fallback = null) {
  try {
    if (!filePath || !fs.existsSync(filePath)) return fallback;
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch (_error) {
    return fallback;
  }
}

function loadDefaultCapabilitySources({ rootDir = process.cwd() } = {}) {
  const catalogPath = path.resolve(
    rootDir,
    "automation",
    "summarySheetContractCatalog.json",
  );
  const overlayPath = path.resolve(
    rootDir,
    "automation",
    "queryCandidateCapabilityOverlay.json",
  );
  return {
    catalogPath,
    overlayPath,
    contractCatalog: readJsonIfExists(catalogPath, {
      version: "",
      contracts: {},
    }),
    overlay: readJsonIfExists(overlayPath, {
      version: QUERY_CANDIDATE_CAPABILITY_OVERLAY_VERSION,
      entries: {},
    }),
  };
}

function sourceRole(item = {}, requiredFallback = false, source = "") {
  const aliases = unique([...asArray(item.aliases), item.label, item.role]);
  return {
    role: normalizeText(item.role || item.label || ""),
    aliases,
    dataType: normalizeText(item.dataType || item.type || ""),
    semanticType: normalizeText(item.semanticType || ""),
    required:
      item.required === true ||
      (item.required == null && requiredFallback === true),
    source: normalizeText(source),
  };
}

function normalizeMetric(metric = {}, source = "") {
  const activation = metric.activation || {};
  return {
    metricId: normalizeText(metric.metricId || metric.id || ""),
    label: normalizeText(metric.label || metric.title || metric.metricId || ""),
    criticality: normalizeText(metric.criticality || "conditional"),
    kind: normalizeText(metric.kind || ""),
    aggregation: normalizeText(metric.aggregation || ""),
    operator: normalizeText(metric.operator || ""),
    valueRole: normalizeText(metric.valueRole || ""),
    groupByRoles: unique(metric.groupByRoles),
    sourceMetricId: normalizeText(metric.sourceMetricId || ""),
    requiresRoles: unique([
      ...asArray(activation.requiresRoles),
      metric.valueRole,
      ...asArray(metric.groupByRoles),
    ]),
    requiresMetricIds: unique(activation.requiresMetricIds),
    source: normalizeText(source),
  };
}

function capabilityTokensForMetric(metric = {}) {
  const result = [];
  const operation = metric.aggregation || metric.operator || metric.kind;
  if (operation) result.push(`operation:${operation}`);
  if (metric.kind) result.push(`metric_kind:${metric.kind}`);
  if (metric.groupByRoles?.length) result.push("group_by");
  if (metric.kind === "rank" || metric.sourceMetricId) result.push("ranking");
  if (metric.requiresMetricIds?.length) result.push("metric_dependency");
  return result;
}

function metricFamilies(metrics = []) {
  return unique(
    metrics
      .map((metric) => normalizeText(metric.metricId).split(".")[0])
      .filter(Boolean),
  );
}

function operations(metrics = []) {
  return unique(
    metrics.flatMap((metric) => [
      metric.aggregation,
      metric.operator,
      metric.kind === "rank" ? "rank" : "",
    ]),
  );
}

function normalizeExecutorSupport(value = {}, candidate = {}) {
  const status = normalizeText(value.status || "").toUpperCase();
  const normalizedStatus = EXECUTOR_SUPPORT_STATUS.includes(status)
    ? status
    : candidate.outputTypes?.length
      ? "GENERIC"
      : "UNKNOWN";
  return {
    status: normalizedStatus,
    outputTypes: unique([
      ...asArray(value.outputTypes),
      ...asArray(candidate.outputTypes),
      candidate.outputType,
    ]),
    reasons: unique(value.reasons),
  };
}

function catalogEntries(contractCatalog = {}) {
  const contracts =
    contractCatalog && typeof contractCatalog.contracts === "object"
      ? contractCatalog.contracts
      : {};
  return Object.entries(contracts).map(([key, contract]) => ({
    key,
    contract: contract || {},
    matchKeys: unique([
      normalizeKey(key),
      normalizeKey(contract?.templateId),
      normalizeKey(contract?.contractId),
    ]),
  }));
}

function overlayEntries(overlay = {}) {
  const entries =
    overlay && typeof overlay.entries === "object" ? overlay.entries : {};
  return Object.entries(entries).map(([key, entry]) => ({
    key,
    entry: entry || {},
    matchKeys: unique([
      normalizeKey(key),
      ...asArray(entry?.aliases).map(normalizeKey),
    ]),
  }));
}

function candidateMatchKeys(candidate = {}) {
  return unique([
    candidate.candidateId,
    candidate.templateId,
    candidate.recipeId,
    ...asArray(candidate.recipeIds),
  ]).map(normalizeKey);
}

function exactMatches(entries = [], keys = []) {
  const keySet = new Set(keys.filter(Boolean));
  return entries.filter((entry) =>
    entry.matchKeys.some((key) => keySet.has(key)),
  );
}

function looseMatches(entries = [], candidate = {}) {
  const candidateValues = unique([
    candidate.candidateId,
    candidate.templateId,
    candidate.recipeId,
    ...asArray(candidate.recipeIds),
  ])
    .map(normalizeLoose)
    .filter((value) => value.length >= 6);
  if (!candidateValues.length) return [];
  return entries.filter((entry) =>
    entry.matchKeys
      .map(normalizeLoose)
      .some((key) =>
        candidateValues.some(
          (value) =>
            key &&
            (key === value || key.includes(value) || value.includes(key)),
        ),
      ),
  );
}

function fromCatalogMatch(match = {}) {
  const contract = match.contract || {};
  const roles = asArray(contract.sourceRoles).map((role) =>
    sourceRole(role, role?.required === true, "contract_catalog"),
  );
  const metrics = asArray(contract.metrics)
    .map((metric) => normalizeMetric(metric, "contract_catalog"))
    .filter((metric) => metric.metricId);
  return {
    bindingSource: "CONTRACT_CATALOG",
    bindingKey: normalizeText(match.key || ""),
    contractIds: unique([contract.contractId]),
    matchedTemplateIds: unique([contract.templateId, match.key]),
    requiredColumnRoles: roles.filter((role) => role.required),
    optionalColumnRoles: roles.filter((role) => !role.required),
    metrics,
    executorSupport: {
      status: "DECLARED",
      outputTypes: ["summarySheet"],
      reasons: ["summary_sheet_contract_catalog"],
    },
  };
}

function fromOverlayMatch(match = {}) {
  const entry = match.entry || {};
  const requiredColumnRoles = asArray(entry.requiredColumnRoles).map((role) =>
    sourceRole(role, true, "capability_overlay"),
  );
  const optionalColumnRoles = asArray(entry.optionalColumnRoles).map((role) =>
    sourceRole(role, false, "capability_overlay"),
  );
  const metrics = asArray(entry.metrics)
    .map((metric) => normalizeMetric(metric, "capability_overlay"))
    .filter((metric) => metric.metricId);
  return {
    bindingSource: "OVERLAY",
    bindingKey: normalizeText(match.key || ""),
    contractIds: unique(entry.contractIds),
    matchedTemplateIds: unique([entry.templateId, ...asArray(entry.aliases)]),
    requiredColumnRoles,
    optionalColumnRoles,
    metrics,
    executorSupport: normalizeExecutorSupport(entry.executorSupport || {}, {}),
  };
}

function explicitCandidateDescriptor(candidate = {}) {
  const roles = asArray(candidate.requiredColumnRoles || candidate.columnRoles);
  const requiredColumnRoles = roles.map((role) =>
    typeof role === "string"
      ? sourceRole({ role, required: true }, true, "explicit_candidate")
      : sourceRole(role, role?.required !== false, "explicit_candidate"),
  );
  const metrics = asArray(candidate.metricContracts || candidate.metrics)
    .map((metric) => normalizeMetric(metric, "explicit_candidate"))
    .filter((metric) => metric.metricId);
  const explicitCapabilities = unique(candidate.requiredCapabilities);
  if (
    !requiredColumnRoles.length &&
    !metrics.length &&
    !explicitCapabilities.length
  ) {
    return null;
  }
  return {
    bindingSource: "EXPLICIT_CANDIDATE",
    bindingKey: candidate.candidateId,
    contractIds: [],
    matchedTemplateIds: unique([candidate.templateId]),
    requiredColumnRoles,
    optionalColumnRoles: [],
    metrics,
    explicitCapabilities,
    executorSupport: normalizeExecutorSupport({}, candidate),
  };
}

function inferDescriptor(candidate = {}) {
  const rawText = normalizeText(
    [
      candidate.candidateId,
      candidate.templateId,
      candidate.recipeId,
      ...asArray(candidate.recipeIds),
      candidate.title,
    ].join(" "),
  ).toLowerCase();
  const compactText = normalizeLoose(rawText);
  const requiredRoles = [];
  const metrics = [];

  const addRole = (role, dataType, semanticType, aliases = []) => {
    if (requiredRoles.some((item) => item.role === role)) return;
    requiredRoles.push(
      sourceRole(
        { role, aliases, dataType, semanticType, required: true },
        true,
        "identifier_inference",
      ),
    );
  };

  const has = (...tokens) =>
    tokens.some((token) => {
      const normalizedToken = normalizeLoose(token);
      if (!normalizedToken) return false;
      if (/^[a-z0-9]+$/u.test(normalizedToken)) {
        const escaped = normalizedToken.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        return new RegExp(`(^|[^a-z0-9])${escaped}(?=$|[^a-z0-9])`, "u").test(
          rawText,
        );
      }
      return compactText.includes(normalizedToken);
    });
  if (has("월별", "기간별", "monthly", "period", "date", "trend", "추이")) {
    addRole("period", "period", "dimension", ["기간", "일자", "연월", "date"]);
  }
  if (has("매출", "sales", "revenue", "금액", "amount")) {
    addRole("amount", "number", "measure", ["금액", "매출액", "amount"]);
  }
  if (has("품목", "상품", "product", "item", "category", "카테고리")) {
    addRole("category", "string", "dimension", ["품목", "상품", "분류"]);
  }
  if (has("출석", "참석", "attendance", "status", "상태")) {
    addRole("status", "string", "dimension", ["출석상태", "참석상태", "상태"]);
  }
  if (has("명단", "roster", "person", "이름", "성명")) {
    addRole("person", "string", "dimension", ["이름", "성명", "대상자"]);
  }
  if (has("부서", "조직", "소속", "department", "organization", "group")) {
    addRole("group", "string", "dimension", ["부서", "조직", "소속"]);
  }

  let operation = "";
  let kind = "aggregate";
  if (has("평균", "average", "avg")) operation = "average";
  else if (has("순위", "ranking", "rank", "top", "최고", "최다")) {
    operation = "rank";
    kind = "rank";
  } else if (has("비율", "rate", "ratio", "percentage", "퍼센트")) {
    operation = "percentage";
    kind = "derived";
  } else if (has("건수", "개수", "인원", "count", "명단", "출석", "상태")) {
    operation = "countRows";
  } else if (has("합계", "총", "매출", "금액", "sum", "sales", "amount")) {
    operation = "sum";
  }

  if (!requiredRoles.length && !operation) return null;
  metrics.push(
    normalizeMetric(
      {
        metricId: `inferred.${normalizeKey(candidate.candidateId || "candidate")}`,
        label: candidate.title || candidate.candidateId,
        criticality: "conditional",
        kind,
        aggregation: ["sum", "average", "countRows"].includes(operation)
          ? operation
          : "",
        operator: ["percentage"].includes(operation) ? operation : "",
        groupByRoles: requiredRoles
          .filter((role) => role.semanticType === "dimension")
          .slice(0, 1)
          .map((role) => role.role),
        activation: {
          requiresRoles: requiredRoles.map((role) => role.role),
          requiresMetricIds: [],
        },
      },
      "identifier_inference",
    ),
  );
  return {
    bindingSource: "IDENTIFIER_INFERENCE",
    bindingKey: candidate.candidateId,
    contractIds: [],
    matchedTemplateIds: unique([candidate.templateId]),
    requiredColumnRoles: requiredRoles,
    optionalColumnRoles: [],
    metrics,
    executorSupport: normalizeExecutorSupport(
      { status: "GENERIC", reasons: ["identifier_inference"] },
      candidate,
    ),
  };
}

function buildDescriptor({ candidate, catalog, overlay } = {}) {
  const keys = candidateMatchKeys(candidate);
  const catalogList = catalogEntries(catalog);
  const overlayList = overlayEntries(overlay);

  const catalogExact = exactMatches(catalogList, keys);
  if (catalogExact.length) {
    return { status: "BOUND", ...fromCatalogMatch(catalogExact[0]) };
  }
  const overlayExact = exactMatches(overlayList, keys);
  if (overlayExact.length) {
    return { status: "BOUND", ...fromOverlayMatch(overlayExact[0]) };
  }

  const explicit = explicitCandidateDescriptor(candidate);
  if (explicit) return { status: "PARTIAL", ...explicit };

  const catalogLoose = looseMatches(catalogList, candidate);
  if (catalogLoose.length === 1) {
    return { status: "PARTIAL", ...fromCatalogMatch(catalogLoose[0]) };
  }
  const overlayLoose = looseMatches(overlayList, candidate);
  if (overlayLoose.length === 1) {
    return { status: "PARTIAL", ...fromOverlayMatch(overlayLoose[0]) };
  }

  const inferred = inferDescriptor(candidate);
  if (inferred) return { status: "INFERRED", ...inferred };

  return {
    status: "UNBOUND",
    bindingSource: "NONE",
    bindingKey: "",
    contractIds: [],
    matchedTemplateIds: [],
    requiredColumnRoles: [],
    optionalColumnRoles: [],
    metrics: [],
    executorSupport: normalizeExecutorSupport({}, candidate),
  };
}

function capabilityItem(candidate = {}, descriptor = {}) {
  const metrics = asArray(descriptor.metrics);
  const explicitCapabilities = unique(descriptor.explicitCapabilities);
  const requiredCapabilities = unique([
    ...explicitCapabilities,
    ...metrics.flatMap(capabilityTokensForMetric),
    ...asArray(descriptor.requiredColumnRoles).map(
      (role) => `column_role:${role.role}`,
    ),
    ...(candidate.sourceTableIds?.length > 1
      ? ["multi_table"]
      : ["single_table"]),
  ]);
  const item = {
    version: QUERY_CANDIDATE_CAPABILITY_ITEM_VERSION,
    candidateId: normalizeText(candidate.candidateId || ""),
    recipeId: normalizeText(candidate.recipeId || ""),
    recipeIds: unique(candidate.recipeIds),
    templateId: normalizeText(candidate.templateId || ""),
    candidateType: normalizeText(candidate.candidateType || "UNKNOWN"),
    bindingStatus: descriptor.status,
    bindingSource: descriptor.bindingSource,
    bindingKey: normalizeText(descriptor.bindingKey || ""),
    contractIds: unique(descriptor.contractIds),
    matchedTemplateIds: unique(descriptor.matchedTemplateIds),
    requiredColumnRoles: asArray(descriptor.requiredColumnRoles),
    optionalColumnRoles: asArray(descriptor.optionalColumnRoles),
    metricContracts: metrics,
    coreMetricIds: unique(
      metrics
        .filter((metric) => metric.criticality === "core")
        .map((metric) => metric.metricId),
    ),
    conditionalMetricIds: unique(
      metrics
        .filter((metric) => metric.criticality !== "core")
        .map((metric) => metric.metricId),
    ),
    supportedMetricIds: unique(metrics.map((metric) => metric.metricId)),
    metricFamilies: metricFamilies(metrics),
    supportedOperations: operations(metrics),
    requiredCapabilities,
    executorSupport: normalizeExecutorSupport(
      descriptor.executorSupport || {},
      candidate,
    ),
    constraints: {
      minimumTableCount: candidate.sourceTableIds?.length ? 1 : 0,
      maximumTableCount:
        candidate.sourceTableIds?.length > 1
          ? candidate.sourceTableIds.length
          : 1,
      sourceScope:
        candidate.sourceTableIds?.length > 1 ? "multiTable" : "singleTable",
      minimumRowCount: 1,
    },
    provenance: {
      candidateContractVersion: normalizeText(candidate.version || ""),
      candidateStatus: normalizeText(candidate.status || ""),
      observedClass: normalizeText(candidate.observedClass || ""),
    },
  };
  item.capabilitySha256 = sha256({ ...item, capabilitySha256: undefined });
  return item;
}

function buildQueryCandidateCapabilityManifest({
  contract = {},
  contractCatalog = {},
  overlay = {},
} = {}) {
  const candidates = asArray(contract.candidates).map((candidate) =>
    capabilityItem(
      candidate,
      buildDescriptor({ candidate, catalog: contractCatalog, overlay }),
    ),
  );
  const manifest = {
    version: QUERY_CANDIDATE_CAPABILITY_MANIFEST_VERSION,
    itemVersion: QUERY_CANDIDATE_CAPABILITY_ITEM_VERSION,
    source: {
      caseId: normalizeText(contract.source?.caseId || ""),
      fileName: normalizeText(contract.source?.fileName || ""),
      contractVersion: normalizeText(contract.version || ""),
      contractSha256: normalizeText(contract.contractSha256 || ""),
    },
    sources: {
      contractCatalogVersion: normalizeText(contractCatalog.version || ""),
      contractCatalogCount: Object.keys(contractCatalog.contracts || {}).length,
      overlayVersion: normalizeText(overlay.version || ""),
      overlayEntryCount: Object.keys(overlay.entries || {}).length,
    },
    counts: {
      total: candidates.length,
      bound: candidates.filter((item) => item.bindingStatus === "BOUND").length,
      partial: candidates.filter((item) => item.bindingStatus === "PARTIAL")
        .length,
      inferred: candidates.filter((item) => item.bindingStatus === "INFERRED")
        .length,
      unbound: candidates.filter((item) => item.bindingStatus === "UNBOUND")
        .length,
      executorDeclared: candidates.filter(
        (item) => item.executorSupport.status === "DECLARED",
      ).length,
      executorGeneric: candidates.filter(
        (item) => item.executorSupport.status === "GENERIC",
      ).length,
      executorUnknown: candidates.filter(
        (item) => item.executorSupport.status === "UNKNOWN",
      ).length,
    },
    candidates,
  };
  manifest.manifestSha256 = sha256({ ...manifest, manifestSha256: undefined });
  return manifest;
}

function issue(pathValue, code, message) {
  return { path: pathValue, code, message };
}

function validateRole(role = {}, pathValue = "") {
  const errors = [];
  if (!normalizeText(role.role)) {
    errors.push(issue(`${pathValue}.role`, "required", "role이 필요합니다."));
  }
  if (!Array.isArray(role.aliases)) {
    errors.push(
      issue(
        `${pathValue}.aliases`,
        "invalid_type",
        "aliases는 배열이어야 합니다.",
      ),
    );
  }
  if (typeof role.required !== "boolean") {
    errors.push(
      issue(
        `${pathValue}.required`,
        "invalid_type",
        "required는 boolean이어야 합니다.",
      ),
    );
  }
  return errors;
}

function validateCapabilityItem(item = {}, index = 0) {
  const pathValue = `candidates[${index}]`;
  const errors = [];
  const warnings = [];
  if (item.version !== QUERY_CANDIDATE_CAPABILITY_ITEM_VERSION) {
    errors.push(
      issue(
        `${pathValue}.version`,
        "invalid_version",
        "item version이 유효하지 않습니다.",
      ),
    );
  }
  if (!normalizeText(item.candidateId)) {
    errors.push(
      issue(
        `${pathValue}.candidateId`,
        "required",
        "candidateId가 필요합니다.",
      ),
    );
  }
  if (!BINDING_STATUS.includes(item.bindingStatus)) {
    errors.push(
      issue(
        `${pathValue}.bindingStatus`,
        "invalid_enum",
        "bindingStatus가 유효하지 않습니다.",
      ),
    );
  }
  if (!BINDING_SOURCE.includes(item.bindingSource)) {
    errors.push(
      issue(
        `${pathValue}.bindingSource`,
        "invalid_enum",
        "bindingSource가 유효하지 않습니다.",
      ),
    );
  }
  if (!EXECUTOR_SUPPORT_STATUS.includes(item.executorSupport?.status)) {
    errors.push(
      issue(
        `${pathValue}.executorSupport.status`,
        "invalid_enum",
        "executor support status가 유효하지 않습니다.",
      ),
    );
  }
  asArray(item.requiredColumnRoles).forEach((role, roleIndex) => {
    errors.push(
      ...validateRole(role, `${pathValue}.requiredColumnRoles[${roleIndex}]`),
    );
  });
  asArray(item.optionalColumnRoles).forEach((role, roleIndex) => {
    errors.push(
      ...validateRole(role, `${pathValue}.optionalColumnRoles[${roleIndex}]`),
    );
  });
  if (item.bindingStatus === "BOUND" && !item.bindingKey) {
    errors.push(
      issue(
        `${pathValue}.bindingKey`,
        "required_for_bound",
        "BOUND 후보에는 bindingKey가 필요합니다.",
      ),
    );
  }
  if (item.bindingStatus === "UNBOUND") {
    warnings.push(
      issue(
        pathValue,
        "candidate_capability_unbound",
        "후보 capability가 아직 manifest에 연결되지 않았습니다.",
      ),
    );
  }
  if (item.bindingStatus === "INFERRED") {
    warnings.push(
      issue(
        pathValue,
        "candidate_capability_inferred",
        "식별자 기반 추론 capability입니다.",
      ),
    );
  }
  if (item.executorSupport?.status !== "DECLARED") {
    warnings.push(
      issue(
        `${pathValue}.executorSupport`,
        "executor_support_not_declared",
        "executor 지원이 명시적으로 선언되지 않았습니다.",
      ),
    );
  }
  const expectedSha = sha256({ ...item, capabilitySha256: undefined });
  if (item.capabilitySha256 !== expectedSha) {
    errors.push(
      issue(
        `${pathValue}.capabilitySha256`,
        "sha_mismatch",
        "capability SHA-256이 일치하지 않습니다.",
      ),
    );
  }
  return { errors, warnings };
}

function validateQueryCandidateCapabilityManifest(manifest = {}) {
  const errors = [];
  const warnings = [];
  if (manifest.version !== QUERY_CANDIDATE_CAPABILITY_MANIFEST_VERSION) {
    errors.push(
      issue(
        "version",
        "invalid_version",
        "manifest version이 유효하지 않습니다.",
      ),
    );
  }
  if (manifest.itemVersion !== QUERY_CANDIDATE_CAPABILITY_ITEM_VERSION) {
    errors.push(
      issue(
        "itemVersion",
        "invalid_version",
        "itemVersion이 유효하지 않습니다.",
      ),
    );
  }
  if (!Array.isArray(manifest.candidates)) {
    errors.push(
      issue("candidates", "invalid_type", "candidates는 배열이어야 합니다."),
    );
  } else {
    manifest.candidates.forEach((item, index) => {
      const validation = validateCapabilityItem(item, index);
      errors.push(...validation.errors);
      warnings.push(...validation.warnings);
    });
  }
  if (manifest.counts?.total !== asArray(manifest.candidates).length) {
    errors.push(
      issue(
        "counts.total",
        "count_mismatch",
        "candidate 수와 counts.total이 다릅니다.",
      ),
    );
  }
  const statusTotal = ["bound", "partial", "inferred", "unbound"].reduce(
    (sum, key) => sum + Number(manifest.counts?.[key] || 0),
    0,
  );
  if (statusTotal !== Number(manifest.counts?.total || 0)) {
    errors.push(
      issue(
        "counts",
        "binding_count_mismatch",
        "binding 상태 합계가 total과 다릅니다.",
      ),
    );
  }
  const expectedSha = sha256({ ...manifest, manifestSha256: undefined });
  if (manifest.manifestSha256 !== expectedSha) {
    errors.push(
      issue(
        "manifestSha256",
        "sha_mismatch",
        "manifest SHA-256이 일치하지 않습니다.",
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
  QUERY_CANDIDATE_CAPABILITY_MANIFEST_VERSION,
  QUERY_CANDIDATE_CAPABILITY_ITEM_VERSION,
  QUERY_CANDIDATE_CAPABILITY_OVERLAY_VERSION,
  BINDING_STATUS,
  BINDING_SOURCE,
  EXECUTOR_SUPPORT_STATUS,
  loadDefaultCapabilitySources,
  buildQueryCandidateCapabilityManifest,
  validateQueryCandidateCapabilityManifest,
  buildDescriptor,
  inferDescriptor,
};
