"use strict";

const catalog = require("./summarySheetContractCatalog.json");

const CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION =
  "contract_driven_summary_recipe_v1";

const ROLE_ALIAS_OVERRIDES = Object.freeze({
  inventory_stock_status: Object.freeze({
    stockQuantity: ["현재재고", "기말재고", "보유수량"],
  }),
  corporate_card_usage_report: Object.freeze({
    category: ["사용유형", "사용구분", "지출유형"],
  }),
});

function normalizeHeader(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s_\-./\\()[\]{}]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function compactText(value = "") {
  return String(value == null ? "" : value).normalize("NFKC").trim();
}

function getRows(table = {}) {
  return Array.isArray(table.rows) ? table.rows : [];
}

function getColumns(table = {}) {
  if (Array.isArray(table.columns) && table.columns.length) return table.columns;
  const first = getRows(table)[0];
  if (first && !Array.isArray(first) && typeof first === "object") {
    return Object.keys(first).map((header) => ({ header }));
  }
  return [];
}

function columnHeader(column = {}) {
  return (
    column.header ||
    column.originalHeader ||
    column.name ||
    column.key ||
    column.accessor ||
    ""
  );
}

function tableHeaders(table = {}) {
  return getColumns(table).map(columnHeader).filter(Boolean);
}

function getRowValue(row = {}, header = "", table = {}) {
  if (!header) return undefined;
  if (Array.isArray(row)) {
    const index = tableHeaders(table).findIndex(
      (candidate) => normalizeHeader(candidate) === normalizeHeader(header),
    );
    return index >= 0 ? row[index] : undefined;
  }
  if (Object.prototype.hasOwnProperty.call(row || {}, header)) return row[header];
  const target = normalizeHeader(header);
  const key = Object.keys(row || {}).find(
    (candidate) => normalizeHeader(candidate) === target,
  );
  return key ? row[key] : undefined;
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null || value === "") return null;
  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/[^\d.+-]/g, "");
  if (!cleaned || cleaned === "." || cleaned === "+" || cleaned === "-") {
    return null;
  }
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function toBoolean(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const text = normalizeHeader(value);
  if (!text) return false;
  if (["true", "yes", "y", "1", "예", "해당", "입사", "퇴사"].includes(text)) {
    return true;
  }
  if (["false", "no", "n", "0", "아니오", "미해당"].includes(text)) {
    return false;
  }
  return true;
}

function normalizePeriod(value) {
  if (value == null || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}`;
  }
  const raw = compactText(value).replace(/\s+/g, "");
  const ym = raw.match(/^((?:19|20)\d{2})[.\-/년]?(1[0-2]|0?[1-9])(?:월)?$/);
  if (ym) return `${ym[1]}-${String(Number(ym[2])).padStart(2, "0")}`;
  const yearOnly = raw.match(/^((?:19|20)\d{2})년?$/);
  if (yearOnly) return yearOnly[1];
  return raw;
}

function headerMatchScore(header = "", aliases = []) {
  const normalized = normalizeHeader(header);
  if (!normalized) return 0;
  let best = 0;
  for (const alias of aliases || []) {
    const target = normalizeHeader(alias);
    if (!target) continue;
    if (normalized === target) best = Math.max(best, 1000 + target.length);
    else if (normalized.includes(target)) best = Math.max(best, 700 + target.length);
    else if (target.includes(normalized) && normalized.length >= 2) {
      best = Math.max(best, 450 + normalized.length);
    }
  }
  return best;
}

function findHeader(table = {}, aliases = []) {
  let best = null;
  for (const header of tableHeaders(table)) {
    const score = headerMatchScore(header, aliases);
    if (score && (!best || score > best.score)) best = { header, score };
  }
  return best?.header || "";
}

function allHeaders(tables = []) {
  return tables.flatMap(tableHeaders);
}

function hasAlias(headers = [], aliases = []) {
  return headers.some((header) => headerMatchScore(header, aliases) >= 700);
}

function chooseContract(templateId = "", tables = []) {
  const base = catalog.contracts?.[templateId] || null;
  if (!base) return null;
  const headers = allHeaders(tables);

  if (
    templateId === "research_budget_report" &&
    hasAlias(headers, ["정부출연금"]) &&
    hasAlias(headers, ["총 연구비", "총연구비"])
  ) {
    return catalog.variants?.research_budget_funding_v1 || base;
  }
  if (
    templateId === "purchase_order_status" &&
    hasAlias(headers, ["요청기관"])
  ) {
    return catalog.variants?.purchase_requesting_organization_v1 || base;
  }
  if (templateId === "sales_report") {
    const hasYear = hasAlias(headers, ["연도", "년도", "year"]);
    const hasMonth = hasAlias(headers, ["월", "month"]);
    const hasProduct = hasAlias(headers, ["품목", "제품", "상품", "product"]);
    if (hasYear && hasMonth && !hasProduct) {
      return catalog.variants?.sales_monthly_without_product_v1 || base;
    }
  }
  if (templateId === "hr_monthly_report") {
    const hasStatus = hasAlias(headers, ["재직상태", "상태", "status"]);
    if (!hasStatus) return catalog.variants?.hr_without_status_v1 || base;
  }
  return base;
}

function scoreColumnsForSurvey(table = {}) {
  return tableHeaders(table).filter((header) => {
    const text = normalizeHeader(header);
    if (!/(만족도|점수|평점|score|rating)/i.test(text)) return false;
    if (/(재참여|추천|의향|nps)/i.test(text)) return false;
    const values = getRows(table)
      .slice(0, 20)
      .map((row) => toNumber(getRowValue(row, header, table)))
      .filter((value) => value != null);
    return values.length > 0;
  });
}

function buildSurveyLongTable(table = {}) {
  const scoreHeaders = scoreColumnsForSurvey(table);
  if (scoreHeaders.length < 2) return null;
  const rows = [];
  for (const sourceRow of getRows(table)) {
    for (const scoreHeader of scoreHeaders) {
      const score = toNumber(getRowValue(sourceRow, scoreHeader, table));
      if (score == null) continue;
      const row = {};
      for (const header of tableHeaders(table)) {
        if (!scoreHeaders.includes(header)) {
          row[header] = getRowValue(sourceRow, header, table);
        }
      }
      row["문항"] = scoreHeader;
      row["점수"] = score;
      rows.push(row);
    }
  }
  if (!rows.length) return null;
  const headers = [...tableHeaders(table).filter((h) => !scoreHeaders.includes(h)), "문항", "점수"];
  return {
    ...table,
    tableId: `${table.tableId || "table"}#CONTRACT_SURVEY_LONG`,
    sheetName: table.sheetName || table.tableName || "",
    columns: headers.map((header) => ({ header })),
    rows,
    rowCount: rows.length,
    isVirtual: true,
  };
}

function candidateTables(templateId = "", tables = []) {
  const base = (tables || []).filter(Boolean);
  if (templateId !== "survey_satisfaction_analysis") return base;
  return base.flatMap((table) => {
    const transformed = buildSurveyLongTable(table);
    return transformed ? [transformed, table] : [table];
  });
}

function resolveRoleMapping(table = {}, role = {}, templateId = "") {
  if (role.role === "period") {
    const yearHeader = findHeader(table, ["연도", "년도", "year"]);
    const monthHeader = findHeader(table, ["월", "월구분", "month"]);
    if (yearHeader && monthHeader && normalizeHeader(yearHeader) !== normalizeHeader(monthHeader)) {
      return { kind: "compositePeriod", yearHeader, monthHeader };
    }
  }
  const aliases = [
    ...(ROLE_ALIAS_OVERRIDES[templateId]?.[role.role] || []),
    ...(role.aliases || []),
  ];
  const direct = findHeader(table, aliases);
  if (direct) return { kind: "column", header: direct };
  return null;
}

function resolveContractSource(contract = {}, tables = []) {
  let best = null;
  for (const table of candidateTables(contract.templateId, tables)) {
    const mappings = {};
    let requiredResolved = 0;
    let optionalResolved = 0;
    for (const role of contract.sourceRoles || []) {
      const mapping = resolveRoleMapping(table, role, contract.templateId);
      if (mapping) {
        mappings[role.role] = mapping;
        if (role.required === false) optionalResolved += 1;
        else requiredResolved += 1;
      }
    }
    const requiredCount = (contract.sourceRoles || []).filter(
      (role) => role.required !== false,
    ).length;
    const score = requiredResolved * 1000 + optionalResolved * 100 + getRows(table).length;
    if (!best || score > best.score) {
      best = { table, mappings, requiredResolved, requiredCount, score };
    }
  }
  return best;
}

function mappedValue(row = {}, mapping = null, table = {}) {
  if (!mapping) return undefined;
  if (mapping.kind === "column") return getRowValue(row, mapping.header, table);
  if (mapping.kind === "compositePeriod") {
    const year = compactText(getRowValue(row, mapping.yearHeader, table)).match(/(?:19|20)\d{2}/)?.[0] || "";
    const monthRaw = compactText(getRowValue(row, mapping.monthHeader, table));
    const month = monthRaw.match(/^(1[0-2]|0?[1-9])(?:월)?$/)?.[1] || "";
    return year && month ? `${year}-${String(Number(month)).padStart(2, "0")}` : year;
  }
  return undefined;
}

function roleDefinitionMap(contract = {}) {
  return new Map((contract.sourceRoles || []).map((role) => [role.role, role]));
}

function typedRoleValue(row = {}, roleName = "", context = {}) {
  const role = context.roleByName.get(roleName);
  const mapping = context.source.mappings[roleName];
  if (!role || !mapping) return { valid: false, blank: true, value: null };
  const raw = mappedValue(row, mapping, context.source.table);
  if (raw == null || compactText(raw) === "") return { valid: true, blank: true, value: null };
  if (role.dataType === "number") {
    const value = toNumber(raw);
    return { valid: value != null, blank: false, value };
  }
  if (role.dataType === "period") {
    const value = normalizePeriod(raw);
    return { valid: Boolean(value), blank: !value, value };
  }
  if (role.dataType === "boolean") {
    return { valid: true, blank: false, value: toBoolean(raw) };
  }
  return { valid: true, blank: false, value: compactText(raw) };
}

function rowPassesFilter(row = {}, filter = {}, context = {}) {
  if (filter.operator === "truthy") {
    return toBoolean(typedRoleValue(row, filter.role, context).value);
  }
  if (filter.operator === "ltRole") {
    const left = toNumber(typedRoleValue(row, filter.leftRole, context).value);
    const right = toNumber(typedRoleValue(row, filter.rightRole, context).value);
    return left != null && right != null && left < right;
  }
  return true;
}

function filterRows(rows = [], filters = [], context = {}) {
  if (!(filters || []).length) return rows;
  return rows.filter((row) =>
    filters.every((filter) => rowPassesFilter(row, filter, context)),
  );
}

function roundHalfUp(value, digits = 2) {
  if (!Number.isFinite(value)) return value;
  const factor = 10 ** digits;
  return (
    Math.sign(value) *
    Math.round(Math.abs(value) * factor + Number.EPSILON) /
    factor
  );
}

function roundMetricValue(value, rounding = {}) {
  const digits = Number.isInteger(rounding.digits) ? rounding.digits : 2;
  if (typeof value === "number") return roundHalfUp(value, digits);
  if (Array.isArray(value)) return value.map((entry) => roundMetricValue(entry, rounding));
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, nested]) => [
        key,
        roundMetricValue(nested, rounding),
      ]),
    );
  }
  return value;
}

function stableKey(parts = []) {
  return JSON.stringify(parts);
}

function aggregateScalar(metric = {}, rows = [], context = {}) {
  if (metric.aggregation === "countRows") return rows.length;
  if (metric.aggregation === "countDistinct") {
    const values = rows
      .map((row) => typedRoleValue(row, metric.valueRole, context))
      .filter((entry) => entry.valid && !entry.blank)
      .map((entry) => stableKey([entry.value]));
    return new Set(values).size;
  }
  const values = rows
    .map((row) => typedRoleValue(row, metric.valueRole, context))
    .filter((entry) => entry.valid && !entry.blank)
    .map((entry) => Number(entry.value));
  if (!values.length) return null;
  if (metric.aggregation === "sum") return values.reduce((sum, value) => sum + value, 0);
  if (metric.aggregation === "average") {
    return values.reduce((sum, value) => sum + value, 0) / values.length;
  }
  return null;
}

function computeAggregate(metric = {}, context = {}) {
  const rows = filterRows(
    context.source.table.rows || [],
    metric.filters || [],
    context,
  );
  const groupByRoles = metric.groupByRoles || [];
  if (!groupByRoles.length) {
    return { valueType: "scalar", value: aggregateScalar(metric, rows, context) };
  }
  const groups = new Map();
  rows.forEach((row, sourceOrder) => {
    const parts = groupByRoles.map((roleName) => typedRoleValue(row, roleName, context));
    if (parts.some((part) => !part.valid || part.blank)) return;
    const keyParts = parts.map((part) => part.value);
    const key = stableKey(keyParts);
    if (!groups.has(key)) groups.set(key, { key: keyParts, rows: [], sourceOrder });
    groups.get(key).rows.push(row);
  });
  return {
    valueType: "grouped",
    groupByRoles,
    entries: [...groups.values()].map((group) => ({
      key: group.key,
      keyObject: Object.fromEntries(groupByRoles.map((role, index) => [role, group.key[index]])),
      value: aggregateScalar(metric, group.rows, context),
      sourceOrder: group.sourceOrder,
    })),
  };
}

function groupedMap(value = {}) {
  return new Map((value.entries || []).map((entry) => [stableKey(entry.key), entry]));
}

function computeDerived(metric = {}, context = {}) {
  if (metric.operator === "filterCount") {
    return {
      valueType: "scalar",
      value: filterRows(
        context.source.table.rows || [],
        metric.filters || [],
        context,
      ).length,
    };
  }
  const operands = (metric.operands || []).map((operand) => {
    if (operand.type === "metric") return context.metricValues.get(operand.ref) || null;
    if (operand.type === "constant") return { valueType: "scalar", value: operand.value };
    return null;
  });
  if (operands.length < 2 || operands.some((operand) => !operand)) return null;
  const [left, right] = operands;
  const operation = (a, b) => {
    if (a == null || b == null) return null;
    if (metric.operator === "subtract") return a - b;
    if (metric.operator === "percentage") return b === 0 ? null : (a / b) * 100;
    return null;
  };
  if (left.valueType === "scalar" && right.valueType === "scalar") {
    return { valueType: "scalar", value: operation(left.value, right.value) };
  }
  if (left.valueType === "grouped" && right.valueType === "grouped") {
    const rightByKey = groupedMap(right);
    return {
      valueType: "grouped",
      groupByRoles: metric.groupByRoles || left.groupByRoles || [],
      entries: (left.entries || [])
        .filter((entry) => rightByKey.has(stableKey(entry.key)))
        .map((entry) => ({
          ...entry,
          value: operation(entry.value, rightByKey.get(stableKey(entry.key)).value),
        })),
    };
  }
  if (left.valueType === "grouped" && right.valueType === "scalar") {
    return {
      valueType: "grouped",
      groupByRoles: metric.groupByRoles || left.groupByRoles || [],
      entries: (left.entries || []).map((entry) => ({
        ...entry,
        value: operation(entry.value, right.value),
      })),
    };
  }
  if (left.valueType === "scalar" && right.valueType === "grouped") {
    return {
      valueType: "grouped",
      groupByRoles: metric.groupByRoles || right.groupByRoles || [],
      entries: (right.entries || []).map((entry) => ({
        ...entry,
        value: operation(left.value, entry.value),
      })),
    };
  }
  return null;
}

function computeRank(metric = {}, context = {}) {
  const source = context.metricValues.get(metric.sourceMetricId);
  if (!source || source.valueType !== "grouped") return null;
  const direction = metric.direction === "asc" ? 1 : -1;
  const sorted = [...(source.entries || [])].sort((a, b) => {
    const delta = (Number(a.value) - Number(b.value)) * direction;
    return delta || Number(a.sourceOrder || 0) - Number(b.sourceOrder || 0);
  });
  const limit = Number(metric.limit || 1);
  if (!sorted.length) return { valueType: "rank", sourceMetricId: metric.sourceMetricId, direction: metric.direction, limit, items: [] };
  const cutoff = sorted[Math.min(limit, sorted.length) - 1]?.value;
  const items = sorted
    .filter((entry, index) => index < limit || Number(entry.value) === Number(cutoff))
    .map((entry, index) => ({ ...entry, rank: index + 1 }));
  return { valueType: "rank", sourceMetricId: metric.sourceMetricId, direction: metric.direction, limit, items };
}

function metricActive(metric = {}, context = {}) {
  const requiresRoles = metric.activation?.requiresRoles || [];
  if (requiresRoles.some((role) => !context.source.mappings[role])) return false;
  const requiresMetrics = metric.activation?.requiresMetricIds || [];
  return requiresMetrics.every((metricId) => context.metricValues.has(metricId));
}

function computeContractMetrics(contract = {}, source = {}) {
  const context = {
    contract,
    source,
    roleByName: roleDefinitionMap(contract),
    metricValues: new Map(),
  };
  const metrics = [];
  const pending = [...(contract.metrics || [])];
  let guard = pending.length + 5;
  while (pending.length && guard-- > 0) {
    let progressed = false;
    for (let index = 0; index < pending.length; index += 1) {
      const metric = pending[index];
      const dependencies = metric.activation?.requiresMetricIds || [];
      if (dependencies.some((metricId) => pending.some((candidate) => candidate.metricId === metricId))) {
        continue;
      }
      pending.splice(index, 1);
      index -= 1;
      progressed = true;
      if (!metricActive(metric, context)) {
        metrics.push({ metric, status: "INACTIVE", value: null });
        continue;
      }
      let value = null;
      if (metric.kind === "aggregate") value = computeAggregate(metric, context);
      else if (metric.kind === "derived") value = computeDerived(metric, context);
      else if (metric.kind === "rank") value = computeRank(metric, context);
      if (!value) {
        metrics.push({ metric, status: "ERROR", value: null });
        continue;
      }
      const roundedValue = roundMetricValue(value, metric.rounding || {});
      context.metricValues.set(metric.metricId, roundedValue);
      metrics.push({ metric, status: "COMPUTED", value: roundedValue });
    }
    if (!progressed) break;
  }
  return metrics;
}

const ROLE_DISPLAY_LABELS = Object.freeze({
  period: "기간",
  product: "품목",
  department: "부서",
  question: "문항",
  program: "프로그램",
  warehouse: "창고",
  item: "품목",
  category: "구분",
  status: "상태",
  requestingOrganization: "요청기관",
  institution: "연구기관",
  project: "과제",
  facility: "시설",
  branch: "지점",
  position: "직급",
  employeeStatus: "재직상태",
});

function roleDisplayLabel(role = "", source = {}) {
  const mapping = source.mappings?.[role];
  return mapping?.header || ROLE_DISPLAY_LABELS[role] || role || "구분";
}

function safeSectionId(metricId = "") {
  return `contract_${String(metricId || "metric").replace(/[^a-zA-Z0-9_]+/g, "_")}`;
}

function scalarCoverageSection(computed = []) {
  if (!computed.length) return null;
  const metricIds = computed.map((entry) => entry.metric.metricId);
  return {
    sectionId: "contract_summary_kpi_coverage",
    sectionType: "contract_metric_coverage",
    title: "계약 핵심지표",
    metricIds,
    result: {
      ok: true,
      resultType: "pivot",
      operation: "contractScalarCoverage",
      rows: computed.map((entry) => ({
        지표: entry.metric.label,
        값: entry.value.value,
        metricId: entry.metric.metricId,
      })),
      rowCount: computed.length,
      meta: {
        contractCoverageVersion: CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION,
        metricIds,
        complete: true,
      },
    },
  };
}

function groupedCoverageSection(entry = {}, source = {}) {
  const metric = entry.metric;
  const value = entry.value;
  const roles = value.groupByRoles || metric.groupByRoles || [];
  const groupHeaders = roles.map((role) => roleDisplayLabel(role, source));
  const rows = (value.entries || []).map((item) => {
    const row = {};
    item.key.forEach((key, index) => {
      row[groupHeaders[index] || `구분${index + 1}`] = key;
    });
    row[metric.label] = item.value;
    row.metricId = metric.metricId;
    return row;
  });
  return {
    sectionId: safeSectionId(metric.metricId),
    sectionType: "contract_metric_coverage",
    title: metric.label,
    metricIds: [metric.metricId],
    result: {
      ok: true,
      resultType: "pivot",
      operation: `contract${metric.aggregation || "Grouped"}`,
      groupBy: { header: groupHeaders[0] || "구분" },
      metric: { header: metric.label },
      rows,
      rowCount: rows.length,
      meta: {
        contractCoverageVersion: CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION,
        metricIds: [metric.metricId],
        complete: true,
      },
    },
  };
}

function rankCoverageSection(entry = {}, source = {}) {
  const metric = entry.metric;
  const value = entry.value;
  const sourceMetric = value.sourceMetricId || metric.sourceMetricId;
  const sourceValue = sourceMetric ? null : null;
  const role = (value.items?.[0]?.keyObject && Object.keys(value.items[0].keyObject)[0]) || "item";
  const groupHeader = roleDisplayLabel(role, source);
  const rows = (value.items || []).map((item) => ({
    [groupHeader]: item.key?.[0] ?? "",
    [metric.label]: item.value,
    순위: item.rank,
    metricId: metric.metricId,
  }));
  return {
    sectionId: safeSectionId(metric.metricId),
    sectionType: "contract_metric_rank_coverage",
    title: metric.label,
    metricIds: [metric.metricId],
    result: {
      ok: true,
      resultType: "pivot",
      operation: "contractRank",
      groupBy: { header: groupHeader },
      metric: { header: metric.label },
      rows,
      rowCount: rows.length,
      meta: {
        contractCoverageVersion: CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION,
        metricIds: [metric.metricId],
        sourceMetricId: sourceMetric,
        complete: true,
      },
    },
  };
}

function sectionsFromComputedMetrics(computed = [], source = {}) {
  const active = computed.filter((entry) => entry.status === "COMPUTED");
  const scalar = active.filter((entry) => entry.value?.valueType === "scalar");
  const grouped = active.filter((entry) => entry.value?.valueType === "grouped");
  const rank = active.filter((entry) => entry.value?.valueType === "rank");
  return [
    scalarCoverageSection(scalar),
    ...grouped.map((entry) => groupedCoverageSection(entry, source)),
    ...rank.map((entry) => rankCoverageSection(entry, source)),
  ].filter(Boolean);
}

function buildContractDrivenSummarySections({
  normalizedQueryTables = [],
  templateId = "",
} = {}) {
  const contract = chooseContract(templateId, normalizedQueryTables);
  if (!contract) {
    return {
      version: CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION,
      contractCatalogVersion: catalog.version,
      templateId,
      status: "UNSUPPORTED",
      sections: [],
      expectedMetricIds: [],
      renderedMetricIds: [],
      inactiveMetricIds: [],
      errorMetricIds: [],
    };
  }
  const source = resolveContractSource(contract, normalizedQueryTables);
  if (!source?.table) {
    return {
      version: CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION,
      contractCatalogVersion: catalog.version,
      templateId,
      contractId: contract.contractId,
      status: "SOURCE_UNRESOLVED",
      sections: [],
      expectedMetricIds: [],
      renderedMetricIds: [],
      inactiveMetricIds: [],
      errorMetricIds: [],
    };
  }
  const computed = computeContractMetrics(contract, source);
  const sections = sectionsFromComputedMetrics(computed, source);
  const renderedMetricIds = computed
    .filter((entry) => entry.status === "COMPUTED")
    .map((entry) => entry.metric.metricId);
  const inactiveMetricIds = computed
    .filter((entry) => entry.status === "INACTIVE")
    .map((entry) => entry.metric.metricId);
  const errorMetricIds = computed
    .filter((entry) => entry.status === "ERROR")
    .map((entry) => entry.metric.metricId);
  return {
    version: CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION,
    contractCatalogVersion: catalog.version,
    sourceContractsVersion: catalog.sourceContractsVersion,
    templateId,
    contractId: contract.contractId,
    status: errorMetricIds.length ? "PARTIAL" : "PASS",
    selectedTableId: source.table.tableId || "",
    selectedSheetName: source.table.sheetName || "",
    expectedMetricIds: renderedMetricIds,
    renderedMetricIds,
    inactiveMetricIds,
    errorMetricIds,
    resolvedRoleIds: Object.keys(source.mappings || {}),
    sections,
  };
}

module.exports = {
  CONTRACT_DRIVEN_SUMMARY_RECIPE_VERSION,
  normalizeHeader,
  normalizePeriod,
  chooseContract,
  resolveContractSource,
  computeContractMetrics,
  buildContractDrivenSummarySections,
};
