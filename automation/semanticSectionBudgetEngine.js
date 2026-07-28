"use strict";

const SEMANTIC_SECTION_BUDGET_ENGINE_VERSION =
  "semantic_section_budget_engine_v2_mandatory_summary_coverage_floor";
const MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION =
  "mandatory_summary_coverage_floor_v1";
const DURATION_SUMMARY_CONTRACT_VERSION =
  "duration_summary_contract_v1_average_median_range";
const DISTINCT_ENTITY_SECTION_VERSION =
  "distinct_entity_section_v1";

const ENTITY_HEADER_PATTERN =
  /(?:^|[_\s])(?:id|code)(?:$|[_\s])|(?:품목|소모품|제품|상품|자산|장비|시설|고객|거래처|업체|기관|사업|프로젝트|과제|서비스|항목)(?:명|이름)$|^(?:이름|명칭|entity|item|product|asset|equipment|facility|customer|vendor|project)$/i;
const EXCLUDED_ENTITY_HEADER_PATTERN =
  /상태|구분|분류|유형|결과|등급|채널|지역|기간|연월|월|일자|날짜|연도|년도|단위|비고|설명|status|category|type|result|grade|channel|region|period|date|month|year|unit|note/i;

function normalizeText(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKey(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "");
}

function cloneValue(value) {
  return JSON.parse(JSON.stringify(value));
}

function median(values = []) {
  const numbers = values
    .filter((value) => typeof value === "number" && Number.isFinite(value))
    .sort((left, right) => left - right);
  if (!numbers.length) return null;
  const middle = Math.floor(numbers.length / 2);
  return numbers.length % 2
    ? numbers[middle]
    : (numbers[middle - 1] + numbers[middle]) / 2;
}

function sectionPolicyForSeries(series = {}, options = {}) {
  const role = normalizeText(series.metricRole);
  const relationRole = normalizeText(series.relationshipRole || "independent");
  const globalMax = Math.max(0, Number(options.maxDimensionsPerSeries ?? 8));
  let maxDimensions = 3;
  let includePeriod = true;
  let maxSections = 5;
  let budgetPriority = Number(series.representativeMetricPriority || 400);

  if (relationRole === "primary_total") {
    maxDimensions = 4;
    maxSections = 6;
    budgetPriority = Math.max(budgetPriority, 1200);
  } else if (relationRole === "primary_basis") {
    maxDimensions = 4;
    maxSections = 6;
    budgetPriority = Math.max(budgetPriority, 980);
  } else if (relationRole === "component") {
    maxDimensions = 1;
    maxSections = 3;
    budgetPriority = Math.min(budgetPriority, 350);
  } else if (relationRole === "unit_component") {
    maxDimensions = 1;
    includePeriod = true;
    maxSections = 3;
    budgetPriority = Math.min(budgetPriority, 260);
  } else if (role === "duration") {
    maxDimensions = 2;
    maxSections = 4;
    budgetPriority = Math.max(budgetPriority, 470);
  } else if (role === "unit_rate" || role === "percentage_rate") {
    maxDimensions = 2;
    maxSections = 4;
  } else if (role === "stock_snapshot") {
    maxDimensions = 4;
    maxSections = 6;
    budgetPriority = Math.max(budgetPriority, 700);
  }

  return {
    version: SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
    maxDimensions: Math.min(globalMax, maxDimensions),
    includePeriod,
    maxSections,
    budgetPriority,
    role,
    relationRole,
  };
}

function sectionPriority(section = {}) {
  const meta = section.result?.meta || {};
  let score = Number(meta.sectionBudgetPriority || 400);
  const type = normalizeText(section.sectionType);
  if (/summary/.test(type)) score += 180;
  if (/distinct/.test(type)) score += 170;
  if (/period/.test(type)) score += 60;
  if (/group/.test(type)) score += 30;
  if (meta.relationshipRole === "primary_total") score += 250;
  if (meta.relationshipRole === "component") score -= 80;
  return score;
}

function applySemanticSectionBudget({ sections = [], maxSections = 28 } = {}) {
  const list = Array.isArray(sections) ? sections : [];
  const cap = Math.max(0, Number(maxSections ?? 28));
  if (!Number.isFinite(cap) || list.length <= cap) {
    return {
      version: SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
      sections: list,
      applied: false,
      inputSectionCount: list.length,
      retainedSectionCount: list.length,
      droppedSectionCount: 0,
      droppedSectionIds: [],
      maxSections: cap,
      effectiveMaxSections: cap,
      mandatorySectionCount: 0,
      mandatoryOverflowCount: 0,
      mandatorySummaryCoverageFloorVersion:
        MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
    };
  }

  const decorated = list.map((section, index) => ({
    section,
    index,
    score: sectionPriority(section),
    metricKey: normalizeKey(section.result?.metric?.header || ""),
    mandatory:
      /summary|distinct/.test(normalizeText(section.sectionType)) ||
      section.result?.operation === "semanticDistinctCount",
  }));

  const mandatoryIndexes = new Set();
  const summaryByMetric = new Map();
  for (const item of decorated) {
    if (item.mandatory) mandatoryIndexes.add(item.index);
    if (!item.metricKey || !/summary/.test(normalizeText(item.section.sectionType))) {
      continue;
    }
    const current = summaryByMetric.get(item.metricKey);
    if (!current || item.score > current.score) summaryByMetric.set(item.metricKey, item);
  }
  for (const item of summaryByMetric.values()) mandatoryIndexes.add(item.index);

  const effectiveCap = Math.max(cap, mandatoryIndexes.size);
  const retainedIndexes = new Set(mandatoryIndexes);
  const ranked = decorated
    .filter((item) => !retainedIndexes.has(item.index))
    .sort((left, right) =>
      right.score - left.score || left.index - right.index,
    );
  for (const item of ranked) {
    if (retainedIndexes.size >= effectiveCap) break;
    retainedIndexes.add(item.index);
  }

  const retained = list.filter((_, index) => retainedIndexes.has(index));
  const dropped = list.filter((_, index) => !retainedIndexes.has(index));
  return {
    version: SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
    sections: retained,
    applied: true,
    inputSectionCount: list.length,
    retainedSectionCount: retained.length,
    droppedSectionCount: dropped.length,
    droppedSectionIds: dropped.map((section, index) =>
      normalizeText(section.sectionId || section.title || `section_${index + 1}`),
    ),
    maxSections: cap,
    effectiveMaxSections: effectiveCap,
    mandatorySectionCount: mandatoryIndexes.size,
    mandatoryOverflowCount: Math.max(0, mandatoryIndexes.size - cap),
    mandatorySummaryCoverageFloorVersion:
      MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
  };
}

function tableColumns(table = {}) {
  return Array.isArray(table.columns) ? table.columns : [];
}

function tableRows(table = {}) {
  return Array.isArray(table.rows) ? table.rows : [];
}

function columnHeader(column = {}, index = 0) {
  return normalizeText(
    column.header || column.originalHeader || column.name || column.label || `열${index + 1}`,
  );
}

function rowValue(row, column = {}, index = 0) {
  if (Array.isArray(row)) return row[index];
  if (!row || typeof row !== "object") return undefined;
  const keys = [
    column.key,
    column.canonicalKey,
    column.accessor,
    column.name,
    column.header,
    column.originalHeader,
    column.label,
    column.id,
  ].map((value) => String(value || "").trim()).filter(Boolean);
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
  }
  const targets = new Set(keys.map(normalizeKey));
  for (const [key, value] of Object.entries(row)) {
    if (targets.has(normalizeKey(key))) return value;
  }
  return Object.values(row)[index];
}

function resolveDistinctEntity(table = {}) {
  const rows = tableRows(table);
  const candidates = tableColumns(table).map((column, index) => {
    const header = columnHeader(column, index);
    const values = rows
      .map((row) => normalizeText(rowValue(row, column, index)))
      .filter(Boolean);
    const distinctCount = new Set(values.map(normalizeKey)).size;
    const coverage = rows.length ? values.length / rows.length : 0;
    const identity = ENTITY_HEADER_PATTERN.test(header);
    const excluded = EXCLUDED_ENTITY_HEADER_PATTERN.test(header);
    let score = identity ? 200 : -Infinity;
    if (excluded || coverage < 0.75 || distinctCount < 2 || distinctCount > 500) {
      score = -Infinity;
    }
    if (Number.isFinite(score)) {
      if (/명$|이름$|명칭$/i.test(header)) score += 40;
      if (coverage >= 0.95) score += 25;
      if (distinctCount === rows.length) score += 15;
      if (distinctCount <= 100) score += 15;
    }
    return { column, index, header, values, distinctCount, coverage, score };
  }).filter((candidate) => Number.isFinite(candidate.score))
    .sort((left, right) =>
      right.score - left.score ||
      right.distinctCount - left.distinctCount ||
      left.header.localeCompare(right.header, "ko"),
    );

  const selected = candidates[0];
  if (!selected) {
    return {
      version: DISTINCT_ENTITY_SECTION_VERSION,
      applied: false,
      reason: "no_distinct_entity_header",
    };
  }
  return {
    version: DISTINCT_ENTITY_SECTION_VERSION,
    applied: true,
    header: selected.header,
    distinctCount: selected.distinctCount,
    rowCount: rows.length,
    duplicateRowCount: Math.max(0, rows.length - selected.distinctCount),
    coverage: selected.coverage,
  };
}

function buildDistinctEntitySection({ table = {}, tableIndex = 0, metricIdFactory } = {}) {
  const resolution = resolveDistinctEntity(table);
  if (!resolution.applied) return null;
  const id = typeof metricIdFactory === "function"
    ? metricIdFactory(tableIndex, `고유 ${resolution.header} 수`, "건", "distinct")
    : `semantic.table_${tableIndex + 1}.distinct.${normalizeKey(resolution.header)}`;
  return {
    sectionId: id,
    title: `고유 ${resolution.header} 수`,
    sectionType: "semantic_distinct_entity_summary",
    metricIds: [id],
    result: {
      ok: true,
      resultType: "pivot",
      operation: "semanticDistinctCount",
      metric: { header: `고유 ${resolution.header} 수`, unit: "건" },
      rows: [
        { 지표: "전체 행 수", 값: resolution.rowCount, 단위: "건" },
        { 지표: `고유 ${resolution.header} 수`, 값: resolution.distinctCount, 단위: "건" },
        { 지표: "중복 행 수", 값: resolution.duplicateRowCount, 단위: "건" },
      ],
      meta: {
        metricIds: [id],
        complete: true,
        distinctEntitySectionVersion: DISTINCT_ENTITY_SECTION_VERSION,
        distinctEntityHeader: resolution.header,
        distinctEntityCount: resolution.distinctCount,
        distinctEntityCoverage: resolution.coverage,
        sectionBudgetPriority: 850,
      },
    },
  };
}

module.exports = {
  SEMANTIC_SECTION_BUDGET_ENGINE_VERSION,
  MANDATORY_SUMMARY_COVERAGE_FLOOR_VERSION,
  DURATION_SUMMARY_CONTRACT_VERSION,
  DISTINCT_ENTITY_SECTION_VERSION,
  applySemanticSectionBudget,
  buildDistinctEntitySection,
  median,
  resolveDistinctEntity,
  sectionPolicyForSeries,
};
