"use strict";

const crypto = require("crypto");
const {
  applySectionMetricIds,
  collectSectionMetricIds,
  normalizeSectionMetricIds,
  uniqueMetricIds,
} = require("./metricIdContract");

const OUTPUT_COMPLETENESS_CONTRACT_VERSION =
  "output_completeness_contract_v1";
const SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION =
  "semantic_output_duplicate_resolver_v1";
const FINAL_OUTPUT_QUALITY_GATE_VERSION =
  "final_output_quality_gate_v1";

const PERIOD_HEADER_PATTERN =
  /기간|년월|연월|월|일자|날짜|연도|년도|date|period|month|quarter|year/i;
const SUMMARY_TYPE_PATTERN = /summary|overview|snapshot|distinct/i;
const DIAGNOSTIC_TYPE_PATTERN =
  /diagnostic|preview|semanticSource|sourceData|rawData/i;

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

function sectionRows(section = {}) {
  const result = section.result || {};
  const candidates = [result.rows, result.data, result.items, result.records];
  for (const candidate of candidates) {
    if (Array.isArray(candidate)) return candidate;
  }
  return [];
}

function sectionMetricHeader(section = {}) {
  return normalizeText(
    section.result?.metric?.header ||
      section.metricHeader ||
      section.metric ||
      "",
  );
}

function sectionGroupHeader(section = {}) {
  return normalizeText(
    section.result?.groupBy?.header ||
      section.groupHeader ||
      section.groupBy ||
      "",
  );
}

function sectionUnit(section = {}) {
  return normalizeText(
    section.result?.metric?.unit ||
      section.unit ||
      "",
  );
}

function sectionOperation(section = {}) {
  return normalizeText(
    section.result?.operation ||
      section.operation ||
      "",
  );
}

function sectionType(section = {}) {
  return normalizeText(section.sectionType || section.resultType || "");
}

function sectionScope(section = {}) {
  const type = sectionType(section);
  const operation = sectionOperation(section);
  const group = sectionGroupHeader(section);

  if (/distinct/i.test(type) || /distinct/i.test(operation)) {
    return "distinct";
  }
  if (
    group &&
    (PERIOD_HEADER_PATTERN.test(group) || /period|time/i.test(type))
  ) {
    return "period";
  }
  if (group) return "group";
  if (
    SUMMARY_TYPE_PATTERN.test(type) ||
    /summary|overview|snapshot|aggregate/i.test(operation)
  ) {
    return "summary";
  }
  return normalizeKey(type || operation || "section") || "section";
}

function operationFamily(section = {}) {
  const operation = normalizeKey(sectionOperation(section));
  if (!operation) return "";
  if (/latest|snapshot/.test(operation)) return "latest";
  if (/average|avg|mean|median/.test(operation)) return "average";
  if (/distinct/.test(operation)) return "distinct";
  if (/count/.test(operation)) return "count";
  if (/ratio|rate|percent|composition/.test(operation)) return "ratio";
  if (/flow|direction|movement|ledger/.test(operation)) return "flow";
  if (/sum|aggregate|total/.test(operation)) return "sum";
  return operation;
}

function canonicalizeValue(value) {
  if (Array.isArray(value)) {
    return value.map(canonicalizeValue);
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.keys(value)
        .sort((left, right) => left.localeCompare(right, "ko"))
        .map((key) => [normalizeText(key), canonicalizeValue(value[key])]),
    );
  }
  if (typeof value === "string") return normalizeText(value);
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return String(value);
    return Number(value.toPrecision(15));
  }
  return value;
}

function sectionRowsHash(section = {}) {
  return crypto
    .createHash("sha256")
    .update(JSON.stringify(canonicalizeValue(sectionRows(section))))
    .digest("hex");
}

function sectionSemanticKey(section = {}) {
  return [
    normalizeKey(sectionMetricHeader(section)),
    normalizeKey(sectionGroupHeader(section)),
    operationFamily(section),
    sectionScope(section),
    normalizeKey(sectionUnit(section)),
  ].join("|");
}

function duplicateSignature(section = {}) {
  const semanticKey = sectionSemanticKey(section);
  const rows = sectionRows(section);
  if (!semanticKey || !rows.length) return "";
  return `${semanticKey}|${sectionRowsHash(section)}`;
}

function isWholeMetricSummary(section = {}) {
  const scope = sectionScope(section);
  if (scope !== "summary") return false;
  return Boolean(sectionMetricHeader(section));
}

function sectionComplete(section = {}) {
  if (section.complete === false) return false;
  if (section.result?.complete === false) return false;
  if (section.result?.meta?.complete === false) return false;
  return true;
}

function sectionPreferenceScore(section = {}, expectedMetricIds = new Set()) {
  const metricIds = collectSectionMetricIds(section);
  let score = metricIds.filter((id) => expectedMetricIds.has(id)).length * 1000;
  if (sectionComplete(section)) score += 200;
  if (isWholeMetricSummary(section)) score += 140;
  if (sectionScope(section) === "distinct") score += 120;
  if (sectionRows(section).length) score += 80;
  if (section.result?.meta?.semanticCoverage?.matchedExistingSection) score += 30;
  score += Math.min(sectionRows(section).length, 50);
  return score;
}

function setSectionMetricIds(section = {}, metricIds = []) {
  const ids = uniqueMetricIds(metricIds);
  const result = section.result || {};
  const meta = result.meta || {};
  return {
    ...section,
    metricIds: ids,
    result: {
      ...result,
      metricIds: ids,
      meta: {
        ...meta,
        metricIds: ids,
      },
    },
  };
}

function resolveDuplicateSections({ sections = [], expectedMetricIds = [] } = {}) {
  const expectedSet = new Set(uniqueMetricIds(expectedMetricIds));
  const input = normalizeSectionMetricIds(cloneValue(Array.isArray(sections) ? sections : []));
  const groups = new Map();

  input.forEach((section, index) => {
    const signature = duplicateSignature(section);
    if (!signature) return;
    if (!groups.has(signature)) groups.set(signature, []);
    groups.get(signature).push({ section, index });
  });

  const removedIndexes = new Set();
  const replacementByIndex = new Map();
  const duplicateGroups = [];
  let mergedMetricIdCount = 0;

  for (const [signature, items] of groups.entries()) {
    if (items.length < 2) continue;
    const ranked = [...items].sort((left, right) => {
      const scoreDelta =
        sectionPreferenceScore(right.section, expectedSet) -
        sectionPreferenceScore(left.section, expectedSet);
      return scoreDelta || left.index - right.index;
    });
    const retained = ranked[0];
    const allMetricIds = uniqueMetricIds(
      ranked.flatMap((item) => collectSectionMetricIds(item.section)),
    );
    const retainedMetricIds = collectSectionMetricIds(retained.section);
    mergedMetricIdCount += allMetricIds.filter(
      (metricId) => !retainedMetricIds.includes(metricId),
    ).length;
    replacementByIndex.set(
      retained.index,
      applySectionMetricIds(retained.section, allMetricIds),
    );
    for (const item of ranked.slice(1)) removedIndexes.add(item.index);
    duplicateGroups.push({
      signature,
      retainedSectionId: normalizeText(
        retained.section.sectionId || retained.section.title || "",
      ),
      removedSectionIds: ranked.slice(1).map((item) =>
        normalizeText(item.section.sectionId || item.section.title || ""),
      ),
      mergedMetricIds: allMetricIds,
    });
  }

  const output = input
    .map((section, index) => replacementByIndex.get(index) || section)
    .filter((_, index) => !removedIndexes.has(index));

  return {
    version: SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
    sections: output,
    applied: removedIndexes.size > 0,
    inputSectionCount: input.length,
    outputSectionCount: output.length,
    removedDuplicateSectionCount: removedIndexes.size,
    removedDuplicateSectionIds: duplicateGroups.flatMap(
      (group) => group.removedSectionIds,
    ),
    duplicateGroups,
    mergedMetricIdCount,
  };
}

function metricIdExpectedScope(metricId = "") {
  const normalized = normalizeText(metricId).toLowerCase();
  if (/\.distinct(?:\.|$)/.test(normalized)) return "distinct";
  if (/\.by_period(?:\.|$)/.test(normalized)) return "period";
  if (/\.by_[^.]+(?:\.|$)/.test(normalized)) return "group";
  if (/\.summary(?:\.|$)/.test(normalized)) return "summary";
  return "";
}

function metricHolderScore(section = {}, metricId = "", expectedSet = new Set()) {
  let score = sectionPreferenceScore(section, expectedSet);
  const expectedScope = metricIdExpectedScope(metricId);
  if (expectedScope && sectionScope(section) === expectedScope) score += 900;
  if (expectedScope === "summary" && isWholeMetricSummary(section)) score += 350;
  if (expectedScope === "group" && sectionGroupHeader(section)) score += 220;
  if (expectedScope === "period" && PERIOD_HEADER_PATTERN.test(sectionGroupHeader(section))) {
    score += 220;
  }
  return score;
}

function normalizeMetricIdOwnership({ sections = [], expectedMetricIds = [] } = {}) {
  const expectedSet = new Set(uniqueMetricIds(expectedMetricIds));
  const working = normalizeSectionMetricIds(cloneValue(Array.isArray(sections) ? sections : []));
  const holders = new Map();

  working.forEach((section, index) => {
    for (const metricId of collectSectionMetricIds(section)) {
      if (!holders.has(metricId)) holders.set(metricId, []);
      holders.get(metricId).push(index);
    }
  });

  const duplicateMetricIdsBefore = [];
  const reassignedMetricIds = [];
  let removedOwnershipCount = 0;

  for (const [metricId, indexes] of holders.entries()) {
    if (indexes.length < 2) continue;
    duplicateMetricIdsBefore.push(metricId);
    const ranked = [...indexes].sort((left, right) => {
      const scoreDelta =
        metricHolderScore(working[right], metricId, expectedSet) -
        metricHolderScore(working[left], metricId, expectedSet);
      return scoreDelta || left - right;
    });
    const winner = ranked[0];
    for (const index of ranked.slice(1)) {
      const ids = collectSectionMetricIds(working[index]).filter(
        (candidate) => candidate !== metricId,
      );
      working[index] = setSectionMetricIds(working[index], ids);
      removedOwnershipCount += 1;
    }
    reassignedMetricIds.push({
      metricId,
      retainedSectionId: normalizeText(
        working[winner].sectionId || working[winner].title || "",
      ),
      removedHolderSectionIds: ranked.slice(1).map((index) =>
        normalizeText(working[index].sectionId || working[index].title || ""),
      ),
    });
  }

  return {
    sections: working,
    applied: duplicateMetricIdsBefore.length > 0,
    duplicateMetricIdsBefore,
    reassignedMetricIds,
    removedOwnershipCount,
  };
}

function defaultSectionTitle(section = {}, index = 0) {
  const metric = sectionMetricHeader(section);
  const group = sectionGroupHeader(section);
  if (metric && group) return `${group}별 ${metric}`;
  if (metric) return `${metric} 분석`;
  return `분석 ${index + 1}`;
}

function ensureUniqueSectionIdentity(sections = []) {
  const sectionIdCounts = new Map();
  const titleCounts = new Map();
  const renamedSectionIds = [];
  const renamedTitles = [];

  const output = (Array.isArray(sections) ? sections : []).map(
    (original, index) => {
      const section = cloneValue(original);
      const rawId = normalizeText(section.sectionId) || `final_quality_section_${index + 1}`;
      const idKey = normalizeKey(rawId) || `section${index + 1}`;
      const idCount = (sectionIdCounts.get(idKey) || 0) + 1;
      sectionIdCounts.set(idKey, idCount);
      const resolvedId = idCount === 1 ? rawId : `${rawId}__${idCount}`;
      if (resolvedId !== normalizeText(section.sectionId)) {
        renamedSectionIds.push({
          from: normalizeText(section.sectionId),
          to: resolvedId,
        });
      }
      section.sectionId = resolvedId;

      const rawTitle = normalizeText(section.title) || defaultSectionTitle(section, index);
      const titleKey = normalizeKey(rawTitle) || `title${index + 1}`;
      const titleCount = (titleCounts.get(titleKey) || 0) + 1;
      titleCounts.set(titleKey, titleCount);
      const resolvedTitle = titleCount === 1 ? rawTitle : `${rawTitle} ${titleCount}`;
      if (resolvedTitle !== normalizeText(section.title)) {
        renamedTitles.push({
          from: normalizeText(section.title),
          to: resolvedTitle,
        });
      }
      section.title = resolvedTitle;
      return section;
    },
  );

  return {
    sections: output,
    renamedSectionIds,
    renamedTitles,
  };
}

function renderedMetricIds(sections = []) {
  return uniqueMetricIds(
    (Array.isArray(sections) ? sections : []).flatMap((section) =>
      collectSectionMetricIds(section),
    ),
  );
}

function duplicateMetricIdOwners(sections = []) {
  const holders = new Map();
  (Array.isArray(sections) ? sections : []).forEach((section, index) => {
    for (const metricId of collectSectionMetricIds(section)) {
      if (!holders.has(metricId)) holders.set(metricId, []);
      holders.get(metricId).push(index);
    }
  });
  return [...holders.entries()]
    .filter(([, indexes]) => indexes.length > 1)
    .map(([metricId, indexes]) => ({
      metricId,
      sectionIds: indexes.map((index) =>
        normalizeText(
          sections[index]?.sectionId || sections[index]?.title || `section_${index + 1}`,
        ),
      ),
    }));
}

function duplicateIdentityValues(sections = [], selector) {
  const values = new Map();
  (Array.isArray(sections) ? sections : []).forEach((section, index) => {
    const raw = normalizeText(selector(section, index));
    const key = normalizeKey(raw);
    if (!key) return;
    if (!values.has(key)) values.set(key, []);
    values.get(key).push(raw);
  });
  return [...values.values()].filter((items) => items.length > 1);
}

function invalidNumberEntries(value, path = "") {
  const found = [];
  if (typeof value === "number" && !Number.isFinite(value)) {
    found.push({ path, value: String(value) });
    return found;
  }
  if (Array.isArray(value)) {
    value.forEach((item, index) => {
      found.push(...invalidNumberEntries(item, `${path}[${index}]`));
    });
    return found;
  }
  if (value && typeof value === "object") {
    for (const [key, item] of Object.entries(value)) {
      found.push(
        ...invalidNumberEntries(item, path ? `${path}.${key}` : key),
      );
    }
  }
  return found;
}

function duplicateSemanticOutputs(sections = []) {
  const groups = new Map();
  (Array.isArray(sections) ? sections : []).forEach((section, index) => {
    const signature = duplicateSignature(section);
    if (!signature) return;
    if (!groups.has(signature)) groups.set(signature, []);
    groups.get(signature).push(index);
  });
  return [...groups.entries()]
    .filter(([, indexes]) => indexes.length > 1)
    .map(([signature, indexes]) => ({
      signature,
      sectionIds: indexes.map((index) =>
        normalizeText(
          sections[index]?.sectionId || sections[index]?.title || `section_${index + 1}`,
        ),
      ),
    }));
}

function sectionBudgetViolations(sectionBudgetSummaries = []) {
  return (Array.isArray(sectionBudgetSummaries) ? sectionBudgetSummaries : [])
    .filter((summary) => {
      const retained = Number(summary.retainedSectionCount ?? 0);
      const effective = Number(
        summary.effectiveMaxSections ?? summary.maxSections ?? Infinity,
      );
      return Number.isFinite(effective) && retained > effective;
    })
    .map((summary) => ({
      tableIndex: summary.tableIndex,
      tableLabel: normalizeText(summary.tableLabel),
      retainedSectionCount: Number(summary.retainedSectionCount ?? 0),
      effectiveMaxSections: Number(
        summary.effectiveMaxSections ?? summary.maxSections ?? 0,
      ),
    }));
}

function analyzeOutputCompleteness({
  sections = [],
  expectedMetricIds = [],
  sectionBudgetSummaries = [],
} = {}) {
  const list = Array.isArray(sections) ? sections : [];
  const expected = uniqueMetricIds(expectedMetricIds);
  const expectedSet = new Set(expected);
  const rendered = renderedMetricIds(list);
  const renderedSet = new Set(rendered);
  const missingMetricIds = expected.filter((id) => !renderedSet.has(id));
  const duplicateMetricIds = duplicateMetricIdOwners(list);
  const emptyRequiredSectionIds = [];
  const incompleteRequiredSectionIds = [];
  const invalidNumbers = [];

  list.forEach((section, index) => {
    const ids = collectSectionMetricIds(section);
    const required = ids.some((id) => expectedSet.has(id));
    if (!required || DIAGNOSTIC_TYPE_PATTERN.test(sectionType(section))) return;
    const id = normalizeText(section.sectionId || section.title || `section_${index + 1}`);
    const rows = sectionRows(section);
    const meta = section.result?.meta || {};
    const enforceNonEmpty = Boolean(
      /^semantic_/i.test(sectionType(section)) ||
      meta.complete === true ||
      meta.addedByBusinessAugmentation === true ||
      meta.restoredByMandatorySummaryCoverageFloor === true ||
      meta.finalOutputQualityRequired === true
    );
    if (!rows.length && enforceNonEmpty) emptyRequiredSectionIds.push(id);
    if (!sectionComplete(section)) incompleteRequiredSectionIds.push(id);
    const entries = invalidNumberEntries(rows, id);
    invalidNumbers.push(...entries);
  });

  const duplicateSectionIds = duplicateIdentityValues(
    list,
    (section, index) => section.sectionId || `section_${index + 1}`,
  );
  const duplicateTitles = duplicateIdentityValues(
    list,
    (section, index) => section.title || `title_${index + 1}`,
  );
  const duplicateOutputs = duplicateSemanticOutputs(list);
  const budgetViolations = sectionBudgetViolations(sectionBudgetSummaries);

  const failureReasons = [];
  if (missingMetricIds.length) failureReasons.push("MISSING_EXPECTED_METRIC_IDS");
  if (duplicateMetricIds.length) failureReasons.push("DUPLICATE_METRIC_ID_OWNERS");
  if (emptyRequiredSectionIds.length) failureReasons.push("EMPTY_REQUIRED_SECTIONS");
  if (incompleteRequiredSectionIds.length) failureReasons.push("INCOMPLETE_REQUIRED_SECTIONS");
  if (duplicateSectionIds.length) failureReasons.push("DUPLICATE_SECTION_IDS");
  if (duplicateTitles.length) failureReasons.push("DUPLICATE_SECTION_TITLES");
  if (duplicateOutputs.length) failureReasons.push("DUPLICATE_SEMANTIC_OUTPUTS");
  if (invalidNumbers.length) failureReasons.push("INVALID_NUMERIC_VALUES");
  if (budgetViolations.length) failureReasons.push("SECTION_BUDGET_VIOLATIONS");

  return {
    version: OUTPUT_COMPLETENESS_CONTRACT_VERSION,
    pass: failureReasons.length === 0,
    status: failureReasons.length === 0 ? "PASS" : "FAIL",
    failureReasons,
    expectedMetricIds: expected,
    renderedMetricIds: rendered,
    expectedMetricCount: expected.length,
    renderedExpectedMetricCount: expected.filter((id) => renderedSet.has(id)).length,
    missingMetricIds,
    duplicateMetricIds,
    duplicateSectionIds,
    duplicateTitles,
    duplicateOutputs,
    emptyRequiredSectionIds,
    incompleteRequiredSectionIds,
    invalidNumbers,
    sectionBudgetViolations: budgetViolations,
    sectionCount: list.length,
  };
}

function finalQualityGateError(analysis = {}) {
  const error = new Error(
    `최종 출력 품질 게이트 실패: ${(analysis.failureReasons || []).join(", ") || "UNKNOWN"}`,
  );
  error.code = "FINAL_OUTPUT_QUALITY_GATE_FAILED";
  error.qualityGate = cloneValue(analysis);
  return error;
}

function applyFinalOutputQualityGate({
  sections = [],
  expectedMetricIds = [],
  sectionBudgetSummaries = [],
  throwOnFailure = false,
} = {}) {
  const inputSnapshot = JSON.stringify(sections);
  const duplicateResolution = resolveDuplicateSections({
    sections,
    expectedMetricIds,
  });
  const ownership = normalizeMetricIdOwnership({
    sections: duplicateResolution.sections,
    expectedMetricIds,
  });
  const identity = ensureUniqueSectionIdentity(ownership.sections);
  const normalizedSections = normalizeSectionMetricIds(identity.sections);
  const analysis = analyzeOutputCompleteness({
    sections: normalizedSections,
    expectedMetricIds,
    sectionBudgetSummaries,
  });

  if (JSON.stringify(sections) !== inputSnapshot) {
    throw new Error("최종 출력 품질 게이트가 입력 Section을 변경했습니다.");
  }

  const result = {
    version: FINAL_OUTPUT_QUALITY_GATE_VERSION,
    completenessVersion: OUTPUT_COMPLETENESS_CONTRACT_VERSION,
    duplicateResolverVersion: SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
    sections: normalizedSections,
    applied:
      duplicateResolution.applied ||
      ownership.applied ||
      identity.renamedSectionIds.length > 0 ||
      identity.renamedTitles.length > 0,
    pass: analysis.pass,
    status: analysis.status,
    failureReasons: analysis.failureReasons,
    analysis,
    duplicateResolution,
    metricOwnership: ownership,
    renamedSectionIds: identity.renamedSectionIds,
    renamedTitles: identity.renamedTitles,
  };

  if (!result.pass && throwOnFailure) {
    throw finalQualityGateError({
      ...analysis,
      finalOutputQualityGateVersion: FINAL_OUTPUT_QUALITY_GATE_VERSION,
    });
  }
  return result;
}

module.exports = {
  OUTPUT_COMPLETENESS_CONTRACT_VERSION,
  SEMANTIC_OUTPUT_DUPLICATE_RESOLVER_VERSION,
  FINAL_OUTPUT_QUALITY_GATE_VERSION,
  analyzeOutputCompleteness,
  applyFinalOutputQualityGate,
  duplicateSignature,
  ensureUniqueSectionIdentity,
  normalizeMetricIdOwnership,
  resolveDuplicateSections,
  sectionRows,
  sectionScope,
  sectionSemanticKey,
};
