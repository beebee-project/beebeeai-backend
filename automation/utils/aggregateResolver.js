const { FORMULA_HEURISTICS } = require("../config/formulaHeuristicsConfig");

function collectSectionText(operation = "", section = {}) {
  return [
    operation,
    section.sectionId,
    section.sectionType,
    section.title,
    section.result?.operation,
    section.result?.metric?.aggregation,
    section.result?.metric?.aggregate,
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();
}

function includesAnyKeyword(text = "", keywords = []) {
  const normalizedText = String(text || "").toLowerCase();

  return keywords.some((keyword) =>
    normalizedText.includes(String(keyword || "").toLowerCase()),
  );
}

function isRankingLikeSection(section = {}) {
  const text = [
    section.sectionId,
    section.sectionType,
    section.title,
    section.result?.operation,
  ]
    .filter(Boolean)
    .join(" ");

  return FORMULA_HEURISTICS.rankingSkipPattern.test(text);
}

function normalizeAggregateOperation(operation = "", section = {}) {
  const text = collectSectionText(operation, section);
  const keywords = FORMULA_HEURISTICS.aggregateOperationKeywords;

  if (includesAnyKeyword(text, keywords.average)) return "average";
  if (includesAnyKeyword(text, keywords.count)) return "count";
  if (includesAnyKeyword(text, keywords.sum)) return "sum";

  return "sum";
}

function isSimpleAggregateOperation(operation = "sum") {
  return FORMULA_HEURISTICS.simpleAggregateOperations.includes(
    String(operation || "").toLowerCase(),
  );
}

function findHeaderIndex(headers = [], candidates = []) {
  const normalizedHeaders = headers.map((h) => String(h || "").trim());
  const normalizedCandidates = candidates
    .filter(Boolean)
    .map((h) => String(h || "").trim());

  for (const candidate of normalizedCandidates) {
    const exact = normalizedHeaders.findIndex((h) => h === candidate);
    if (exact >= 0) return exact;
  }

  for (const candidate of normalizedCandidates) {
    const partial = normalizedHeaders.findIndex(
      (h) => h && candidate && (h.includes(candidate) || candidate.includes(h)),
    );
    if (partial >= 0) return partial;
  }

  return -1;
}

function findExactAggregateHeaderIndex(headers = [], candidates = []) {
  const normalizedHeaders = headers.map((h) => String(h || "").trim());
  const normalizedCandidates = candidates
    .filter(Boolean)
    .map((h) => String(h || "").trim());

  for (const candidate of normalizedCandidates) {
    const index = normalizedHeaders.findIndex((header) => header === candidate);
    if (index >= 0) return index;
  }

  return -1;
}

function resolveAggregateFormulaTargets({
  headers = [],
  operation = "sum",
  formulaPlan = {},
  criteriaColIndex = -1,
} = {}) {
  const targets = [];
  const used = new Set([criteriaColIndex]);

  function pushTarget(op, index) {
    if (index < 0 || used.has(index)) return;
    targets.push({ operation: op, columnIndex: index });
    used.add(index);
  }

  const targetHeaders = FORMULA_HEURISTICS.aggregateTargetHeaders;

  const countIndex = findExactAggregateHeaderIndex(
    headers,
    targetHeaders.count,
  );
  const sumIndex = findExactAggregateHeaderIndex(headers, targetHeaders.sum);
  const averageIndex = findExactAggregateHeaderIndex(
    headers,
    targetHeaders.average,
  );

  const hasMultiAggregateColumns =
    countIndex >= 0 || sumIndex >= 0 || averageIndex >= 0;

  if (hasMultiAggregateColumns) {
    pushTarget("count", countIndex);

    if (formulaPlan.metric?.letter) {
      pushTarget("sum", sumIndex);
      pushTarget("average", averageIndex);
    }

    return targets;
  }

  const normalizedOperation = normalizeAggregateOperation(
    operation,
    formulaPlan,
  );

  const singleValueIndex = findExactAggregateHeaderIndex(headers, [
    ...targetHeaders.value,
    normalizedOperation,
    formulaPlan.metric?.header,
  ]);

  if (singleValueIndex < 0) {
    return [];
  }

  if (normalizedOperation !== "count" && !formulaPlan.metric?.letter) {
    return [];
  }

  pushTarget(normalizedOperation, singleValueIndex);

  return targets;
}

function resolveCriteriaColumnIndex({ headers = [], formulaPlan = {} } = {}) {
  return findHeaderIndex(headers, [
    formulaPlan.group?.header,
    ...FORMULA_HEURISTICS.criteriaHeaderFallbacks,
  ]);
}

module.exports = {
  isRankingLikeSection,
  normalizeAggregateOperation,
  isSimpleAggregateOperation,
  findHeaderIndex,
  findExactAggregateHeaderIndex,
  resolveAggregateFormulaTargets,
  resolveCriteriaColumnIndex,
};
