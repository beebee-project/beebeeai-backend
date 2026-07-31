const METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION =
  "metric_relationship_priority_engine_v1";
const DERIVED_TOTAL_RELATION_VERSION =
  "derived_total_relation_v1_additive_product";
const REPRESENTATIVE_METRIC_PRIORITY_VERSION =
  "representative_metric_priority_v1";

const RELATION_ROLE = Object.freeze({
  PRIMARY_TOTAL: "primary_total",
  COMPONENT: "component",
  INDEPENDENT: "independent",
  UNIT_COMPONENT: "unit_component",
  PRIMARY_BASIS: "primary_basis",
});

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

function finiteNumber(value) {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function seriesLabel(series = {}) {
  return normalizeText(series.metricLabel || series.valueHeader || "");
}

function seriesKey(series = {}, index = 0) {
  return (
    normalizeText(series.key) ||
    `${normalizeKey(seriesLabel(series))}::${index}`
  );
}

function sameUnit(left = {}, right = {}) {
  const a = normalizeKey(left.unit || "");
  const b = normalizeKey(right.unit || "");
  return !a || !b || a === b;
}

function rowValueMap(series = {}) {
  const map = new Map();
  for (const record of Array.isArray(series.records) ? series.records : []) {
    const value = finiteNumber(record?.value);
    if (value == null) continue;
    const rowIndex = Number(record?.rowIndex);
    if (!Number.isFinite(rowIndex)) continue;
    map.set(rowIndex, value);
  }
  return map;
}

function combinations(
  items = [],
  size = 2,
  start = 0,
  prefix = [],
  output = [],
) {
  if (prefix.length === size) {
    output.push(prefix);
    return output;
  }
  for (let index = start; index < items.length; index += 1) {
    combinations(items, size, index + 1, [...prefix, items[index]], output);
  }
  return output;
}

function approximatelyEqual(left, right) {
  const scale = Math.max(1, Math.abs(left), Math.abs(right));
  return Math.abs(left - right) <= scale * 1e-7;
}

function relationMatchStats(target = {}, components = [], mode = "sum") {
  const targetMap = rowValueMap(target);
  const componentMaps = components.map(rowValueMap);
  const matches = [];
  let comparableCount = 0;
  let matchedCount = 0;

  for (const [rowIndex, targetValue] of targetMap.entries()) {
    const values = componentMaps.map((map) => map.get(rowIndex));
    if (values.some((value) => value == null)) continue;
    comparableCount += 1;
    const expected =
      mode === "product"
        ? values.reduce((product, value) => product * value, 1)
        : values.reduce((sum, value) => sum + value, 0);
    const matched = approximatelyEqual(targetValue, expected);
    if (matched) matchedCount += 1;
    matches.push({ rowIndex, targetValue, expected, matched });
  }

  const targetCount = targetMap.size;
  return {
    comparableCount,
    matchedCount,
    matchRatio: comparableCount ? matchedCount / comparableCount : 0,
    coverage: targetCount ? comparableCount / targetCount : 0,
    matches,
  };
}

function additiveEligible(series = {}) {
  return (
    series.operation === "sum" &&
    ["money_flow", "quantity_flow", "stock_snapshot"].includes(
      normalizeText(series.metricRole),
    )
  );
}

function additiveTargetEvidenceScore(series = {}) {
  const label = seriesLabel(series);
  let score = 0;
  if (/^(?:총|전체)|(?:합계|총계)$|\btotal\b/i.test(label)) score += 180;
  if (
    /금액|비용|사용량|수량|매출|지출|예산|원가|amount|cost|usage|quantity/i.test(
      label,
    )
  ) {
    score += 35;
  }
  if (/단가|평균|비율|율|기간|일수|rate|average|duration/i.test(label)) {
    score -= 100;
  }
  return score;
}

function detectAdditiveRelations(seriesList = []) {
  const eligible = seriesList.filter(additiveEligible);
  const relations = [];

  for (const target of eligible) {
    const candidates = eligible.filter(
      (component) => component !== target && sameUnit(target, component),
    );
    if (candidates.length < 2) continue;

    let best = null;
    const maxSize = Math.min(4, candidates.length);
    for (let size = 2; size <= maxSize; size += 1) {
      for (const componentSet of combinations(candidates, size)) {
        const stats = relationMatchStats(target, componentSet, "sum");
        if (
          stats.comparableCount < 3 ||
          stats.coverage < 0.75 ||
          stats.matchRatio < 0.95
        ) {
          continue;
        }

        const positiveContribution = componentSet.every((component) =>
          [...rowValueMap(component).values()].some(
            (value) => Math.abs(value) > 0,
          ),
        );
        if (!positiveContribution) continue;

        const score =
          stats.matchRatio * 1000 +
          stats.coverage * 200 +
          additiveTargetEvidenceScore(target) -
          size * 2;
        if (!best || score > best.score) {
          best = { target, components: componentSet, stats, score };
        }
      }
    }

    if (best) {
      relations.push({
        relationType: "additive_total",
        targetKey: seriesKey(best.target),
        targetMetric: seriesLabel(best.target),
        componentKeys: best.components.map((item) => seriesKey(item)),
        componentMetrics: best.components.map(seriesLabel),
        matchRatio: best.stats.matchRatio,
        coverage: best.stats.coverage,
        comparableRowCount: best.stats.comparableCount,
        version: DERIVED_TOTAL_RELATION_VERSION,
      });
    }
  }

  return relations;
}

function detectProductRelations(seriesList = []) {
  const targets = seriesList.filter((series) =>
    ["money_flow", "stock_snapshot"].includes(normalizeText(series.metricRole)),
  );
  const quantities = seriesList.filter((series) =>
    ["quantity_flow", "stock_snapshot"].includes(
      normalizeText(series.metricRole),
    ),
  );
  const rates = seriesList.filter(
    (series) => normalizeText(series.metricRole) === "unit_rate",
  );
  const relations = [];

  for (const target of targets) {
    let best = null;
    for (const quantity of quantities) {
      for (const rate of rates) {
        const stats = relationMatchStats(target, [quantity, rate], "product");
        if (
          stats.comparableCount < 3 ||
          stats.coverage < 0.75 ||
          stats.matchRatio < 0.95
        ) {
          continue;
        }
        const score = stats.matchRatio * 1000 + stats.coverage * 200;
        if (!best || score > best.score) {
          best = { quantity, rate, stats, score };
        }
      }
    }
    if (best) {
      relations.push({
        relationType: "multiplicative_total",
        targetKey: seriesKey(target),
        targetMetric: seriesLabel(target),
        componentKeys: [seriesKey(best.quantity), seriesKey(best.rate)],
        componentMetrics: [seriesLabel(best.quantity), seriesLabel(best.rate)],
        quantityMetric: seriesLabel(best.quantity),
        rateMetric: seriesLabel(best.rate),
        matchRatio: best.stats.matchRatio,
        coverage: best.stats.coverage,
        comparableRowCount: best.stats.comparableCount,
        version: DERIVED_TOTAL_RELATION_VERSION,
      });
    }
  }

  return relations;
}

function chooseNonConflictingRelations(relations = []) {
  const ordered = [...relations].sort(
    (left, right) =>
      right.matchRatio - left.matchRatio ||
      right.coverage - left.coverage ||
      (right.componentKeys?.length || 0) - (left.componentKeys?.length || 0),
  );
  const targetSeen = new Set();
  const output = [];
  for (const relation of ordered) {
    if (targetSeen.has(relation.targetKey)) continue;
    targetSeen.add(relation.targetKey);
    output.push(relation);
  }
  return output;
}

function baseRepresentativePriority(series = {}) {
  const role = normalizeText(series.metricRole);
  const label = seriesLabel(series);
  let score = 400;
  if (role === "stock_snapshot") score += 220;
  if (role === "money_flow") score += 160;
  if (role === "quantity_flow") score += 130;
  if (role === "count") score += 90;
  if (role === "duration") score += 40;
  if (role === "unit_rate" || role === "percentage_rate") score += 20;
  if (/^(?:총|전체)|(?:합계|총계)$|\btotal\b/i.test(label)) score += 100;
  return score;
}

function applyMetricRelationshipPriorities(seriesList = []) {
  const originalSnapshot = JSON.stringify(seriesList);
  const cloned = (Array.isArray(seriesList) ? seriesList : []).map(
    (series, index) => ({
      ...cloneValue(series),
      relationshipRole: RELATION_ROLE.INDEPENDENT,
      representativeMetricPriority: baseRepresentativePriority(series),
      metricRelationships: [],
      originalSeriesIndex: index,
    }),
  );

  const keyMap = new Map(
    cloned.map((series, index) => [seriesKey(series, index), series]),
  );
  const detected = chooseNonConflictingRelations([
    ...detectAdditiveRelations(cloned),
    ...detectProductRelations(cloned),
  ]);

  for (const relation of detected) {
    const target = keyMap.get(relation.targetKey);
    if (!target) continue;
    target.relationshipRole = RELATION_ROLE.PRIMARY_TOTAL;
    target.representativeMetricPriority = Math.max(
      target.representativeMetricPriority,
      relation.relationType === "additive_total" ? 1200 : 1100,
    );
    target.metricRelationships.push(cloneValue(relation));

    for (const componentKey of relation.componentKeys || []) {
      const component = keyMap.get(componentKey);
      if (!component || component === target) continue;
      const unitComponent =
        relation.relationType === "multiplicative_total" &&
        normalizeText(component.metricRole) === "unit_rate";
      const primaryBasis =
        relation.relationType === "multiplicative_total" &&
        normalizeText(component.metricRole) === "stock_snapshot";
      component.relationshipRole = unitComponent
        ? RELATION_ROLE.UNIT_COMPONENT
        : primaryBasis
          ? RELATION_ROLE.PRIMARY_BASIS
          : RELATION_ROLE.COMPONENT;
      component.componentOfMetric = relation.targetMetric;
      component.componentOfKey = relation.targetKey;
      component.representativeMetricPriority = primaryBasis
        ? Math.max(component.representativeMetricPriority, 980)
        : Math.min(
            component.representativeMetricPriority,
            unitComponent ? 250 : 320,
          );
      component.metricRelationships.push(cloneValue(relation));
    }
  }

  cloned.sort(
    (left, right) =>
      right.representativeMetricPriority - left.representativeMetricPriority ||
      left.originalSeriesIndex - right.originalSeriesIndex,
  );

  if (JSON.stringify(seriesList) !== originalSnapshot) {
    throw new Error("Metric relationship engine mutated input series.");
  }

  return {
    version: METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
    relationVersion: DERIVED_TOTAL_RELATION_VERSION,
    priorityVersion: REPRESENTATIVE_METRIC_PRIORITY_VERSION,
    series: cloned,
    relations: detected,
    primaryMetricLabels: detected.map((relation) => relation.targetMetric),
    componentMetricLabels: Array.from(
      new Set(detected.flatMap((relation) => relation.componentMetrics || [])),
    ),
  };
}

function sectionMetricHeader(section = {}) {
  return normalizeText(
    section.result?.metric?.header ||
      section.metricHeader ||
      section.result?.meta?.metricHeader ||
      "",
  );
}

function prioritizeBusinessSections({
  sections = [],
  primaryMetricLabels = [],
  componentMetricLabels = [],
} = {}) {
  const primary = new Set(
    primaryMetricLabels.map(normalizeKey).filter(Boolean),
  );
  const components = new Set(
    componentMetricLabels.map(normalizeKey).filter(Boolean),
  );
  const decorated = (Array.isArray(sections) ? sections : []).map(
    (section, index) => {
      const metric = normalizeKey(sectionMetricHeader(section));
      let score = 500;
      if (primary.has(metric)) score = 1200;
      else if (components.has(metric)) score = 300;
      if (/계약|핵심지표|contract/i.test(normalizeText(section.title)))
        score += 100;
      if (String(section.sectionType || "").includes("overview")) score += 80;
      return { section, index, score, metric };
    },
  );

  const sorted = [...decorated].sort(
    (left, right) => right.score - left.score || left.index - right.index,
  );
  const reorderedCount = sorted.reduce(
    (count, item, index) => count + (item.index === index ? 0 : 1),
    0,
  );
  return {
    sections: sorted.map((item) => item.section),
    applied: reorderedCount > 0 && primary.size > 0,
    reorderedSectionCount: reorderedCount,
    primaryMetricLabels: [...primaryMetricLabels],
    componentMetricLabels: [...componentMetricLabels],
  };
}

module.exports = {
  METRIC_RELATIONSHIP_PRIORITY_ENGINE_VERSION,
  DERIVED_TOTAL_RELATION_VERSION,
  REPRESENTATIVE_METRIC_PRIORITY_VERSION,
  RELATION_ROLE,
  applyMetricRelationshipPriorities,
  detectAdditiveRelations,
  detectProductRelations,
  prioritizeBusinessSections,
  relationMatchStats,
};
