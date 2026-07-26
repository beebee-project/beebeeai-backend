"use strict";

const METRIC_ID_CONTRACT_VERSION =
  "metric_id_contract_common_v1";

function normalizeMetricId(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .trim();
}

function metricIdValues(value) {
  if (Array.isArray(value)) return value;
  if (value == null || value === "") return [];
  return [value];
}

function uniqueMetricIds(values = []) {
  return Array.from(
    new Set(
      values
        .flatMap(metricIdValues)
        .map(normalizeMetricId)
        .filter(Boolean),
    ),
  );
}

function collectSectionMetricIds(section = {}) {
  const result = section.result || {};
  const meta = result.meta || {};

  return uniqueMetricIds([
    section.metricIds,
    section.metricId,
    result.metricIds,
    result.metricId,
    meta.metricIds,
    meta.metricId,
  ]);
}

function applySectionMetricIds(section = {}, explicitMetricIds = []) {
  const metricIds = uniqueMetricIds([
    collectSectionMetricIds(section),
    explicitMetricIds,
  ]);

  const result = section.result || {};
  const meta = result.meta || {};

  return {
    ...section,
    metricIds,
    result: {
      ...result,
      metricIds,
      meta: {
        ...meta,
        metricIds,
        metricIdContractVersion: METRIC_ID_CONTRACT_VERSION,
      },
    },
  };
}

function normalizeSectionMetricIds(sections = []) {
  return (Array.isArray(sections) ? sections : []).map((section) =>
    applySectionMetricIds(section),
  );
}

module.exports = {
  METRIC_ID_CONTRACT_VERSION,
  normalizeMetricId,
  uniqueMetricIds,
  collectSectionMetricIds,
  applySectionMetricIds,
  normalizeSectionMetricIds,
};
