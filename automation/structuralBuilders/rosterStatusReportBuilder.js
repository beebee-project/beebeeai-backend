const {
  findColumnHeader,
  makeTemplateCandidate,
  executeTemplateSections,
} = require("../businessTemplates/commonTemplateHelpers");

function uniqueCandidates(candidates = []) {
  const seen = new Set();

  return candidates.filter((candidate) => {
    const c = candidate.columns || {};
    const key = [
      candidate.recipeType || "",
      candidate.tableId || "",
      c.dimension || "",
      c.metric || "",
      candidate.sectionId || "",
    ].join("|");

    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildCategorySummaryCandidates({ table, config = {} }) {
  if (!table?.tableId) return [];

  const tableId = table.tableId;
  const metricHeader =
    config.metricHeader ||
    findColumnHeader(table, config.metricHints || [], { type: "number" });

  const candidates = [];

  for (const dim of config.dimensions || []) {
    const dimensionHeader = dim.header || findColumnHeader(table, dim.hints || []);
    if (!dimensionHeader) continue;

    if (metricHeader && dim.includeSummary !== false) {
      candidates.push(
        makeTemplateCandidate({
          sectionId: dim.summarySectionId || dim.sectionId,
          sectionType: dim.summarySectionType || dim.sectionType || "category_summary",
          recipeType: "group_summary",
          title: dim.summaryTitle || `${dimensionHeader}별 ${metricHeader} 요약`,
          tableId,
          columns: {
            dimension: dimensionHeader,
            metric: metricHeader,
          },
          chartHint: {
            preferredType: dim.chartType || "bar",
            categoryField: dimensionHeader,
            valueField: metricHeader,
          },
          narrativeHint: {
            focus: "category_summary",
            dimension: dimensionHeader,
            metric: metricHeader,
          },
        }),
      );
    }

    if (dim.includeCount) {
      candidates.push(
        makeTemplateCandidate({
          sectionId: dim.countSectionId || `${dim.sectionId || dimensionHeader}_count`,
          sectionType: dim.countSectionType || "category_count",
          recipeType: "category_count",
          title: dim.countTitle || `${dimensionHeader}별 건수`,
          tableId,
          columns: {
            dimension: dimensionHeader,
          },
          chartHint: {
            preferredType: "bar",
            categoryField: dimensionHeader,
            valueField: "count",
          },
          narrativeHint: {
            focus: "category_count",
            dimension: dimensionHeader,
          },
        }),
      );
    }
  }

  const topBottom = config.topBottom || null;
  if (topBottom && metricHeader) {
    const dimensionHeader =
      topBottom.dimensionHeader ||
      findColumnHeader(table, topBottom.dimensionHints || []) ||
      findColumnHeader(table, (config.dimensions || []).flatMap((d) => d.hints || []));

    if (dimensionHeader) {
      candidates.push(
        makeTemplateCandidate({
          sectionId: topBottom.sectionId || "top_bottom_category_metric",
          sectionType: topBottom.sectionType || "top_bottom_category_metric",
          recipeType: "top_bottom",
          title: topBottom.title || `${metricHeader} 상위/하위 항목`,
          tableId,
          columns: {
            dimension: dimensionHeader,
            metric: metricHeader,
          },
          chartHint: {
            preferredType: "bar",
            categoryField: dimensionHeader,
            valueField: metricHeader,
          },
          narrativeHint: {
            focus: "top_bottom",
            dimension: dimensionHeader,
            metric: metricHeader,
          },
        }),
      );
    }
  }

  return uniqueCandidates(candidates);
}

function buildCategorySummaryReportSections({
  normalizedQueryTables = [],
  table,
  templateCandidate = {},
  config = {},
}) {
  const candidates = buildCategorySummaryCandidates({ table, config });

  if (!candidates.length) return [];

  return executeTemplateSections({
    normalizedQueryTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });
}

module.exports = {
  buildCategorySummaryCandidates,
  buildCategorySummaryReportSections,
};
