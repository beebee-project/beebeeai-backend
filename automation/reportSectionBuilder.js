const { buildNarrativeSections } = require("./reportNarrativeBuilder");
const { recommendChartSpec } = require("./chartRecommendationBuilder");

function takeRows(rows = [], limit = 12) {
  return Array.isArray(rows) ? rows.slice(0, limit) : [];
}

function buildReportSections({ fileName, message, result } = {}) {
  const narrative = buildNarrativeSections(result, {
    message,
    fileName,
  });

  const chartSpec = recommendChartSpec(result);
  const rows = Array.isArray(result?.rows) ? result.rows : [];

  const sections = [
    {
      type: "cover",
      title: narrative.title,
      subtitle: fileName || "",
    },
    {
      type: "summary",
      title: "핵심 요약",
      summary: narrative.summary,
      bullets: narrative.highlights || [],
    },
  ];

  if (chartSpec) {
    sections.push({
      type: "chart",
      title: chartSpec.title || "차트",
      chartSpec,
      rows: takeRows(rows, 50),
    });
  }

  if (rows.length) {
    sections.push({
      type: "table",
      title: "분석 결과",
      rows: takeRows(rows, 12),
      rowCount: rows.length,
    });
  }

  sections.push({
    type: "insight",
    title: "분석 인사이트",
    bullets: narrative.highlights || [],
  });

  return {
    version: "report_sections_v1",
    title: narrative.title,
    source: {
      fileName: fileName || "",
      message: message || "",
    },
    resultType: result?.resultType || "",
    operation: result?.operation || "",
    sections,
  };
}

module.exports = {
  buildReportSections,
};
