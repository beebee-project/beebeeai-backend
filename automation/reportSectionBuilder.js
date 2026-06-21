const { buildNarrativeSections } = require("./reportNarrativeBuilder");
const { recommendChartSpec } = require("./chartRecommendationBuilder");
const {
  isBusinessTemplateResult,
  normalizeBusinessTemplateResult,
} = require("./businessTemplateContract");

function takeRows(rows = [], limit = 12) {
  return Array.isArray(rows) ? rows.slice(0, limit) : [];
}

function buildReportSections({ fileName, message, result } = {}) {
  const narrative = buildNarrativeSections(result, {
    message,
    fileName,
  });

  const normalizedBusinessResult = isBusinessTemplateResult(result)
    ? normalizeBusinessTemplateResult(result)
    : null;

  const businessSections = Array.isArray(normalizedBusinessResult?.sections)
    ? normalizedBusinessResult.sections
    : [];

  if (businessSections.length) {
    const sections = [
      {
        type: "cover",
        title: result.title || "업무 템플릿 보고서",
        subtitle: fileName || "",
      },
      {
        type: "summary",
        title: "핵심 요약",
        summary: `${businessSections.length}개 분석 섹션이 생성되었습니다.`,
        bullets: businessSections.map(
          (s) => s.title || s.sectionId || "분석 섹션",
        ),
      },
    ];

    for (const section of businessSections) {
      const sectionResult = section.result || {};
      const rows = Array.isArray(sectionResult.rows) ? sectionResult.rows : [];
      const chartSpec = recommendChartSpec(sectionResult);

      if (chartSpec) {
        sections.push({
          type: "chart",
          title: section.title || chartSpec.title || "차트",
          chartSpec,
          rows: takeRows(rows, 50),
        });
      }

      if (rows.length) {
        sections.push({
          type: "table",
          title: section.title || "분석 결과",
          rows: takeRows(rows, 12),
          rowCount: rows.length,
        });
      }
    }

    sections.push({
      type: "insight",
      title: "분석 인사이트",
      bullets: businessSections.map(
        (s) => `${s.title || s.sectionId || "섹션"} 결과가 생성되었습니다.`,
      ),
    });

    return {
      version: "report_sections_v1",
      reportType: "analysisReport",
      title: normalizedBusinessResult.title || "업무 템플릿 보고서",
      source: {
        fileName: fileName || "",
        message: message || "",
      },
      resultType: normalizedBusinessResult.resultType || "",
      operation: normalizedBusinessResult.templateId || "",
      sections,
    };
  }

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
    reportType: "analysisReport",
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
