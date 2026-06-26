const { buildNarrativeSections } = require("./reportNarrativeBuilder");
const { recommendChartSpec } = require("./chartRecommendationBuilder");
const {
  isBusinessTemplateResult,
  normalizeBusinessTemplateResult,
} = require("./businessTemplateContract");

function takeRows(rows = [], limit = 12) {
  return Array.isArray(rows) ? rows.slice(0, limit) : [];
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null || value === "") return null;
  const n = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

function formatValue(value) {
  const n = toNumber(value);
  if (n == null) return String(value ?? "");
  return Number.isInteger(n)
    ? n.toLocaleString()
    : Number(n.toFixed(2)).toLocaleString();
}

function inferLabelKey(row = {}) {
  const keys = Object.keys(row || {});

  return (
    keys.find((key) =>
      /name|이름|성명|title|제목|label|항목|부서|팀|구분|분류|연월|월|일자|date|id$/i.test(
        key,
      ),
    ) ||
    keys.find((key) => typeof row[key] === "string" && row[key].trim()) ||
    keys[0] ||
    "항목"
  );
}

function numericKeys(row = {}, exclude = []) {
  const excluded = new Set(exclude);
  return Object.keys(row || {}).filter(
    (key) => !excluded.has(key) && toNumber(row[key]) != null,
  );
}

function rowCountOfSection(section = {}) {
  const rows = Array.isArray(section.result?.rows) ? section.result.rows : [];
  return rows.length;
}

function totalRowsOfSections(sections = []) {
  return sections.reduce((sum, section) => sum + rowCountOfSection(section), 0);
}

function uniqueTruthy(values = []) {
  return [
    ...new Set(values.map((v) => String(v || "").trim()).filter(Boolean)),
  ];
}

function buildTopBottomInsight(rows = []) {
  const first = rows[0] || {};
  const labelKey = inferLabelKey(first);
  const valueKey = numericKeys(first, [labelKey, "rowCount"])[0];

  if (!valueKey) return null;

  const values = rows
    .map((row) => ({
      label: row[labelKey],
      value: toNumber(row[valueKey]),
    }))
    .filter((row) => row.label != null && row.value != null);

  if (!values.length) return null;

  const top = values.reduce((a, b) => (b.value > a.value ? b : a));
  return `${valueKey} 기준 최상위 항목은 ${top.label}(${formatValue(top.value)})입니다.`;
}

function buildSectionInsight(section = {}, chartSpec = null) {
  const title = section.title || section.sectionId || "분석 섹션";
  const rows = Array.isArray(section.result?.rows) ? section.result.rows : [];

  if (!rows.length)
    return `${title}는 결과 행이 없어 표 생성을 건너뛰었습니다.`;

  return (
    chartSpec?.insight ||
    buildTopBottomInsight(rows) ||
    `${title} 결과 ${rows.length.toLocaleString()}건이 생성되었습니다.`
  );
}

function sectionScore(section = {}) {
  const rows = Array.isArray(section.result?.rows) ? section.result.rows : [];
  const chartSpec = recommendChartSpec(section.result || {});

  let score = 0;
  if (rows.length) score += Math.min(rows.length, 50);
  if (chartSpec) score += 30;
  if (section.title) score += 5;
  if (/summary|요약|현황|추이|상위|하위/i.test(section.title || "")) score += 5;

  return score;
}

function sortReportSections(sections = []) {
  return [...sections].sort((a, b) => sectionScore(b) - sectionScore(a));
}

function buildExecutiveSummary({ title = "분석 보고서", sections = [] } = {}) {
  const chartableCount = sections.filter((section) =>
    recommendChartSpec(section.result || {}),
  ).length;
  const tableCount = sections.filter(
    (section) => rowCountOfSection(section) > 0,
  ).length;
  const totalRows = totalRowsOfSections(sections);

  return {
    title,
    sectionCount: sections.length,
    tableSectionCount: tableCount,
    chartSectionCount: chartableCount,
    totalRows,
    generatedAt: new Date().toISOString(),
  };
}

function buildBusinessSummaryBullets(sections = []) {
  const summary = buildExecutiveSummary({ sections });
  const topSections = sortReportSections(sections)
    .slice(0, 5)
    .map((section) => {
      const title = section.title || section.sectionId || "분석 섹션";
      const rows = rowCountOfSection(section);
      return `${title}: ${rows.toLocaleString()}건`;
    });

  return uniqueTruthy([
    `${summary.sectionCount.toLocaleString()}개 분석 섹션이 생성되었습니다.`,
    `${summary.tableSectionCount.toLocaleString()}개 표 섹션과 ${summary.chartSectionCount.toLocaleString()}개 차트 후보를 확인했습니다.`,
    `${summary.totalRows.toLocaleString()}건의 결과 행을 보고서 구조로 정리했습니다.`,
    ...topSections,
  ]).slice(0, 7);
}

function pushBusinessSections(reportSections = [], businessSections = []) {
  for (const section of sortReportSections(businessSections)) {
    const sectionResult = section.result || {};
    const rows = Array.isArray(sectionResult.rows) ? sectionResult.rows : [];
    const chartSpec = recommendChartSpec(sectionResult);
    const title = section.title || section.sectionId || "분석 결과";

    if (chartSpec && rows.length >= 2) {
      reportSections.push({
        type: "chart",
        title: chartSpec.title || title,
        chartSpec,
        rows: takeRows(rows, 50),
        insight: buildSectionInsight(section, chartSpec),
      });
    }

    if (rows.length) {
      reportSections.push({
        type: "table",
        title,
        rows: takeRows(rows, 12),
        rowCount: rows.length,
        note:
          rows.length > 12
            ? `상위 12건만 미리보기로 표시했습니다. 전체 ${rows.length.toLocaleString()}건`
            : "",
      });
    }
  }
}

function buildBusinessInsights(businessSections = []) {
  return sortReportSections(businessSections)
    .map((section) =>
      buildSectionInsight(section, recommendChartSpec(section.result || {})),
    )
    .filter(Boolean)
    .slice(0, 6);
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
    const title =
      normalizedBusinessResult.title || result.title || "업무 템플릿 보고서";
    const executiveSummary = buildExecutiveSummary({
      title,
      sections: businessSections,
    });

    const sections = [
      {
        type: "cover",
        title,
        subtitle: fileName || "",
        generatedAt: executiveSummary.generatedAt,
      },
      {
        type: "summary",
        title: "핵심 요약",
        summary: `${businessSections.length.toLocaleString()}개 분석 섹션을 보고서 형태로 정리했습니다.`,
        bullets: buildBusinessSummaryBullets(businessSections),
      },
    ];

    pushBusinessSections(sections, businessSections);

    sections.push({
      type: "insight",
      title: "분석 인사이트",
      bullets: buildBusinessInsights(businessSections),
    });

    return {
      version: "report_sections_v2",
      reportType: "analysisReport",
      title,
      generatedAt: executiveSummary.generatedAt,
      source: {
        fileName: fileName || "",
        message: message || "",
      },
      resultType: normalizedBusinessResult.resultType || "",
      operation: normalizedBusinessResult.templateId || "",
      executiveSummary,
      sections,
    };
  }

  const chartSpec = recommendChartSpec(result);
  const rows = Array.isArray(result?.rows) ? result.rows : [];
  const generatedAt = new Date().toISOString();

  const sections = [
    {
      type: "cover",
      title: narrative.title,
      subtitle: fileName || "",
      generatedAt,
    },
    {
      type: "summary",
      title: "핵심 요약",
      summary: narrative.summary,
      bullets: narrative.highlights || [],
    },
  ];

  if (chartSpec && rows.length >= 2) {
    sections.push({
      type: "chart",
      title: chartSpec.title || "차트",
      chartSpec,
      rows: takeRows(rows, 50),
      insight: chartSpec.insight || "",
    });
  }

  if (rows.length) {
    sections.push({
      type: "table",
      title: "분석 결과",
      rows: takeRows(rows, 12),
      rowCount: rows.length,
      note:
        rows.length > 12
          ? `상위 12건만 미리보기로 표시했습니다. 전체 ${rows.length.toLocaleString()}건`
          : "",
    });
  }

  sections.push({
    type: "insight",
    title: "분석 인사이트",
    bullets: narrative.highlights || [],
  });

  return {
    version: "report_sections_v2",
    reportType: "analysisReport",
    title: narrative.title,
    generatedAt,
    source: {
      fileName: fileName || "",
      message: message || "",
    },
    resultType: result?.resultType || "",
    operation: result?.operation || "",
    executiveSummary: {
      title: narrative.title,
      sectionCount: sections.length,
      tableSectionCount: rows.length ? 1 : 0,
      chartSectionCount: chartSpec ? 1 : 0,
      totalRows: rows.length,
      generatedAt,
    },
    sections,
  };
}

module.exports = {
  buildReportSections,
};
