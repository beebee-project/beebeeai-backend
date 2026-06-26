const { buildNarrativeSections } = require("./reportNarrativeBuilder");
const { recommendChartSpec } = require("./chartRecommendationBuilder");
const {
  decorateReportSections,
  tableColumnLabels,
} = require("./reportDisplayUtils");
const {
  isBusinessTemplateResult,
  normalizeBusinessTemplateResult,
} = require("./businessTemplateContract");

const PREVIEW_ROW_LIMIT = 12;
const CHART_ROW_LIMIT = 50;

function takeRows(rows = [], limit = PREVIEW_ROW_LIMIT) {
  return Array.isArray(rows) ? rows.slice(0, limit) : [];
}

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function compactText(value = "") {
  return String(value || "").trim();
}

function isNumericValue(value) {
  if (value === null || value === undefined || value === "") return false;
  const n = Number(value);
  return Number.isFinite(n);
}

function safeRows(value = {}) {
  if (Array.isArray(value?.rows)) return value.rows;
  if (Array.isArray(value?.data)) return value.data;
  if (Array.isArray(value?.items)) return value.items;
  if (Array.isArray(value?.result?.rows)) return value.result.rows;
  if (Array.isArray(value?.result?.data)) return value.result.data;
  if (Array.isArray(value?.result?.items)) return value.result.items;
  return [];
}

function getRowKeys(rows = []) {
  const keys = [];
  rows.forEach((row) => {
    if (!isPlainObject(row)) return;
    Object.keys(row).forEach((key) => {
      if (!keys.includes(key)) keys.push(key);
    });
  });
  return keys;
}

function isMostlyNumeric(rows = [], key = "") {
  const values = rows
    .map((row) => row?.[key])
    .filter((value) => value !== null && value !== undefined && value !== "");
  if (!values.length) return false;
  const numericCount = values.filter(isNumericValue).length;
  return numericCount / values.length >= 0.7;
}

function getSectionTitle(section = {}, fallback = "분석 결과") {
  return (
    compactText(
      section.title || section.name || section.label || section.sectionTitle,
    ) || fallback
  );
}

function getSectionResult(section = {}) {
  return isPlainObject(section.result) ? section.result : section;
}

function getSectionRowCount(section = {}, rows = []) {
  const result = getSectionResult(section);
  const explicit = Number(result.rowCount ?? section.rowCount);
  if (Number.isFinite(explicit) && explicit >= 0) return explicit;
  return rows.length;
}

function normalizeReportSourceSection(section = {}) {
  const result = getSectionResult(section);
  const rows = safeRows(result).length ? safeRows(result) : safeRows(section);
  const title = getSectionTitle(section, getSectionTitle(result));
  const rowCount = getSectionRowCount(section, rows);

  return {
    ...result,
    ...section,
    result,
    title,
    rows,
    rowCount,
    columnLabels: {
      ...(isPlainObject(result.columnLabels) ? result.columnLabels : {}),
      ...(isPlainObject(section.columnLabels) ? section.columnLabels : {}),
    },
  };
}

function getBusinessSections(result = {}, normalizedBusinessResult = null) {
  if (Array.isArray(normalizedBusinessResult?.sections)) {
    return normalizedBusinessResult.sections
      .map(normalizeReportSourceSection)
      .filter((section) => section.title || safeRows(section).length);
  }

  const candidates = [
    result.businessSections,
    result.sections,
    result.summarySections,
    result.reportSections,
    result.outputs,
  ];

  for (const value of candidates) {
    if (!Array.isArray(value)) continue;
    const sections = value
      .filter((section) => isPlainObject(section))
      .filter(
        (section) =>
          !["cover", "summary", "insight", "chart"].includes(section.type),
      )
      .map(normalizeReportSourceSection)
      .filter(
        (section) =>
          safeRows(section).length || section.rowCount || section.title,
      );

    if (sections.length) return sections;
  }

  const rows = safeRows(result);
  if (rows.length) {
    return [
      normalizeReportSourceSection({
        title: result.title || "분석 결과",
        ...result,
        rows,
        rowCount: rows.length,
      }),
    ];
  }

  return [];
}

function pickCategoryField(rows = []) {
  const keys = getRowKeys(rows);

  if (keys.includes("label")) return "label";

  const preferred = keys.find(
    (key) =>
      /부서|직급|입사연월|연월|월|일자|날짜|기간|항목|구분|카테고리|category|name|title/i.test(
        key,
      ) && !isMostlyNumeric(rows, key),
  );
  if (preferred) return preferred;

  return keys.find((key) => !isMostlyNumeric(rows, key)) || "";
}

function pickValueField(rows = [], section = {}) {
  const keys = getRowKeys(rows).filter((key) => isMostlyNumeric(rows, key));
  const title = getSectionTitle(section);

  if (/연봉|급여|임금|salary|pay|wage/i.test(title)) {
    if (keys.includes("average")) return "average";
    if (keys.includes("value")) return "value";
    if (keys.includes("sum")) return "sum";
  }

  const preferredOrder = ["value", "count", "average", "avg", "sum"];
  for (const key of preferredOrder) {
    if (keys.includes(key)) return key;
  }

  return keys[0] || "";
}

function formatDisplayValue(value, field = "") {
  if (!isNumericValue(value)) return String(value ?? "");
  const n = Number(value);
  const options = /(average|avg|mean)/i.test(field)
    ? { maximumFractionDigits: 1 }
    : { maximumFractionDigits: Number.isInteger(n) ? 0 : 1 };
  return n.toLocaleString("ko-KR", options);
}

function displayLabelForField(field = "", section = {}) {
  const labels = tableColumnLabels(safeRows(section), {
    title: section.title,
    sectionTitle: section.title,
    fieldLabels: section.columnLabels,
  });
  return labels[field] || String(field || "값");
}

function buildChartTitle(section = {}, categoryField = "", valueField = "") {
  const title = getSectionTitle(section);
  const categoryLabel = displayLabelForField(categoryField, section);
  const valueLabel = displayLabelForField(valueField, section);

  if (/상위|하위/.test(title)) return title;
  if (valueField === "average" || valueField === "avg") {
    return `${categoryLabel}별 ${valueLabel}`;
  }
  if (/입사/.test(title) && valueField === "count") {
    return `${categoryLabel}별 입사 건수`;
  }
  if (/인원/.test(title) && valueField === "count") return title;
  return title || `${categoryLabel}별 ${valueLabel}`;
}

function buildChartInsight(
  rows = [],
  section = {},
  categoryField = "",
  valueField = "",
) {
  if (!rows.length || !categoryField || !valueField) return "";

  const top = rows
    .filter((row) => isNumericValue(row?.[valueField]))
    .slice()
    .sort((a, b) => Number(b[valueField]) - Number(a[valueField]))[0];

  if (!top) return "";

  const categoryLabel = displayLabelForField(categoryField, section);
  const valueLabel = displayLabelForField(valueField, section);
  const categoryValue = top[categoryField];
  const value = formatDisplayValue(top[valueField], valueField);

  return `${categoryLabel} 기준 ${valueLabel} 최상위 항목은 ${categoryValue}(${value})입니다.`;
}

function isUsableChartSpec(spec = null, rows = []) {
  if (!spec || !rows.length) return false;
  const keys = getRowKeys(rows);
  if (!spec.categoryField || !keys.includes(spec.categoryField)) return false;
  const seriesFields = Array.isArray(spec.seriesFields)
    ? spec.seriesFields
    : spec.valueField
      ? [spec.valueField]
      : [];
  return seriesFields.some((field) => keys.includes(field));
}

function isChartableSection(section = {}) {
  const rows = safeRows(section);
  if (rows.length < 2) return false;

  const title = getSectionTitle(section);
  if (
    /전체\s*대상\s*수|총\s*건수|전체\s*건수/.test(title) &&
    rows.length <= 1
  ) {
    return false;
  }

  const categoryField = pickCategoryField(rows);
  const valueField = pickValueField(rows, section);
  return Boolean(categoryField && valueField && categoryField !== valueField);
}

function inferChartSpec(section = {}) {
  const rows = safeRows(section);
  if (!isChartableSection(section)) return null;

  const categoryField = pickCategoryField(rows);
  const valueField = pickValueField(rows, section);
  const fieldLabels = tableColumnLabels(rows, {
    title: section.title,
    sectionTitle: section.title,
    fieldLabels: section.columnLabels,
  });
  const recommendedType = /연월|월|일자|날짜|기간|date|month/i.test(
    categoryField,
  )
    ? "line"
    : rows.length > 6
      ? "horizontal_bar"
      : "bar";

  return {
    version: "chart_spec_v1",
    recommendedType,
    title: buildChartTitle(section, categoryField, valueField),
    categoryField,
    valueField,
    seriesFields: [valueField],
    rowCount: section.rowCount || rows.length,
    fieldLabels,
    insight: buildChartInsight(rows, section, categoryField, valueField),
    generatedBy: "reportSectionBuilder:fallback",
  };
}

function safeRecommendChartSpec(section = {}) {
  const rows = safeRows(section);
  try {
    const spec = recommendChartSpec({
      ...getSectionResult(section),
      ...section,
      rows,
    });
    if (isUsableChartSpec(spec, rows)) return spec;
  } catch (error) {
    // 차트 추천 실패는 보고서 생성 실패로 전파하지 않고 공통 추론 폴백을 사용한다.
  }
  return null;
}

function buildChartSection(section = {}) {
  const rows = safeRows(section);
  const recommended = safeRecommendChartSpec(section);
  const inferred = inferChartSpec(section);
  const baseSpec = recommended || inferred;

  if (!baseSpec) return null;

  const fieldLabels = {
    ...tableColumnLabels(rows, {
      title: section.title,
      sectionTitle: section.title,
      fieldLabels: section.columnLabels,
    }),
    ...(baseSpec.fieldLabels || {}),
  };

  const chartSpec = {
    ...baseSpec,
    fieldLabels,
    title: baseSpec.title || inferred?.title || getSectionTitle(section),
    insight: baseSpec.insight || inferred?.insight || "",
  };

  return {
    type: "chart",
    title: chartSpec.title || getSectionTitle(section),
    chartSpec,
    rows: takeRows(rows, CHART_ROW_LIMIT),
    rowCount: section.rowCount || rows.length,
    columnLabels: fieldLabels,
    insight: chartSpec.insight || "",
  };
}

function buildTableSection(section = {}) {
  const rows = safeRows(section);
  const rowCount = getSectionRowCount(section, rows);
  const previewRows = takeRows(rows, PREVIEW_ROW_LIMIT);
  const note =
    section.note ||
    (rowCount > PREVIEW_ROW_LIMIT
      ? `상위 ${PREVIEW_ROW_LIMIT}건만 미리보기로 표시했습니다. 전체 ${rowCount}건`
      : "");

  return {
    type: "table",
    title: getSectionTitle(section),
    rows: previewRows,
    rowCount,
    columnLabels: tableColumnLabels(rows, {
      title: section.title,
      sectionTitle: section.title,
      fieldLabels: section.columnLabels,
    }),
    note,
  };
}

function countSectionsByType(sections = [], type = "") {
  return sections.filter((section) => section.type === type).length;
}

function buildExecutiveSummary({ title, sections, sectionCount, totalRows }) {
  return {
    title,
    sectionCount: Number(
      sectionCount || countSectionsByType(sections, "table"),
    ),
    tableSectionCount: countSectionsByType(sections, "table"),
    chartSectionCount: countSectionsByType(sections, "chart"),
    totalRows: Number(totalRows || 0),
    generatedAt: new Date().toISOString(),
  };
}

function buildSummaryBullets({
  businessSections = [],
  sections = [],
  totalRows = 0,
}) {
  const tableSections = sections.filter((section) => section.type === "table");
  const chartSections = sections.filter((section) => section.type === "chart");
  const bullets = [
    `${businessSections.length || tableSections.length}개 분석 섹션이 생성되었습니다.`,
    `${tableSections.length}개 표 섹션과 ${chartSections.length}개 차트 후보를 확인했습니다.`,
  ];

  if (totalRows)
    bullets.push(`${totalRows}건의 결과 행을 보고서 구조로 정리했습니다.`);

  for (const section of tableSections.slice(0, 4)) {
    bullets.push(
      `${section.title || "분석 결과"}: ${section.rowCount || section.rows?.length || 0}건`,
    );
  }

  return bullets.slice(0, 8);
}

function resolveBusinessTotalRows(
  result = {},
  businessSections = [],
  normalizedBusinessResult = null,
) {
  const explicit = Number(
    normalizedBusinessResult?.rowCount ?? result?.rowCount ?? result?.totalRows,
  );
  if (Number.isFinite(explicit) && explicit >= 0) return explicit;
  return businessSections.reduce(
    (sum, section) =>
      sum + Number(section.rowCount || safeRows(section).length || 0),
    0,
  );
}

function buildBusinessReportSections({
  fileName,
  message,
  result,
  normalizedBusinessResult,
}) {
  const businessSections = getBusinessSections(
    result,
    normalizedBusinessResult,
  );
  const title =
    normalizedBusinessResult?.title || result?.title || "업무 템플릿 보고서";
  const generatedAt = new Date().toISOString();

  const bodySections = [];
  for (const section of businessSections) {
    const chartSection = buildChartSection(section);
    const tableSection = buildTableSection(section);

    if (chartSection) bodySections.push(chartSection);
    if (tableSection) bodySections.push(tableSection);
  }

  const totalRows = resolveBusinessTotalRows(
    result,
    businessSections,
    normalizedBusinessResult,
  );
  const decoratedBodySections = decorateReportSections(bodySections);

  const sections = decorateReportSections([
    {
      type: "cover",
      title,
      subtitle: fileName || "",
      generatedAt,
    },
    {
      type: "summary",
      title: "핵심 요약",
      summary: `${businessSections.length}개 분석 섹션을 보고서 형태로 정리했습니다.`,
      bullets: buildSummaryBullets({
        businessSections,
        sections: decoratedBodySections,
        totalRows,
      }),
    },
    ...decoratedBodySections,
    {
      type: "insight",
      title: "분석 인사이트",
      bullets: decoratedBodySections
        .filter((section) => section.type === "chart" && section.insight)
        .map((section) => section.insight)
        .concat(
          businessSections.map(
            (section) =>
              `${section.title || section.sectionId || "섹션"} 결과가 생성되었습니다.`,
          ),
        )
        .slice(0, 8),
    },
  ]);

  return {
    version: "report_sections_v2",
    reportType: "analysisReport",
    title,
    generatedAt,
    source: {
      fileName: fileName || "",
      message: message || "",
    },
    resultType:
      normalizedBusinessResult?.resultType || result?.resultType || "",
    operation: normalizedBusinessResult?.templateId || result?.operation || "",
    executiveSummary: buildExecutiveSummary({
      title,
      sections,
      sectionCount: businessSections.length,
      totalRows,
    }),
    sections,
  };
}

function buildGenericReportSections({ fileName, message, result } = {}) {
  const narrative = buildNarrativeSections(result, {
    message,
    fileName,
  });
  const chartSpec = recommendChartSpec(result);
  const rows = safeRows(result);
  const title = narrative.title || result?.title || "분석 보고서";
  const generatedAt = new Date().toISOString();

  const sections = [
    {
      type: "cover",
      title,
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

  if (chartSpec) {
    sections.push({
      type: "chart",
      title: chartSpec.title || "차트",
      chartSpec,
      rows: takeRows(rows, CHART_ROW_LIMIT),
      insight: chartSpec.insight || "",
    });
  }

  if (rows.length)
    sections.push(
      buildTableSection({ title: "분석 결과", rows, rowCount: rows.length }),
    );

  sections.push({
    type: "insight",
    title: "분석 인사이트",
    bullets: narrative.highlights || [],
  });

  const finalSections = decorateReportSections(sections);

  return {
    version: "report_sections_v2",
    reportType: "analysisReport",
    title,
    generatedAt,
    source: {
      fileName: fileName || "",
      message: message || "",
    },
    resultType: result?.resultType || "",
    operation: result?.operation || "",
    executiveSummary: buildExecutiveSummary({
      title,
      sections: finalSections,
      totalRows: rows.length,
    }),
    sections: finalSections,
  };
}

function buildReportSections({ fileName, message, result } = {}) {
  const normalizedBusinessResult = isBusinessTemplateResult(result)
    ? normalizeBusinessTemplateResult(result)
    : null;

  if (normalizedBusinessResult?.sections?.length) {
    return buildBusinessReportSections({
      fileName,
      message,
      result,
      normalizedBusinessResult,
    });
  }

  return buildGenericReportSections({ fileName, message, result });
}

module.exports = {
  buildReportSections,
};
