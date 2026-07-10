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

const ANALYSIS_REPORT_QUALITY_VERSION = "analysis_report_quality_v1";

const DOMAIN_NARRATIVE_DEFS = Object.freeze({
  sales: Object.freeze({
    label: "매출·영업",
    summary:
      "매출 규모, 구성비, 상위·하위 항목을 중심으로 실적 흐름을 정리했습니다.",
    action: "매출 비중이 큰 항목과 변동 폭이 큰 항목을 우선 확인하세요.",
  }),
  budget: Object.freeze({
    label: "예산·지출·정산",
    summary:
      "집행 규모, 비목별 구성, 잔액 또는 초과 가능성을 중심으로 정산 관점의 흐름을 정리했습니다.",
    action:
      "집행 비중이 큰 비목과 증빙 확인이 필요한 지출 항목을 우선 점검하세요.",
  }),
  hr: Object.freeze({
    label: "인사·조직",
    summary:
      "부서·직급별 인원 구성과 주요 변동 항목을 중심으로 조직 현황을 정리했습니다.",
    action:
      "특정 부서 또는 직급에 인원이 집중되는지, 변동이 큰 구간이 있는지 확인하세요.",
  }),
  inventory: Object.freeze({
    label: "재고·물류·자산",
    summary:
      "입출고 흐름, 보유 현황, 품목·창고별 편차를 중심으로 재고 관점의 흐름을 정리했습니다.",
    action:
      "재고 과다·부족 가능성이 있는 품목과 이동량이 큰 창고를 우선 확인하세요.",
  }),
  survey: Object.freeze({
    label: "설문·평가",
    summary:
      "평균 점수, 문항별 분포, 낮은 점수 항목을 중심으로 응답 경향을 정리했습니다.",
    action:
      "평균보다 낮은 문항과 응답 편차가 큰 문항을 개선 후보로 검토하세요.",
  }),
  operation: Object.freeze({
    label: "운영·상태",
    summary:
      "상태별 분포, 완료율, 미처리·지연 항목을 중심으로 운영 현황을 정리했습니다.",
    action: "미완료·지연 상태의 항목을 우선 조치 대상으로 분류하세요.",
  }),
  project: Object.freeze({
    label: "프로젝트·성과",
    summary:
      "목표 대비 실적, 달성률, 성과 편차를 중심으로 핵심 지표 흐름을 정리했습니다.",
    action: "목표 대비 미달 항목과 가중치가 큰 핵심 지표를 우선 확인하세요.",
  }),
  energy: Object.freeze({
    label: "에너지·비용",
    summary:
      "사용량, 절감량, 비용, 회수기간을 중심으로 에너지 운영 현황을 정리했습니다.",
    action:
      "사용량 또는 비용이 큰 항목과 절감 효과가 낮은 항목을 우선 확인하세요.",
  }),
  general: Object.freeze({
    label: "일반",
    summary:
      "주요 그룹, 수치 규모, 상위·하위 항목을 중심으로 결과를 정리했습니다.",
    action: "수치가 큰 항목과 변동 폭이 큰 항목을 우선 확인하세요.",
  }),
});

function cleanReportText(value = "") {
  return String(value ?? "")
    .replace(/undefined|null|NaN|\[object Object\]/gi, "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function uniqueBullets(items = [], limit = 8) {
  const seen = new Set();
  const result = [];
  for (const item of items.map(cleanReportText).filter(Boolean)) {
    const key = item.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(item);
    if (result.length >= limit) break;
  }
  return result;
}

function inferReportDomain({
  title = "",
  fileName = "",
  message = "",
  result = {},
  normalizedBusinessResult = null,
} = {}) {
  const source = cleanReportText(
    [
      normalizedBusinessResult?.domain,
      normalizedBusinessResult?.templateId,
      result?.templateId,
      result?.operation,
      title,
      fileName,
      message,
    ]
      .filter(Boolean)
      .join(" "),
  ).toLowerCase();

  if (/energy|에너지|전력|전기|절감|회수기간/.test(source)) return "energy";
  if (/kpi|성과|달성률|목표|실적|프로젝트|project/.test(source))
    return "project";
  if (/매출|영업|판매|revenue|sales|손익|profit/.test(source)) return "sales";
  if (/예산|집행|정산|카드|출장|회의비|budget|expense|cost|비용/.test(source))
    return "budget";
  if (/인사|직원|부서|직급|입사|퇴사|연봉|hr|employee/.test(source))
    return "hr";
  if (/재고|입고|출고|창고|자산|장비|inventory|asset|warehouse/.test(source))
    return "inventory";
  if (/만족도|설문|문항|평점|점수|survey|feedback/.test(source))
    return "survey";
  if (/상태|완료|처리|이수|점검|배송|운영|status|operation/.test(source))
    return "operation";
  return "general";
}

function getDomainNarrativeDef(domain = "general") {
  return DOMAIN_NARRATIVE_DEFS[domain] || DOMAIN_NARRATIVE_DEFS.general;
}

function getNumericFieldsForRows(rows = []) {
  return getRowKeys(rows).filter((key) => isMostlyNumeric(rows, key));
}

function formatKoreanNumber(value, field = "") {
  if (!isNumericValue(value)) return cleanReportText(value);
  const n = Number(value);
  const digits = /율|비율|rate|ratio|증감|달성/.test(String(field))
    ? 1
    : Number.isInteger(n)
      ? 0
      : 1;
  return n.toLocaleString("ko-KR", {
    maximumFractionDigits: digits,
    minimumFractionDigits: 0,
  });
}

function getExtremeRows(rows = [], valueField = "") {
  if (!rows.length || !valueField) return { max: null, min: null };
  const numericRows = rows
    .filter((row) => isNumericValue(row?.[valueField]))
    .slice()
    .sort((a, b) => Number(b[valueField]) - Number(a[valueField]));
  return {
    max: numericRows[0] || null,
    min: numericRows[numericRows.length - 1] || null,
  };
}

function buildSectionInsight(section = {}, domain = "general") {
  const rows = safeRows(section);
  const title = cleanReportText(getSectionTitle(section));
  const rowCount = getSectionRowCount(section, rows);
  const categoryField = pickCategoryField(rows);
  const valueField = pickValueField(rows, section);
  const categoryLabel = displayLabelForField(categoryField, section);
  const valueLabel = displayLabelForField(valueField, section);
  const def = getDomainNarrativeDef(domain);

  if (!rows.length) {
    return `${title}은(는) 요약 항목 중심으로 확인했습니다.`;
  }

  const { max, min } = getExtremeRows(rows, valueField);
  if (max && categoryField && valueField) {
    const maxLabel = cleanReportText(
      max?.[categoryField] ?? max?.label ?? max?.항목 ?? "상위 항목",
    );
    const maxValue = formatKoreanNumber(max?.[valueField], valueField);
    if (min && min !== max) {
      const minLabel = cleanReportText(
        min?.[categoryField] ?? min?.label ?? min?.항목 ?? "하위 항목",
      );
      const minValue = formatKoreanNumber(min?.[valueField], valueField);
      return `${title}: ${categoryLabel} 기준 ${valueLabel}은 ${maxLabel}(${maxValue})이 가장 크고, ${minLabel}(${minValue})이 가장 작습니다.`;
    }
    return `${title}: ${categoryLabel} 기준 ${valueLabel} 최상위 항목은 ${maxLabel}(${maxValue})입니다.`;
  }

  return `${title}: ${rowCount.toLocaleString("ko-KR")}건의 결과를 기준으로 ${def.label} 흐름을 확인했습니다.`;
}

function buildReportQualitySignals({
  sections = [],
  businessSections = [],
  domain = "general",
} = {}) {
  const summarySections = sections.filter(
    (section) => section.type === "summary",
  ).length;
  const insightSections = sections.filter(
    (section) => section.type === "insight",
  ).length;
  const chartSections = sections.filter(
    (section) => section.type === "chart",
  ).length;
  const tableSections = sections.filter(
    (section) => section.type === "table",
  ).length;
  return {
    version: ANALYSIS_REPORT_QUALITY_VERSION,
    domain,
    domainAwareNarrative: domain !== "general",
    summarySections,
    insightSections,
    chartSections,
    tableSections,
    businessSectionCount: businessSections.length,
    sanitizedText: true,
    duplicateBulletGuard: true,
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
  domain = "general",
}) {
  const tableSections = sections.filter((section) => section.type === "table");
  const chartSections = sections.filter((section) => section.type === "chart");
  const def = getDomainNarrativeDef(domain);
  const bullets = [
    `${def.label} 관점에서 ${businessSections.length || tableSections.length}개 분석 섹션을 정리했습니다.`,
    `${tableSections.length}개 표 섹션과 ${chartSections.length}개 차트 후보를 함께 확인했습니다.`,
  ];

  if (totalRows) {
    bullets.push(
      `${totalRows.toLocaleString("ko-KR")}건의 결과 행을 기준으로 핵심 흐름을 요약했습니다.`,
    );
  }

  const sectionInsights = businessSections
    .map((section) => buildSectionInsight(section, domain))
    .filter(Boolean);

  bullets.push(...sectionInsights.slice(0, 4));
  bullets.push(def.action);

  return uniqueBullets(bullets, 8);
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
  const domain = inferReportDomain({
    title,
    fileName,
    message,
    result,
    normalizedBusinessResult,
  });
  const domainDef = getDomainNarrativeDef(domain);
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
  const summaryBullets = buildSummaryBullets({
    businessSections,
    sections: decoratedBodySections,
    totalRows,
    domain,
  });
  const chartInsights = decoratedBodySections
    .filter((section) => section.type === "chart" && section.insight)
    .map((section) => section.insight);
  const sectionInsights = businessSections.map((section) =>
    buildSectionInsight(section, domain),
  );
  const insightBullets = uniqueBullets(
    [...chartInsights, ...sectionInsights, domainDef.action],
    8,
  );

  const sections = decorateReportSections([
    {
      type: "cover",
      title,
      subtitle: [fileName || "", domainDef.label].filter(Boolean).join(" · "),
      generatedAt,
    },
    {
      type: "summary",
      title: "핵심 요약",
      summary: `${domainDef.summary} 총 ${businessSections.length.toLocaleString("ko-KR")}개 분석 섹션과 ${totalRows.toLocaleString("ko-KR")}건의 결과 행을 기준으로 작성했습니다.`,
      bullets: summaryBullets,
    },
    ...decoratedBodySections,
    {
      type: "insight",
      title: "분석 인사이트",
      bullets: insightBullets,
    },
  ]);

  return {
    version: "report_sections_v2",
    reportType: "analysisReport",
    qualityVersion: ANALYSIS_REPORT_QUALITY_VERSION,
    title,
    domain,
    domainLabel: domainDef.label,
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
    qualitySignals: buildReportQualitySignals({
      sections,
      businessSections,
      domain,
    }),
    sections,
  };
}

function buildGenericReportSections({ fileName, message, result } = {}) {
  const narrative = buildNarrativeSections(result, {
    message,
    fileName,
  });
  const domain =
    narrative.domain ||
    inferReportDomain({
      title: narrative.title,
      fileName,
      message,
      result,
    });
  const domainDef = getDomainNarrativeDef(domain);
  const chartSpec = recommendChartSpec(result);
  const rows = safeRows(result);
  const title = narrative.title || result?.title || "분석 보고서";
  const generatedAt = new Date().toISOString();

  const sections = [
    {
      type: "cover",
      title,
      subtitle: [fileName || "", domainDef.label].filter(Boolean).join(" · "),
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
    qualityVersion: ANALYSIS_REPORT_QUALITY_VERSION,
    title,
    domain,
    domainLabel: domainDef.label,
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
    qualitySignals: buildReportQualitySignals({
      sections: finalSections,
      businessSections: [],
      domain,
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
