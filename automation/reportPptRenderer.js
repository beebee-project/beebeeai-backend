const pptxgen = require("pptxgenjs");
const {
  cleanDisplayTitle,
  formatDisplayValue,
  getDisplayFieldLabel,
  tableColumnLabels,
  toDisplayNumber,
} = require("./reportDisplayUtils");

function text(v) {
  return String(v ?? "");
}

function truncate(v, max = 90) {
  const s = text(v);
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

function formatDateTime(value) {
  const d = value ? new Date(value) : new Date();
  if (Number.isNaN(d.getTime())) return text(value || "");
  return new Intl.DateTimeFormat("ko-KR", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(d);
}

const PPT_TEMPLATES = {
  default: {
    name: "default",
    fontFace: "Malgun Gothic",
    titleFontSize: 22,
    bodyFontSize: 14,
    tableFontSize: 8.5,
    coverTitleFontSize: 28,
    titleX: 0.5,
    titleY: 0.36,
    titleW: 9,
    titleH: 0.56,
    contentX: 0.7,
    contentY: 1.16,
    contentW: 8.6,
    chartY: 1.32,
    chartH: 3.25,
    tableY: 1.08,
    tableH: 4.08,
    footerY: 5.43,
    maxTableRowsPerSlide: 8,
    maxTableCols: 6,
  },
  minimal: {
    name: "minimal",
    fontFace: "Malgun Gothic",
    titleFontSize: 20,
    bodyFontSize: 13,
    tableFontSize: 8,
    coverTitleFontSize: 26,
    titleX: 0.6,
    titleY: 0.42,
    titleW: 8.8,
    titleH: 0.5,
    contentX: 0.75,
    contentY: 1.2,
    contentW: 8.4,
    chartY: 1.3,
    chartH: 3.2,
    tableY: 1.1,
    tableH: 4.05,
    footerY: 5.43,
    maxTableRowsPerSlide: 9,
    maxTableCols: 6,
  },
};

const PPT_QUALITY_VERSION = "ppt_quality_v1";
const PPT_QUALITY_DEFAULT_MAX_SLIDES = 16;

const PPT_DOMAIN_DEFS = Object.freeze({
  sales: Object.freeze({ label: "매출·영업", base: "매출" }),
  budget: Object.freeze({ label: "예산·지출", base: "예산" }),
  hr: Object.freeze({ label: "인사·조직", base: "인원" }),
  inventory: Object.freeze({ label: "재고·자산", base: "재고" }),
  survey: Object.freeze({ label: "설문·만족도", base: "만족도" }),
  operation: Object.freeze({ label: "운영·상태", base: "처리 현황" }),
  project: Object.freeze({ label: "프로젝트·성과", base: "성과" }),
  customer: Object.freeze({ label: "고객·CS", base: "고객 문의" }),
  energy: Object.freeze({ label: "에너지", base: "에너지 사용량" }),
  general: Object.freeze({ label: "분석", base: "주요 지표" }),
});

function compactWhitespace(value = "") {
  return text(value).replace(/\s+/g, " ").trim();
}

function sanitizePptText(value = "") {
  return compactWhitespace(cleanDisplayTitle(value))
    .replace(/undefined|NaN|\[object Object\]|Infinity/g, "")
    .trim();
}

function getPptMaxSlides(options = {}) {
  const raw = Number(
    options.maxSlides ||
      process.env.PPT_QUALITY_MAX_SLIDES ||
      PPT_QUALITY_DEFAULT_MAX_SLIDES,
  );
  if (!Number.isFinite(raw) || raw < 8) return PPT_QUALITY_DEFAULT_MAX_SLIDES;
  return Math.min(Math.floor(raw), 24);
}

function inferPptDomain(report = {}) {
  const source = compactWhitespace(
    [
      report.domain,
      report.domainLabel,
      report.title,
      report.operation,
      report.resultType,
      report.source?.fileName,
      report.source?.message,
    ]
      .filter(Boolean)
      .join(" "),
  );

  if (/에너지|사용량|절감|회수기간|energy/i.test(source)) return "energy";
  if (/매출|판매|손익|비용|이익|profit|revenue|sales/i.test(source))
    return "sales";
  if (/예산|집행|정산|카드|회의비|출장비|budget|expense/i.test(source))
    return "budget";
  if (/인사|직원|구성원|참여연구원|채용|부서|직급|hr|employee/i.test(source))
    return "hr";
  if (
    /재고|입출고|창고|자산|장비|소모품|inventory|warehouse|asset/i.test(source)
  )
    return "inventory";
  if (/설문|만족도|평가|점수|survey|score|feedback/i.test(source))
    return "survey";
  if (/프로젝트|성과|KPI|달성률|지원사업|project|performance/i.test(source))
    return "project";
  if (/고객|문의|민원|customer|cs/i.test(source)) return "customer";
  if (/상태|처리|이수|점검|배송|운영|status|operation/i.test(source))
    return "operation";
  return "general";
}

function domainDef(report = {}) {
  return PPT_DOMAIN_DEFS[inferPptDomain(report)] || PPT_DOMAIN_DEFS.general;
}

function looksMachineGeneratedTitle(title = "") {
  const value = text(title);
  const underscoreCount = (value.match(/_/g) || []).length;
  return (
    /\d{4}[-_]\d{2}/.test(value) || underscoreCount >= 3 || value.length > 48
  );
}

function compressSlideTitle(title = "", section = {}, report = {}) {
  const cleaned = sanitizePptText(title || section.title || "분석 결과");
  const domain = domainDef(report);
  const source = `${cleaned} ${section.operation || ""} ${section.resultType || ""}`;
  const machineLike = looksMachineGeneratedTitle(cleaned);

  if (machineLike || cleaned.length > 42) {
    if (/상위|하위|top|bottom/i.test(source))
      return `${domain.base} 상위·하위 항목`;
    if (/구성비|비중|composition|ratio/i.test(source))
      return `${domain.base} 구성비`;
    if (/평균|average|avg|mean/i.test(source)) return `${domain.base} 평균`;
    if (/합계|총액|sum|total/i.test(source)) return `${domain.base} 합계`;
    if (/추이|time|series|월|기간|연도/i.test(source))
      return `${domain.base} 추이`;
    if (/달성|목표|KPI|성과/i.test(source)) return "목표 대비 성과";
    if (/상태|처리|완료|이수/i.test(source)) return `${domain.base} 상태 요약`;
    return `${domain.base} 분석 요약`;
  }

  return truncate(cleaned, 42);
}

function summarizeSourceName(report = {}) {
  return sanitizePptText(report.source?.fileName || report.fileName || "");
}

function buildCoverSection(report = {}) {
  const existing = (Array.isArray(report.sections) ? report.sections : []).find(
    (section) => section.type === "cover",
  );
  const domain = domainDef(report);
  return {
    ...(existing || {}),
    type: "cover",
    title: sanitizePptText(existing?.title || report.title || "분석 보고서"),
    subtitle: [domain.label, summarizeSourceName(report)]
      .filter(Boolean)
      .join(" · "),
    generatedAt:
      existing?.generatedAt || report.generatedAt || new Date().toISOString(),
    pptQualityVersion: PPT_QUALITY_VERSION,
  };
}

function collectSummaryBullets(report = {}, originalSections = []) {
  const executive = report.executiveSummary || {};
  const summarySections = originalSections.filter(
    (section) => section.type === "summary",
  );
  const bullets = [];

  if (Number.isFinite(Number(executive.sectionCount))) {
    bullets.push(
      `${Number(executive.sectionCount)}개 분석 섹션을 핵심 흐름 중심으로 정리했습니다.`,
    );
  }
  if (Number.isFinite(Number(executive.chartSectionCount))) {
    bullets.push(
      `${Number(executive.chartSectionCount)}개 차트 후보와 주요 표를 PPT로 요약했습니다.`,
    );
  }
  if (
    Number.isFinite(Number(executive.totalRows)) &&
    Number(executive.totalRows) > 0
  ) {
    bullets.push(
      `총 ${new Intl.NumberFormat("ko-KR").format(Number(executive.totalRows))}건의 결과 행을 검토했습니다.`,
    );
  }

  for (const section of summarySections) {
    if (section.summary) bullets.push(section.summary);
    if (Array.isArray(section.bullets)) bullets.push(...section.bullets);
  }

  return Array.from(
    new Set(bullets.map(sanitizePptText).filter(Boolean)),
  ).slice(0, 5);
}

function buildExecutiveSummarySlide(report = {}, originalSections = []) {
  const domain = domainDef(report);
  return {
    type: "summary",
    title: "Executive Summary",
    summary: `${domain.label} 관점에서 핵심 지표와 확인 포인트를 요약했습니다.`,
    bullets: collectSummaryBullets(report, originalSections),
    pptQualityRole: "executiveSummary",
  };
}

function getSectionRows(section = {}) {
  return Array.isArray(section.rows) ? section.rows : [];
}

function scoreBodySection(section = {}, index = 0) {
  const title = text(section.title);
  let score = 1000 - index;
  if (section.type === "chart") score += 500;
  if (section.type === "table") score += 150;
  if (/요약|추이|구성비|상위|하위|평균|합계|달성|상태|현황/.test(title))
    score += 200;
  if (/전체\s*대상\s*수|총\s*건수/.test(title)) score -= 250;
  if (getSectionRows(section).length <= 1 && section.type === "table")
    score -= 120;
  return score;
}

function dedupeBodySections(bodySections = [], report = {}) {
  const seen = new Set();
  const deduped = [];
  for (const section of bodySections) {
    const key = compressSlideTitle(
      section.title,
      section,
      report,
    ).toLowerCase();
    const typeKey = `${section.type}:${key}`;
    if (seen.has(typeKey)) continue;
    seen.add(typeKey);
    deduped.push(section);
  }
  return deduped;
}

function limitTableRowsForSingleSlide(
  section = {},
  template = PPT_TEMPLATES.default,
) {
  if (section.type !== "table") return section;
  const rows = getSectionRows(section);
  const maxRows = Number(template.maxTableRowsPerSlide || 8);
  if (rows.length <= maxRows) return section;
  return {
    ...section,
    rows: rows.slice(0, maxRows),
    note: section.note || `PPT 가독성을 위해 상위 ${maxRows}건만 표시했습니다.`,
    pptQualityRowLimitApplied: true,
  };
}

function prepareBodySections(
  report = {},
  originalSections = [],
  template = PPT_TEMPLATES.default,
  maxSlides = PPT_QUALITY_DEFAULT_MAX_SLIDES,
) {
  const bodyLimit = Math.max(3, maxSlides - 4);
  const body = originalSections.filter(
    (section) => section.type === "chart" || section.type === "table",
  );
  const deduped = dedupeBodySections(body, report);
  const selected = deduped
    .map((section, index) => ({
      section,
      index,
      score: scoreBodySection(section, index),
    }))
    .sort((a, b) => b.score - a.score)
    .slice(0, bodyLimit)
    .sort((a, b) => a.index - b.index)
    .map(({ section }) => {
      const title = compressSlideTitle(section.title, section, report);
      return limitTableRowsForSingleSlide(
        {
          ...section,
          pptOriginalTitle: section.title || "",
          title,
          pptQualityTitleApplied: title !== section.title,
        },
        template,
      );
    });
  return selected;
}

function collectInsightBullets(originalSections = []) {
  const bullets = [];
  for (const section of originalSections) {
    if (section.type === "insight" && Array.isArray(section.bullets))
      bullets.push(...section.bullets);
    if (section.type === "chart" && section.insight)
      bullets.push(section.insight);
    if (section.chartSpec?.insight) bullets.push(section.chartSpec.insight);
  }
  return Array.from(
    new Set(bullets.map(sanitizePptText).filter(Boolean)),
  ).slice(0, 5);
}

function buildClosingSection(report = {}, originalSections = []) {
  const domain = domainDef(report);
  const bullets = collectInsightBullets(originalSections);
  return {
    type: "closing",
    title: "검토 및 다음 단계",
    bullets: bullets.length
      ? bullets.slice(0, 4)
      : [
          `${domain.label} 관점에서 주요 수치와 변동 항목을 추가 확인하세요.`,
          "원본데이터와 자동화시트를 함께 검토하면 세부 원인을 확인할 수 있습니다.",
        ],
    pptQualityRole: "closing",
  };
}

function buildPptQualityPlan(
  report = {},
  template = PPT_TEMPLATES.default,
  options = {},
) {
  const sourceSections = Array.isArray(report.sections) ? report.sections : [];
  const maxSlides = getPptMaxSlides(options);
  const cover = buildCoverSection(report);
  const executiveSummary = buildExecutiveSummarySlide(report, sourceSections);
  const body = prepareBodySections(report, sourceSections, template, maxSlides);
  const closing = buildClosingSection(report, sourceSections);
  const plannedSections = [cover, executiveSummary, ...body, closing].slice(
    0,
    maxSlides,
  );
  const chartCount = plannedSections.filter(
    (section) => section.type === "chart",
  ).length;
  const tableCount = plannedSections.filter(
    (section) => section.type === "table",
  ).length;

  return {
    report: {
      ...report,
      pptQualityVersion: PPT_QUALITY_VERSION,
      pptQualitySignals: {
        version: PPT_QUALITY_VERSION,
        originalSectionCount: sourceSections.length,
        plannedSectionCount: plannedSections.length,
        maxSlides,
        titleCompression: true,
        executiveSummarySlide: true,
        closingSlide: true,
        tableRowLimit: true,
        chartSlideCount: chartCount,
        tableSlideCount: tableCount,
      },
      sections: plannedSections,
    },
    meta: {
      version: PPT_QUALITY_VERSION,
      originalSectionCount: sourceSections.length,
      plannedSectionCount: plannedSections.length,
      maxSlides,
      chartSlideCount: chartCount,
      tableSlideCount: tableCount,
      executiveSummarySlide: true,
      closingSlide: true,
      titleCompression: true,
    },
  };
}

function getTemplate(options = {}) {
  const key = options.template || "default";
  return PPT_TEMPLATES[key] || PPT_TEMPLATES.default;
}

function addFooter(
  slide,
  report = {},
  pageNo = 0,
  template = PPT_TEMPLATES.default,
) {
  const source = report.source?.fileName || report.fileName || "";
  slide.addText(`BeeBee AI${source ? ` | ${source}` : ""}`, {
    x: template.contentX,
    y: template.footerY,
    w: 6.8,
    h: 0.22,
    fontFace: template.fontFace,
    fontSize: 8,
    color: "666666",
  });
  slide.addText(String(pageNo || ""), {
    x: 9.0,
    y: template.footerY,
    w: 0.4,
    h: 0.22,
    fontFace: template.fontFace,
    fontSize: 8,
    color: "666666",
    align: "right",
  });
}

function addTitle(slide, title, y = null, template = PPT_TEMPLATES.default) {
  slide.addText(truncate(cleanDisplayTitle(title), 70), {
    x: template.titleX,
    y: y ?? template.titleY,
    w: template.titleW,
    h: template.titleH,
    fontFace: template.fontFace,
    fontSize: template.titleFontSize,
    bold: true,
    fit: "shrink",
  });
}

function addBullets(
  slide,
  bullets = [],
  y = null,
  template = PPT_TEMPLATES.default,
) {
  const safe = Array.isArray(bullets) ? bullets.slice(0, 5) : [];

  slide.addText(
    safe.map((b) => `• ${truncate(cleanDisplayTitle(b), 100)}`).join("\n") ||
      "요약 내용이 없습니다.",
    {
      x: template.contentX,
      y: y ?? template.contentY,
      w: template.contentW,
      h: 3.8,
      fontFace: template.fontFace,
      fontSize: template.bodyFontSize,
      breakLine: false,
      fit: "shrink",
      valign: "top",
    },
  );
}

function chunkRows(rows = [], size = 8) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const chunks = [];

  for (let i = 0; i < safeRows.length; i += size) {
    chunks.push(safeRows.slice(i, i + size));
  }

  return chunks.length ? chunks : [[]];
}

function tableColumnKeys(rows = [], maxCols = 6) {
  const first = Array.isArray(rows) && rows.length ? rows[0] : {};
  return Object.keys(first || {}).slice(0, maxCols);
}

function tableContext(section = {}) {
  return {
    title: section.title,
    sectionTitle: section.title,
    chartSpec: section.chartSpec,
    fieldLabels: section.columnLabels || section.chartSpec?.fieldLabels,
  };
}

function rowsToTableWithKeys(rows = [], keys = [], context = {}) {
  if (!rows.length || !keys.length) return [["결과 없음"]];
  const labels = context.fieldLabels || tableColumnLabels(rows, context);
  const header = keys.map(
    (key) => labels[key] || getDisplayFieldLabel(key, context),
  );
  const body = rows.map((row) =>
    keys.map((key) => formatDisplayValue(row[key], key, context)),
  );
  return [header, ...body];
}

function rowsToTable(rows = [], limit = 8, context = {}) {
  const safeRows = Array.isArray(rows) ? rows.slice(0, limit) : [];
  const keys = tableColumnKeys(safeRows);
  return rowsToTableWithKeys(safeRows, keys, context);
}

function addTable(
  slide,
  rows = [],
  template = PPT_TEMPLATES.default,
  context = {},
) {
  slide.addTable(rowsToTable(rows, template.maxTableRowsPerSlide, context), {
    x: 0.5,
    y: template.tableY,
    w: 9,
    h: template.tableH,
    fontFace: template.fontFace,
    fontSize: template.tableFontSize,
    border: { type: "solid", pt: 0.5 },
    fit: "shrink",
  });
}

function addPagedTableSlides(
  pptx,
  section = {},
  template = PPT_TEMPLATES.default,
  report = {},
  pageState,
) {
  const rows = Array.isArray(section.rows) ? section.rows : [];
  const keys = tableColumnKeys(rows, template.maxTableCols);
  const chunks = chunkRows(rows, template.maxTableRowsPerSlide);
  const totalPages = chunks.length;
  const context = tableContext(section);

  chunks.forEach((chunk, idx) => {
    const slide = pptx.addSlide();
    pageState.page += 1;
    const title =
      totalPages > 1
        ? `${section.title || "분석 결과"} (${idx + 1}/${totalPages})`
        : section.title || "분석 결과";

    addTitle(slide, title, null, template);
    slide.addTable(rowsToTableWithKeys(chunk, keys, context), {
      x: 0.5,
      y: template.tableY,
      w: 9,
      h: template.tableH,
      fontFace: template.fontFace,
      fontSize: template.tableFontSize,
      border: { type: "solid", pt: 0.5 },
      fit: "shrink",
    });

    const note =
      section.note || `전체 행 수: ${section.rowCount || rows.length}`;
    slide.addText(cleanDisplayTitle(note), {
      x: 0.5,
      y: 5.18,
      w: 7.8,
      h: 0.28,
      fontFace: template.fontFace,
      fontSize: 8.5,
      color: "666666",
      fit: "shrink",
    });
    addFooter(slide, report, pageState.page, template);
  });
}

function mapChartType(type = "") {
  if (type === "line") return "line";
  if (type === "horizontal_bar") return "bar";
  if (type === "grouped_bar") return "bar";
  if (type === "stacked_bar") return "bar";
  return "bar";
}

function getChartRows(section = {}) {
  return Array.isArray(section.rows) ? section.rows.slice(0, 12) : [];
}

function buildChartSeries(section = {}) {
  const spec = section.chartSpec || {};
  const rows = getChartRows(section);

  if (!rows.length) return null;

  const categoryField = spec.categoryField;
  const seriesFields = Array.isArray(spec.seriesFields)
    ? spec.seriesFields
    : spec.valueField
      ? [spec.valueField]
      : [];

  if (!categoryField || !seriesFields.length) return null;

  const context = {
    title: section.title,
    sectionTitle: section.title,
    chartSpec: spec,
    fieldLabels: spec.fieldLabels,
  };
  const labels = rows.map((r) => text(r[categoryField] ?? ""));

  const series = seriesFields.map((field) => ({
    name: getDisplayFieldLabel(field, context),
    labels,
    values: rows.map((r) => {
      const raw =
        r[field] ?? (r.metric === field ? r.value : undefined) ?? r.value;
      return toDisplayNumber(raw, field, context);
    }),
  }));

  return {
    labels,
    series,
  };
}

function addChart(slide, section = {}, template = PPT_TEMPLATES.default) {
  const spec = section.chartSpec || {};
  const chartData = buildChartSeries(section);

  if (!chartData) {
    addTable(slide, section.rows || [], template, tableContext(section));
    return;
  }

  const chartType = mapChartType(spec.recommendedType);

  slide.addChart(chartType, chartData.series, {
    x: template.contentX,
    y: template.chartY,
    w: template.contentW,
    h: template.chartH,
    showLegend: chartData.series.length > 1,
    showTitle: false,
    showValue: false,
    catAxisLabelFontFace: template.fontFace,
    catAxisLabelFontSize: 8.5,
    valAxisLabelFontFace: template.fontFace,
    valAxisLabelFontSize: 8.5,
  });

  const insight = spec.insight || section.insight;
  if (insight) {
    slide.addText(truncate(cleanDisplayTitle(insight), 110), {
      x: template.contentX,
      y: 4.62,
      w: template.contentW,
      h: 0.28,
      fontFace: template.fontFace,
      fontSize: 9.5,
      color: "555555",
      fit: "shrink",
    });
  }

  if (section.rows?.length) {
    slide.addText(`차트 데이터: ${section.rows.length}건`, {
      x: template.contentX,
      y: 4.92,
      w: 8,
      h: 0.25,
      fontFace: template.fontFace,
      fontSize: 8.5,
      color: "777777",
    });
  }
}

function renderReportPpt(report = {}, options = {}) {
  const template = getTemplate(options);
  const qualityPlan = buildPptQualityPlan(report, template, options);
  const pptReport = qualityPlan.report;
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "BeeBeeAI";
  pptx.subject = pptReport.title || "분석 보고서";
  pptx.title = pptReport.title || "분석 보고서";
  pptx.company = "BeeBee AI";
  pptx.lang = "ko-KR";

  const sections = Array.isArray(pptReport.sections) ? pptReport.sections : [];
  const pageState = { page: 0 };

  for (const section of sections) {
    if (section.type === "table") {
      addPagedTableSlides(pptx, section, template, pptReport, pageState);
      continue;
    }

    const slide = pptx.addSlide();
    pageState.page += 1;

    if (section.type === "cover") {
      slide.addText(text(section.title || pptReport.title || "분석 보고서"), {
        x: template.contentX,
        y: 1.45,
        w: template.contentW,
        h: 0.8,
        fontFace: template.fontFace,
        fontSize: template.coverTitleFontSize,
        bold: true,
        fit: "shrink",
      });

      slide.addText(
        text(section.subtitle || pptReport.source?.fileName || ""),
        {
          x: template.contentX,
          y: 2.45,
          w: template.contentW,
          h: 0.34,
          fontFace: template.fontFace,
          fontSize: template.bodyFontSize,
          color: "555555",
        },
      );

      if (pptReport.source?.message) {
        slide.addText(cleanDisplayTitle(pptReport.source.message), {
          x: template.contentX,
          y: 2.88,
          w: template.contentW,
          h: 0.34,
          fontFace: template.fontFace,
          fontSize: 10.5,
          color: "666666",
          fit: "shrink",
        });
      }

      slide.addText(
        `생성일시: ${formatDateTime(section.generatedAt || pptReport.generatedAt)}`,
        {
          x: template.contentX,
          y: 3.35,
          w: template.contentW,
          h: 0.3,
          fontFace: template.fontFace,
          fontSize: 10,
          color: "666666",
        },
      );
      addFooter(slide, pptReport, pageState.page, template);
      continue;
    }

    if (section.type === "summary") {
      addTitle(slide, section.title || "핵심 요약", null, template);
      addBullets(
        slide,
        [section.summary, ...(section.bullets || [])].filter(Boolean),
        null,
        template,
      );
      addFooter(slide, pptReport, pageState.page, template);
      continue;
    }

    if (section.type === "chart") {
      addTitle(
        slide,
        section.title || section.chartSpec?.title || "차트",
        null,
        template,
      );
      addChart(slide, section, template);
      addFooter(slide, pptReport, pageState.page, template);
      continue;
    }

    if (section.type === "insight") {
      addTitle(slide, section.title || "분석 인사이트", null, template);
      addBullets(slide, section.bullets || [], null, template);
      addFooter(slide, pptReport, pageState.page, template);
      continue;
    }

    if (section.type === "closing") {
      addTitle(slide, section.title || "검토 및 다음 단계", null, template);
      addBullets(slide, section.bullets || [], null, template);
      slide.addText(
        "자동 생성 PPT는 원본데이터와 함께 검토하는 것을 권장합니다.",
        {
          x: template.contentX,
          y: 4.95,
          w: template.contentW,
          h: 0.28,
          fontFace: template.fontFace,
          fontSize: 9,
          color: "777777",
          fit: "shrink",
        },
      );
      addFooter(slide, pptReport, pageState.page, template);
      continue;
    }

    addTitle(slide, section.title || section.type || "섹션", null, template);
    addFooter(slide, pptReport, pageState.page, template);
  }

  if (!sections.length) {
    const slide = pptx.addSlide();
    pageState.page += 1;
    addTitle(slide, "분석 보고서", null, template);
    slide.addText("보고서 섹션이 없습니다.", {
      x: template.contentX,
      y: template.contentY,
      w: template.contentW,
      h: 0.5,
      fontFace: template.fontFace,
      fontSize: template.bodyFontSize,
    });
    addFooter(slide, pptReport, pageState.page, template);
  }

  pptx._beebeeSlideCount = pageState.page;
  pptx._beebeePptQualityVersion = PPT_QUALITY_VERSION;
  pptx._beebeePptQuality = {
    ...qualityPlan.meta,
    renderedSlideCount: pageState.page,
  };
  pptx._beebeePptReport = pptReport;

  return pptx;
}

module.exports = {
  PPT_QUALITY_VERSION,
  renderReportPpt,
};
