const pptxgen = require("pptxgenjs");

function text(v) {
  return String(v ?? "");
}

function truncate(v, max = 140) {
  const s = text(v).replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

function formatDateKst(value = null) {
  const date = value ? new Date(value) : new Date();
  if (Number.isNaN(date.getTime())) return "";

  const kst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  const y = kst.getUTCFullYear();
  const m = String(kst.getUTCMonth() + 1).padStart(2, "0");
  const d = String(kst.getUTCDate()).padStart(2, "0");
  const hh = String(kst.getUTCHours()).padStart(2, "0");
  const mm = String(kst.getUTCMinutes()).padStart(2, "0");
  return `${y}.${m}.${d} ${hh}:${mm}`;
}

const PPT_TEMPLATES = {
  default: {
    name: "default",
    fontFace: "Malgun Gothic",
    titleFontSize: 22,
    bodyFontSize: 14,
    tableFontSize: 9,
    coverTitleFontSize: 30,
    titleX: 0.55,
    titleY: 0.35,
    titleW: 9.0,
    titleH: 0.55,
    contentX: 0.75,
    contentY: 1.15,
    contentW: 8.8,
    chartY: 1.35,
    chartH: 3.25,
    tableY: 1.15,
    tableH: 4.05,
    footerY: 5.35,
    maxTableRowsPerSlide: 8,
    maxTableCols: 6,
    maxSummaryBullets: 5,
    chartValueAxisFontSize: 9,
    chartCategoryAxisFontSize: 9,
    chartLegendFontSize: 9,
    chartShowGrid: true,
    chartShowValue: false,
    chartShowCategoryName: false,
    chartRoundedCorners: true,
  },
  minimal: {
    name: "minimal",
    fontFace: "Malgun Gothic",
    titleFontSize: 20,
    bodyFontSize: 13,
    tableFontSize: 8,
    coverTitleFontSize: 28,
    titleX: 0.6,
    titleY: 0.4,
    titleW: 8.8,
    titleH: 0.55,
    contentX: 0.8,
    contentY: 1.15,
    contentW: 8.5,
    chartY: 1.35,
    chartH: 3.2,
    tableY: 1.15,
    tableH: 4.0,
    footerY: 5.35,
    maxTableRowsPerSlide: 9,
    maxTableCols: 6,
    maxSummaryBullets: 5,
    chartValueAxisFontSize: 9,
    chartCategoryAxisFontSize: 9,
    chartLegendFontSize: 9,
    chartShowGrid: true,
    chartShowValue: false,
    chartShowCategoryName: false,
    chartRoundedCorners: true,
  },
};

function getTemplate(options = {}) {
  const key = options.template || "default";
  return PPT_TEMPLATES[key] || PPT_TEMPLATES.default;
}

function addFooter(
  slide,
  report = {},
  pageNo = 1,
  template = PPT_TEMPLATES.default,
) {
  const sourceFile = report.source?.fileName
    ? ` | ${report.source.fileName}`
    : "";
  slide.addText(`BeeBee AI${sourceFile}`, {
    x: 0.55,
    y: template.footerY,
    w: 6.8,
    h: 0.25,
    fontFace: template.fontFace,
    fontSize: 8,
    color: "666666",
  });

  slide.addText(String(pageNo), {
    x: 9.2,
    y: template.footerY,
    w: 0.45,
    h: 0.25,
    fontFace: template.fontFace,
    fontSize: 8,
    color: "666666",
    align: "right",
  });
}

function addTitle(slide, title, y = null, template = PPT_TEMPLATES.default) {
  slide.addText(truncate(title, 70), {
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
  const safe = Array.isArray(bullets)
    ? bullets.filter(Boolean).slice(0, template.maxSummaryBullets || 5)
    : [];

  slide.addText(
    safe.map((b) => `• ${truncate(b, 110)}`).join("\n") ||
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

function rowsToTableWithKeys(rows = [], keys = []) {
  if (!rows.length || !keys.length) return [["결과 없음"]];

  const body = rows.map((row) => keys.map((h) => truncate(row[h], 80)));
  return [keys, ...body];
}

function rowsToTable(rows = [], limit = 8, maxCols = 6) {
  const safeRows = Array.isArray(rows) ? rows.slice(0, limit) : [];
  const keys = tableColumnKeys(safeRows, maxCols);

  return rowsToTableWithKeys(safeRows, keys);
}

function addTable(slide, rows = [], template = PPT_TEMPLATES.default) {
  slide.addTable(
    rowsToTable(rows, template.maxTableRowsPerSlide, template.maxTableCols),
    {
      x: 0.55,
      y: template.tableY,
      w: 9.1,
      h: template.tableH,
      fontFace: template.fontFace,
      fontSize: template.tableFontSize,
      border: { type: "solid", pt: 0.5 },
      fit: "shrink",
    },
  );
}

function addPagedTableSlides(
  pptx,
  section = {},
  template = PPT_TEMPLATES.default,
  report = {},
  pageNo = 1,
) {
  const rows = Array.isArray(section.rows) ? section.rows : [];
  const keys = tableColumnKeys(rows, template.maxTableCols);
  const chunks = chunkRows(rows, template.maxTableRowsPerSlide);
  const totalPages = chunks.length;
  let currentPage = pageNo;

  chunks.forEach((chunk, idx) => {
    const slide = pptx.addSlide();
    const title =
      totalPages > 1
        ? `${section.title || "분석 결과"} (${idx + 1}/${totalPages})`
        : section.title || "분석 결과";

    addTitle(slide, title, null, template);
    slide.addTable(rowsToTableWithKeys(chunk, keys), {
      x: 0.55,
      y: template.tableY,
      w: 9.1,
      h: template.tableH,
      fontFace: template.fontFace,
      fontSize: template.tableFontSize,
      border: { type: "solid", pt: 0.5 },
      fit: "shrink",
    });

    const note =
      section.note ||
      `전체 행 수: ${Number(section.rowCount || rows.length).toLocaleString()}`;
    slide.addText(note, {
      x: 0.55,
      y: 5.05,
      w: 8.8,
      h: 0.25,
      fontFace: template.fontFace,
      fontSize: 9,
      color: "666666",
      fit: "shrink",
    });

    addFooter(slide, report, currentPage, template);
    currentPage += 1;
  });

  return currentPage;
}

function renderReportPpt(report = {}, options = {}) {
  const template = getTemplate(options);
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "BeeBeeAI";
  pptx.subject = report.title || "분석 보고서";
  pptx.title = report.title || "분석 보고서";
  pptx.company = "BeeBee AI";
  pptx.lang = "ko-KR";

  const sections = Array.isArray(report.sections) ? report.sections : [];
  let pageNo = 1;

  for (const section of sections) {
    if (section.type === "table") {
      pageNo = addPagedTableSlides(pptx, section, template, report, pageNo);
      continue;
    }

    const slide = pptx.addSlide();

    if (section.type === "cover") {
      slide.addText(text(section.title || report.title || "분석 보고서"), {
        x: template.contentX,
        y: 1.35,
        w: template.contentW,
        h: 0.85,
        fontFace: template.fontFace,
        fontSize: template.coverTitleFontSize,
        bold: true,
        fit: "shrink",
      });

      const metaLines = [
        section.subtitle || report.source?.fileName || "",
        report.source?.message || "",
        `생성일시: ${formatDateKst(section.generatedAt || report.generatedAt)}`,
      ].filter(Boolean);

      slide.addText(metaLines.map((line) => truncate(line, 100)).join("\n"), {
        x: template.contentX,
        y: 2.45,
        w: template.contentW,
        h: 1.0,
        fontFace: template.fontFace,
        fontSize: template.bodyFontSize,
        color: "555555",
        fit: "shrink",
      });

      addFooter(slide, report, pageNo, template);
      pageNo += 1;
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
      addFooter(slide, report, pageNo, template);
      pageNo += 1;
      continue;
    }

    if (section.type === "chart") {
      addTitle(slide, section.title || "차트", null, template);
      addChart(slide, section, template);

      const insight = section.insight || section.chartSpec?.insight || "";
      if (insight) {
        slide.addText(truncate(insight, 120), {
          x: template.contentX,
          y: 4.85,
          w: template.contentW,
          h: 0.3,
          fontFace: template.fontFace,
          fontSize: 10,
          color: "555555",
          fit: "shrink",
        });
      }

      addFooter(slide, report, pageNo, template);
      pageNo += 1;
      continue;
    }

    if (section.type === "insight") {
      addTitle(slide, section.title || "분석 인사이트", null, template);
      addBullets(slide, section.bullets || [], null, template);
      addFooter(slide, report, pageNo, template);
      pageNo += 1;
      continue;
    }

    addTitle(slide, section.title || section.type || "섹션", null, template);
    addFooter(slide, report, pageNo, template);
    pageNo += 1;
  }

  if (!sections.length) {
    const slide = pptx.addSlide();
    addTitle(slide, "분석 보고서", null, template);
    slide.addText("보고서 섹션이 없습니다.", {
      x: template.contentX,
      y: template.contentY,
      w: template.contentW,
      h: 0.5,
      fontFace: template.fontFace,
      fontSize: template.bodyFontSize,
    });
    addFooter(slide, report, pageNo, template);
  }

  return pptx;
}

function mapChartType(type = "") {
  if (type === "line") return "line";
  if (type === "horizontal_bar") return "bar";
  if (type === "grouped_bar") return "bar";
  if (type === "stacked_bar") return "bar";
  return "bar";
}

function isHorizontalChart(type = "") {
  return type === "horizontal_bar";
}

function isStackedChart(type = "") {
  return type === "stacked_bar";
}

function isLineChart(type = "") {
  return type === "line";
}

function shouldShowLegend(series = [], type = "") {
  if (isLineChart(type)) return series.length > 1;
  return series.length > 1;
}

function chartValueFormat(series = []) {
  const firstName = String(series?.[0]?.name || "");

  if (/%|율|비율|증감/.test(firstName)) {
    return "0.0%";
  }

  return "#,##0";
}

function buildChartOptions(
  section = {},
  chartData = {},
  template = PPT_TEMPLATES.default,
) {
  const spec = section.chartSpec || {};
  const type = spec.recommendedType || "";
  const series = chartData.series || [];

  const opts = {
    x: template.contentX,
    y: template.chartY,
    w: template.contentW,
    h: template.chartH,
    showLegend: shouldShowLegend(series, type),
    showTitle: false,
    showValue: template.chartShowValue,
    showCategoryName: template.chartShowCategoryName,
    showCatName: template.chartShowCategoryName,
    catAxisLabelFontFace: template.fontFace,
    catAxisLabelFontSize: template.chartCategoryAxisFontSize,
    valAxisLabelFontFace: template.fontFace,
    valAxisLabelFontSize: template.chartValueAxisFontSize,
    legendFontFace: template.fontFace,
    legendFontSize: template.chartLegendFontSize,
    valAxisNumFmt: chartValueFormat(series),
    showValAxis: true,
    showCatAxis: true,
    showMajorGridLines: template.chartShowGrid,
  };

  if (isHorizontalChart(type)) {
    opts.barDir = "bar";
  }

  if (isStackedChart(type)) {
    opts.grouping = "stacked";
  }

  if (isLineChart(type)) {
    opts.showMarker = true;
    opts.smooth = false;
  }

  return opts;
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

  const labels = rows.map((r) => truncate(r[categoryField] ?? "", 30));

  const series = seriesFields.map((field) => ({
    name: String(field),
    labels,
    values: rows.map((r) => {
      const raw =
        r[field] ?? (r.metric === field ? r.value : undefined) ?? r.value;
      const n = Number(String(raw ?? "").replace(/,/g, ""));
      return Number.isFinite(n) ? n : 0;
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
    addTable(slide, section.rows || [], template);
    return;
  }

  const chartType = mapChartType(spec.recommendedType);

  slide.addChart(
    chartType,
    chartData.series,
    buildChartOptions(section, chartData, template),
  );

  if (section.rows?.length) {
    slide.addText(`차트 데이터: ${section.rows.length.toLocaleString()}건`, {
      x: template.contentX,
      y: 4.6,
      w: template.contentW,
      h: 0.25,
      fontFace: template.fontFace,
      fontSize: 9,
      color: "666666",
    });
  }
}

module.exports = {
  renderReportPpt,
};
