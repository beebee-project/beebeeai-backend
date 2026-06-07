const pptxgen = require("pptxgenjs");

function text(v) {
  return String(v ?? "");
}

const PPT_TEMPLATES = {
  default: {
    name: "default",
    fontFace: "Arial",
    titleFontSize: 22,
    bodyFontSize: 14,
    tableFontSize: 9,
    coverTitleFontSize: 28,
    titleX: 0.5,
    titleY: 0.4,
    titleW: 9,
    titleH: 0.5,
    contentX: 0.7,
    contentY: 1.2,
    contentW: 8.6,
    chartY: 1.25,
    chartH: 3.4,
    tableY: 1.1,
    tableH: 4.2,
    maxTableRowsPerSlide: 8,
    maxTableCols: 6,
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
    fontFace: "Arial",
    titleFontSize: 20,
    bodyFontSize: 13,
    tableFontSize: 8,
    coverTitleFontSize: 26,
    titleX: 0.6,
    titleY: 0.45,
    titleW: 8.8,
    titleH: 0.5,
    contentX: 0.75,
    contentY: 1.2,
    contentW: 8.4,
    chartY: 1.25,
    chartH: 3.3,
    tableY: 1.1,
    tableH: 4.1,
    maxTableRowsPerSlide: 9,
    maxTableCols: 6,
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

function addTitle(slide, title, y = null, template = PPT_TEMPLATES.default) {
  slide.addText(text(title), {
    x: template.titleX,
    y: y ?? template.titleY,
    w: template.titleW,
    h: template.titleH,
    fontFace: template.fontFace,
    fontSize: template.titleFontSize,
    bold: true,
  });
}

function addBullets(
  slide,
  bullets = [],
  y = null,
  template = PPT_TEMPLATES.default,
) {
  const safe = Array.isArray(bullets) ? bullets.slice(0, 6) : [];

  slide.addText(
    safe.map((b) => `• ${text(b)}`).join("\n") || "요약 내용이 없습니다.",
    {
      x: template.contentX,
      y: y ?? template.contentY,
      w: template.contentW,
      h: 3.8,
      fontFace: template.fontFace,
      fontSize: template.bodyFontSize,
      breakLine: false,
      fit: "shrink",
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

  const body = rows.map((row) => keys.map((h) => text(row[h])));
  return [keys, ...body];
}

function rowsToTable(rows = [], limit = 8) {
  const safeRows = Array.isArray(rows) ? rows.slice(0, limit) : [];
  const keys = tableColumnKeys(safeRows);

  return rowsToTableWithKeys(safeRows, keys);
}

function addTable(slide, rows = [], template = PPT_TEMPLATES.default) {
  slide.addTable(rowsToTable(rows, template.maxTableRowsPerSlide), {
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
) {
  const rows = Array.isArray(section.rows) ? section.rows : [];
  const keys = tableColumnKeys(rows, template.maxTableCols);
  const chunks = chunkRows(rows, template.maxTableRowsPerSlide);
  const totalPages = chunks.length;

  chunks.forEach((chunk, idx) => {
    const slide = pptx.addSlide();
    const title =
      totalPages > 1
        ? `${section.title || "분석 결과"} (${idx + 1}/${totalPages})`
        : section.title || "분석 결과";

    addTitle(slide, title, null, template);
    slide.addTable(rowsToTableWithKeys(chunk, keys), {
      x: 0.5,
      y: template.tableY,
      w: 9,
      h: template.tableH,
      fontFace: template.fontFace,
      fontSize: template.tableFontSize,
      border: { type: "solid", pt: 0.5 },
      fit: "shrink",
    });

    slide.addText(`전체 행 수: ${rows.length}`, {
      x: 0.5,
      y: 5.35,
      w: 4,
      h: 0.3,
      fontFace: template.fontFace,
      fontSize: 9,
    });
  });
}

function renderReportPpt(report = {}, options = {}) {
  const template = getTemplate(options);
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "BeeBeeAI";

  const sections = Array.isArray(report.sections) ? report.sections : [];

  for (const section of sections) {
    if (section.type === "table") {
      addPagedTableSlides(pptx, section, template);
      continue;
    }

    const slide = pptx.addSlide();

    if (section.type === "cover") {
      slide.addText(text(section.title || report.title || "분석 보고서"), {
        x: template.contentX,
        y: 1.6,
        w: template.contentW,
        h: 0.8,
        fontFace: template.fontFace,
        fontSize: template.coverTitleFontSize,
        bold: true,
      });

      slide.addText(text(section.subtitle || ""), {
        x: template.contentX,
        y: 2.6,
        w: template.contentW,
        h: 0.4,
        fontFace: template.fontFace,
        fontSize: template.bodyFontSize,
      });

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
      continue;
    }

    if (section.type === "chart") {
      addTitle(slide, section.title || "차트", null, template);

      slide.addText(
        `차트 유형: ${section.chartSpec?.recommendedType || "chart"}`,
        {
          x: template.contentX,
          y: 0.95,
          w: template.contentW,
          h: 0.3,
          fontFace: template.fontFace,
          fontSize: 11,
        },
      );

      addChart(slide, section, template);
      continue;
    }

    if (section.type === "insight") {
      addTitle(slide, section.title || "분석 인사이트", null, template);
      addBullets(slide, section.bullets || [], null, template);
      continue;
    }

    addTitle(slide, section.title || section.type || "섹션", null, template);
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
    y: template.chartY + template.chartH + 0.2,
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

  const labels = rows.map((r) => String(r[categoryField] ?? ""));

  const series = seriesFields.map((field) => ({
    name: String(field),
    labels,
    values: rows.map((r) => {
      const raw =
        r[field] ?? (r.metric === field ? r.value : undefined) ?? r.value;

      const n = Number(raw);
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
    slide.addText(`데이터 행 수: ${section.rows.length}`, {
      x: template.contentX,
      y: 4.85,
      w: 8,
      h: 0.3,
      fontFace: template.fontFace,
      fontSize: 10,
    });
  }
}

module.exports = {
  renderReportPpt,
};
