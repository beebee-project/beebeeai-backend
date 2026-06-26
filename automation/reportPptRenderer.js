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
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "BeeBeeAI";
  pptx.subject = report.title || "분석 보고서";
  pptx.title = report.title || "분석 보고서";
  pptx.company = "BeeBee AI";

  const sections = Array.isArray(report.sections) ? report.sections : [];
  const pageState = { page: 0 };

  for (const section of sections) {
    if (section.type === "table") {
      addPagedTableSlides(pptx, section, template, report, pageState);
      continue;
    }

    const slide = pptx.addSlide();
    pageState.page += 1;

    if (section.type === "cover") {
      slide.addText(text(section.title || report.title || "분석 보고서"), {
        x: template.contentX,
        y: 1.45,
        w: template.contentW,
        h: 0.8,
        fontFace: template.fontFace,
        fontSize: template.coverTitleFontSize,
        bold: true,
        fit: "shrink",
      });

      slide.addText(text(section.subtitle || report.source?.fileName || ""), {
        x: template.contentX,
        y: 2.45,
        w: template.contentW,
        h: 0.34,
        fontFace: template.fontFace,
        fontSize: template.bodyFontSize,
        color: "555555",
      });

      if (report.source?.message) {
        slide.addText(cleanDisplayTitle(report.source.message), {
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
        `생성일시: ${formatDateTime(section.generatedAt || report.generatedAt)}`,
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
      addFooter(slide, report, pageState.page, template);
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
      addFooter(slide, report, pageState.page, template);
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
      addFooter(slide, report, pageState.page, template);
      continue;
    }

    if (section.type === "insight") {
      addTitle(slide, section.title || "분석 인사이트", null, template);
      addBullets(slide, section.bullets || [], null, template);
      addFooter(slide, report, pageState.page, template);
      continue;
    }

    addTitle(slide, section.title || section.type || "섹션", null, template);
    addFooter(slide, report, pageState.page, template);
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
    addFooter(slide, report, pageState.page, template);
  }

  return pptx;
}

module.exports = {
  renderReportPpt,
};
