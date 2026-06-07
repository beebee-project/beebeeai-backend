const pptxgen = require("pptxgenjs");

function text(v) {
  return String(v ?? "");
}

function addTitle(slide, title, y = 0.4) {
  slide.addText(text(title), {
    x: 0.5,
    y,
    w: 9,
    h: 0.5,
    fontSize: 22,
    bold: true,
  });
}

function addBullets(slide, bullets = [], y = 1.2) {
  const safe = Array.isArray(bullets) ? bullets.slice(0, 6) : [];

  slide.addText(
    safe.map((b) => `• ${text(b)}`).join("\n") || "요약 내용이 없습니다.",
    {
      x: 0.7,
      y,
      w: 8.6,
      h: 3.8,
      fontSize: 14,
      breakLine: false,
      fit: "shrink",
    },
  );
}

function rowsToTable(rows = [], limit = 8) {
  const safeRows = Array.isArray(rows) ? rows.slice(0, limit) : [];
  if (!safeRows.length) return [["결과 없음"]];

  const headers = Object.keys(safeRows[0] || {}).slice(0, 6);
  const body = safeRows.map((row) => headers.map((h) => text(row[h])));

  return [headers, ...body];
}

function addTable(slide, rows = []) {
  slide.addTable(rowsToTable(rows), {
    x: 0.5,
    y: 1.1,
    w: 9,
    h: 4.2,
    fontSize: 9,
    border: { type: "solid", pt: 0.5 },
    fit: "shrink",
  });
}

function renderReportPpt(report = {}) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "BeeBeeAI";

  const sections = Array.isArray(report.sections) ? report.sections : [];

  for (const section of sections) {
    const slide = pptx.addSlide();

    if (section.type === "cover") {
      slide.addText(text(section.title || report.title || "분석 보고서"), {
        x: 0.7,
        y: 1.6,
        w: 8.8,
        h: 0.8,
        fontSize: 28,
        bold: true,
      });

      slide.addText(text(section.subtitle || ""), {
        x: 0.7,
        y: 2.6,
        w: 8.8,
        h: 0.4,
        fontSize: 14,
      });

      continue;
    }

    if (section.type === "summary") {
      addTitle(slide, section.title || "핵심 요약");
      addBullets(
        slide,
        [section.summary, ...(section.bullets || [])].filter(Boolean),
      );
      continue;
    }

    if (section.type === "chart") {
      addTitle(slide, section.title || "차트");
      slide.addText(
        `추천 차트: ${section.chartSpec?.recommendedType || "chart"}`,
        { x: 0.7, y: 1.0, w: 8.5, h: 0.3, fontSize: 12 },
      );
      addTable(slide, section.rows || []);
      continue;
    }

    if (section.type === "table") {
      addTitle(slide, section.title || "분석 결과");
      addTable(slide, section.rows || []);
      continue;
    }

    if (section.type === "insight") {
      addTitle(slide, section.title || "분석 인사이트");
      addBullets(slide, section.bullets || []);
      continue;
    }

    addTitle(slide, section.title || section.type || "섹션");
  }

  if (!sections.length) {
    const slide = pptx.addSlide();
    addTitle(slide, "분석 보고서");
    slide.addText("보고서 섹션이 없습니다.", {
      x: 0.7,
      y: 1.2,
      w: 8,
      h: 0.5,
      fontSize: 14,
    });
  }

  return pptx;
}

module.exports = {
  renderReportPpt,
};
