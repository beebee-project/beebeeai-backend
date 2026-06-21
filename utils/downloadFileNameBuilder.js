const path = require("path");

const OUTPUT_TYPE_LABELS = {
  summarySheet: "자동화시트",
  analysisReport: "데이터분석",
  ppt: "PPT",
};

const OUTPUT_TYPE_ALIASES = {
  reportJson: "analysisReport",
  reportJSON: "analysisReport",
  report_json: "analysisReport",
  json: "analysisReport",
  pptx: "ppt",
  xlsx: "summarySheet",
};

function stripExtension(fileName = "") {
  const base = path.basename(String(fileName || ""));
  const ext = path.extname(base);
  return ext ? base.slice(0, -ext.length) : base;
}

function sanitizePart(value = "", fallback = "파일") {
  const cleaned = String(value || "")
    .replace(/[^\p{Letter}\p{Number}\s-]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();

  return (cleaned || fallback).slice(0, 60).trim();
}

function formatKstTimestamp(date = new Date()) {
  const kst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  const y = kst.getUTCFullYear();
  const m = String(kst.getUTCMonth() + 1).padStart(2, "0");
  const d = String(kst.getUTCDate()).padStart(2, "0");
  const hh = String(kst.getUTCHours()).padStart(2, "0");
  const mm = String(kst.getUTCMinutes()).padStart(2, "0");
  return `${y}${m}${d}_${hh}${mm}`;
}

function normalizeExtension(extension = "") {
  const ext = String(extension || "").trim().replace(/^\./, "").toLowerCase();
  return ext || "xlsx";
}

function buildDownloadFileName({
  sourceFileName,
  templateTitle,
  outputType,
  extension,
  date = new Date(),
}) {
  const source = sanitizePart(stripExtension(sourceFileName), "원본파일");
  const template = sanitizePart(templateTitle, "보고서");
  const normalizedOutputType = OUTPUT_TYPE_ALIASES[outputType] || outputType;
  const output =
    OUTPUT_TYPE_LABELS[normalizedOutputType] ||
    sanitizePart(normalizedOutputType, "결과");
  const timestamp = formatKstTimestamp(date);
  const ext = normalizeExtension(extension);

  return `${source}_${template}_${output}_${timestamp}.${ext}`;
}

module.exports = {
  OUTPUT_TYPE_LABELS,
  buildDownloadFileName,
};
