const OUTPUT_TYPES = Object.freeze({
  SUMMARY_SHEET: "summarySheet",
  ANALYSIS_REPORT: "analysisReport",
  PPT: "ppt",
});

const OUTPUT_ARTIFACTS = Object.freeze({
  [OUTPUT_TYPES.SUMMARY_SHEET]: Object.freeze({
    outputType: OUTPUT_TYPES.SUMMARY_SHEET,
    label: "자동화시트",
    uiLabel: "자동화 시트",
    extension: "xlsx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    defaultTitle: "자동화 시트",
    storagePrefix: "generated/summary-sheets",
    localDirName: "summary-sheets",
  }),

  [OUTPUT_TYPES.ANALYSIS_REPORT]: Object.freeze({
    outputType: OUTPUT_TYPES.ANALYSIS_REPORT,
    label: "데이터분석",
    uiLabel: "데이터 분석",
    extension: "json",
    mimeType: "application/json; charset=utf-8",
    defaultTitle: "데이터 분석",
    storagePrefix: "reports",
    localDirName: "reports",
    version: "analysis_report_export_v1",
  }),

  [OUTPUT_TYPES.PPT]: Object.freeze({
    outputType: OUTPUT_TYPES.PPT,
    label: "PPT",
    uiLabel: "PPT 생성",
    extension: "pptx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    defaultTitle: "PPT",
    storagePrefix: "ppt",
    localDirName: "ppt",
  }),
});

const OUTPUT_TYPE_ALIASES = Object.freeze({
  reportJson: OUTPUT_TYPES.ANALYSIS_REPORT,
  reportJSON: OUTPUT_TYPES.ANALYSIS_REPORT,
  report_json: OUTPUT_TYPES.ANALYSIS_REPORT,
  json: OUTPUT_TYPES.ANALYSIS_REPORT,
  analysis: OUTPUT_TYPES.ANALYSIS_REPORT,
  analysisJson: OUTPUT_TYPES.ANALYSIS_REPORT,

  template: OUTPUT_TYPES.SUMMARY_SHEET,
  workbook: OUTPUT_TYPES.SUMMARY_SHEET,
  xlsx: OUTPUT_TYPES.SUMMARY_SHEET,

  powerpoint: OUTPUT_TYPES.PPT,
  pptx: OUTPUT_TYPES.PPT,
});

const ALLOWED_OUTPUT_TYPES = Object.freeze(Object.keys(OUTPUT_ARTIFACTS));

function normalizeOutputType(outputType) {
  const raw = String(outputType || "").trim();
  if (!raw) return null;

  const aliased = OUTPUT_TYPE_ALIASES[raw] || raw;

  return Object.prototype.hasOwnProperty.call(OUTPUT_ARTIFACTS, aliased)
    ? aliased
    : null;
}

function normalizeOutputTypes(outputTypes = []) {
  const source = Array.isArray(outputTypes) ? outputTypes : [outputTypes];
  const normalized = source.map(normalizeOutputType).filter(Boolean);
  const deduped = [...new Set(normalized)];

  return deduped.length ? deduped : [...ALLOWED_OUTPUT_TYPES];
}

function getOutputArtifact(outputType) {
  const normalized = normalizeOutputType(outputType);
  return normalized ? OUTPUT_ARTIFACTS[normalized] : null;
}

function getOutputTypeLabel(outputType) {
  return getOutputArtifact(outputType)?.label || String(outputType || "");
}

function getOutputTypeUiLabel(outputType) {
  return getOutputArtifact(outputType)?.uiLabel || String(outputType || "");
}

function getOutputExtension(outputType, fallback = "xlsx") {
  return getOutputArtifact(outputType)?.extension || fallback;
}

function getOutputMimeType(outputType, fallback = "application/octet-stream") {
  return getOutputArtifact(outputType)?.mimeType || fallback;
}

function getOutputDefaultTitle(outputType, fallback = "보고서") {
  return getOutputArtifact(outputType)?.defaultTitle || fallback;
}

function getOutputVersion(outputType) {
  return getOutputArtifact(outputType)?.version || "";
}

function getOutputArtifactByExtension(extension = "") {
  const ext = String(extension || "")
    .replace(/^\./, "")
    .toLowerCase();

  return (
    Object.values(OUTPUT_ARTIFACTS).find(
      (artifact) => artifact.extension === ext,
    ) || null
  );
}

function inferOutputArtifact({ outputType = "", fileName = "" } = {}) {
  const byType = getOutputArtifact(outputType);
  if (byType) return byType;

  const ext = String(fileName || "")
    .split(".")
    .pop();
  return getOutputArtifactByExtension(ext);
}

module.exports = {
  OUTPUT_TYPES,
  OUTPUT_ARTIFACTS,
  OUTPUT_TYPE_ALIASES,
  ALLOWED_OUTPUT_TYPES,
  normalizeOutputType,
  normalizeOutputTypes,
  getOutputArtifact,
  getOutputTypeLabel,
  getOutputTypeUiLabel,
  getOutputExtension,
  getOutputMimeType,
  getOutputDefaultTitle,
  getOutputVersion,
  getOutputArtifactByExtension,
  inferOutputArtifact,
};
