const SHEET_NAMES = Object.freeze({
  SUMMARY: "요약",
  SOURCE_DATA: "원본데이터",
  ANALYSIS_RESULT: "분석결과",
  EXECUTION_PLAN: "실행계획",
  EXECUTION_META: "실행메타",
  CHART_DATA: "차트데이터",
  CHART_CONFIG: "차트설정",
  INSIGHTS: "인사이트",
  REPORT_SECTIONS: "보고서구성",

  AUTOMATION_GUIDE: "사용방법",
  AUTOMATION_SETTINGS: "자동화설정",
  AUTOMATION_TEMPLATE: "자동화시트",
  SUMMARY_ROWS: "요약행",
  DIAGNOSTICS: "진단정보",
  EXECUTION_PREVIEW: "실행결과_미리보기",
});

const SUMMARY_SHEET_MODES = Object.freeze([
  "static",
  "formula",
  "hybrid",
  "sourceDataOnly",
]);
const SUMMARY_SHEET_MODE_SET = new Set(SUMMARY_SHEET_MODES);

function normalizeSummarySheetMode(mode = "static") {
  const normalized = String(mode || "static").trim();
  return SUMMARY_SHEET_MODE_SET.has(normalized) ? normalized : "static";
}

function isFormulaEnabledMode(mode = "static") {
  return mode === "formula" || mode === "hybrid";
}

function sourceSheetNameForTableIndex(index = 0, total = 1) {
  const safeIndex = Math.max(0, Number(index) || 0);
  const safeTotal = Math.max(1, Number(total) || 1);

  if (safeTotal <= 1 || safeIndex === 0) return SHEET_NAMES.SOURCE_DATA;

  return `${SHEET_NAMES.SOURCE_DATA}${safeIndex + 1}`;
}

module.exports = {
  SHEET_NAMES,
  SUMMARY_SHEET_MODES,
  normalizeSummarySheetMode,
  isFormulaEnabledMode,
  sourceSheetNameForTableIndex,
};
