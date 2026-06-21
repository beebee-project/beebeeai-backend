const BUSINESS_TEMPLATE_RESULT_TYPE = "businessTemplate";
const BUSINESS_TEMPLATE_CONTRACT_VERSION = "business_template_result_v1";

const ALLOWED_OUTPUT_TYPES = ["summarySheet", "analysisReport", "ppt"];

const OUTPUT_TYPE_ALIASES = {
  reportJson: "analysisReport",
  reportJSON: "analysisReport",
  report_json: "analysisReport",
  json: "analysisReport",
  analysis: "analysisReport",
  analysisJson: "analysisReport",
  template: "summarySheet",
  workbook: "summarySheet",
  xlsx: "summarySheet",
  powerpoint: "ppt",
  pptx: "ppt",
};

const OUTPUT_TYPE_LABELS = {
  summarySheet: "자동화시트",
  analysisReport: "데이터분석",
  ppt: "PPT",
};

function normalizeOutputType(outputType) {
  const raw = String(outputType || "").trim();
  if (!raw) return null;
  const aliased = OUTPUT_TYPE_ALIASES[raw] || raw;
  return ALLOWED_OUTPUT_TYPES.includes(aliased) ? aliased : null;
}

function normalizeOutputTypes(outputTypes = []) {
  const source = Array.isArray(outputTypes) ? outputTypes : [outputTypes];
  const normalized = source.map(normalizeOutputType).filter(Boolean);
  const deduped = [...new Set(normalized)];
  return deduped.length ? deduped : [...ALLOWED_OUTPUT_TYPES];
}

function outputTypeLabel(outputType) {
  const normalized = normalizeOutputType(outputType);
  return normalized ? OUTPUT_TYPE_LABELS[normalized] : String(outputType || "");
}

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function ensureRows(value) {
  return Array.isArray(value) ? value : [];
}

function inferRowCount(result = {}) {
  const explicit = Number(result.rowCount);
  if (Number.isFinite(explicit) && explicit >= 0) return explicit;
  return ensureRows(result.rows).length;
}

function normalizeSectionResult(result = {}, fallbackColumns = {}) {
  const rows = ensureRows(result.rows);
  const columns = isPlainObject(result.columns)
    ? result.columns
    : isPlainObject(fallbackColumns)
      ? fallbackColumns
      : {};

  return {
    ...result,
    rows,
    rowCount: inferRowCount({ ...result, rows }),
    columns,
  };
}

function normalizeBusinessTemplateSection(section = {}, index = 0) {
  const candidate = isPlainObject(section.candidate) ? section.candidate : {};
  const rawResult = isPlainObject(section.result) ? section.result : {};

  const sectionType =
    section.sectionType ||
    candidate.sectionType ||
    candidate.meta?.sectionType ||
    candidate.recipeType ||
    rawResult.recipeType ||
    rawResult.resultType ||
    "custom_metric";

  const sectionId =
    section.sectionId ||
    candidate.sectionId ||
    candidate.id ||
    candidate.recipeId ||
    `${sectionType}_${index + 1}`;

  const title =
    section.title ||
    candidate.title ||
    candidate.name ||
    rawResult.title ||
    `분석 섹션 ${index + 1}`;

  return {
    sectionId: String(sectionId),
    sectionType: String(sectionType),
    title: String(title),
    result: normalizeSectionResult(
      rawResult,
      candidate.columns || section.columns,
    ),
    chartHint: isPlainObject(section.chartHint)
      ? section.chartHint
      : isPlainObject(candidate.chartHint)
        ? candidate.chartHint
        : {},
    narrativeHint: isPlainObject(section.narrativeHint)
      ? section.narrativeHint
      : isPlainObject(candidate.narrativeHint)
        ? candidate.narrativeHint
        : {},
    candidate,
  };
}

function normalizeBusinessTemplateResult(result = {}, templateCandidate = {}) {
  const sections = ensureRows(result.sections).map((section, index) =>
    normalizeBusinessTemplateSection(section, index),
  );

  return {
    ...result,
    ok: result.ok !== false,
    resultType: BUSINESS_TEMPLATE_RESULT_TYPE,
    templateId: result.templateId || templateCandidate.templateId || "",
    title:
      result.title ||
      templateCandidate.title ||
      result.templateId ||
      templateCandidate.templateId ||
      "업무 템플릿",
    description: result.description || templateCandidate.description || "",
    outputTypes: normalizeOutputTypes(
      result.outputTypes ||
        templateCandidate.outputTypes ||
        ALLOWED_OUTPUT_TYPES,
    ),
    sections,
    contractVersion: BUSINESS_TEMPLATE_CONTRACT_VERSION,
  };
}

function validateBusinessTemplateResultContract(result = {}) {
  const issues = [];

  if (result.resultType !== BUSINESS_TEMPLATE_RESULT_TYPE) {
    issues.push("INVALID_RESULT_TYPE");
  }

  if (!result.templateId) issues.push("MISSING_TEMPLATE_ID");
  if (!Array.isArray(result.sections)) issues.push("MISSING_SECTIONS");

  (result.sections || []).forEach((section, index) => {
    if (!section.sectionId) issues.push(`SECTION_${index}_MISSING_ID`);
    if (!section.sectionType) issues.push(`SECTION_${index}_MISSING_TYPE`);
    if (!section.title) issues.push(`SECTION_${index}_MISSING_TITLE`);
    if (!isPlainObject(section.result))
      issues.push(`SECTION_${index}_MISSING_RESULT`);
    if (!Array.isArray(section.result?.rows))
      issues.push(`SECTION_${index}_ROWS_NOT_ARRAY`);
    if (!Number.isFinite(Number(section.result?.rowCount))) {
      issues.push(`SECTION_${index}_ROWCOUNT_INVALID`);
    }
    if (!isPlainObject(section.result?.columns)) {
      issues.push(`SECTION_${index}_COLUMNS_NOT_OBJECT`);
    }
    if (!isPlainObject(section.chartHint))
      issues.push(`SECTION_${index}_CHARTHINT_NOT_OBJECT`);
    if (!isPlainObject(section.narrativeHint)) {
      issues.push(`SECTION_${index}_NARRATIVEHINT_NOT_OBJECT`);
    }
  });

  return {
    ok: issues.length === 0,
    issues,
  };
}

function isBusinessTemplateResult(result = {}) {
  return (
    result?.resultType === BUSINESS_TEMPLATE_RESULT_TYPE ||
    Array.isArray(result?.sections)
  );
}

module.exports = {
  BUSINESS_TEMPLATE_RESULT_TYPE,
  BUSINESS_TEMPLATE_CONTRACT_VERSION,
  ALLOWED_OUTPUT_TYPES,
  OUTPUT_TYPE_LABELS,
  normalizeOutputType,
  normalizeOutputTypes,
  outputTypeLabel,
  normalizeSectionResult,
  normalizeBusinessTemplateSection,
  normalizeBusinessTemplateResult,
  validateBusinessTemplateResultContract,
  isBusinessTemplateResult,
};
