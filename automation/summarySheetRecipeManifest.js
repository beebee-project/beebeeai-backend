let xlsxModule = null;

function getXlsxModule() {
  if (xlsxModule) return xlsxModule;
  // Lazy load keeps non-workbook smoke tests independent from xlsx installation.
  // Production workbook generation still requires the normal project dependency.
  xlsxModule = require("xlsx");
  return xlsxModule;
}

const { collectSectionMetricIds } = require("./metricIdContract");

const SUMMARY_SHEET_RECIPE_MANIFEST_VERSION =
  "summary_sheet_recipe_manifest_v2";
const SUMMARY_SHEET_RECIPE_MANIFEST_SHEET = "_beebee_recipe_manifest";
const FINAL_OUTPUT_QUALITY_GATE_MANIFEST_SERIALIZATION_VERSION =
  "final_output_quality_gate_manifest_serialization_v1";

const FINAL_OUTPUT_QUALITY_GATE_META_KEYS = Object.freeze([
  "outputCompletenessContractVersion",
  "semanticOutputDuplicateResolverVersion",
  "finalOutputQualityGateVersion",
  "finalOutputQualityGateApplied",
  "finalOutputQualityGatePass",
  "finalOutputQualityGateStatus",
  "finalOutputQualityGateStrict",
  "finalOutputQualityGateFailureReasons",
  "finalOutputQualityGateExpectedMetricCount",
  "finalOutputQualityGateRenderedExpectedMetricCount",
  "finalOutputQualityGateMissingMetricIds",
  "finalOutputQualityGateDuplicateMetricIds",
  "finalOutputQualityGateEmptyRequiredSectionIds",
  "finalOutputQualityGateIncompleteRequiredSectionIds",
  "finalOutputQualityGateInvalidNumberCount",
  "finalOutputQualityGateBudgetViolationCount",
  "finalOutputQualityGateInputSectionCount",
  "finalOutputQualityGateOutputSectionCount",
  "finalOutputQualityGateRemovedDuplicateSectionCount",
  "finalOutputQualityGateRemovedDuplicateSectionIds",
  "finalOutputQualityGateMergedMetricIdCount",
  "finalOutputQualityGateReassignedMetricIdCount",
  "finalOutputQualityGateRenamedSectionIdCount",
  "finalOutputQualityGateRenamedTitleCount",
]);

const FINAL_OUTPUT_QUALITY_GATE_BOOLEAN_KEYS = new Set([
  "finalOutputQualityGateApplied",
  "finalOutputQualityGatePass",
  "finalOutputQualityGateStrict",
]);

const FINAL_OUTPUT_QUALITY_GATE_NUMBER_KEYS = new Set([
  "finalOutputQualityGateExpectedMetricCount",
  "finalOutputQualityGateRenderedExpectedMetricCount",
  "finalOutputQualityGateInvalidNumberCount",
  "finalOutputQualityGateBudgetViolationCount",
  "finalOutputQualityGateInputSectionCount",
  "finalOutputQualityGateOutputSectionCount",
  "finalOutputQualityGateRemovedDuplicateSectionCount",
  "finalOutputQualityGateMergedMetricIdCount",
  "finalOutputQualityGateReassignedMetricIdCount",
  "finalOutputQualityGateRenamedSectionIdCount",
  "finalOutputQualityGateRenamedTitleCount",
]);

const FINAL_OUTPUT_QUALITY_GATE_JSON_KEYS = new Set([
  "finalOutputQualityGateFailureReasons",
  "finalOutputQualityGateMissingMetricIds",
  "finalOutputQualityGateDuplicateMetricIds",
  "finalOutputQualityGateEmptyRequiredSectionIds",
  "finalOutputQualityGateIncompleteRequiredSectionIds",
  "finalOutputQualityGateRemovedDuplicateSectionIds",
]);

function normalizeText(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .trim();
}

function cloneJsonValue(value) {
  if (value == null) return value;
  if (Array.isArray(value)) {
    return value.map((item) => cloneJsonValue(item));
  }
  if (typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, cloneJsonValue(item)]),
    );
  }
  return value;
}

function hasOwn(target, key) {
  return Object.prototype.hasOwnProperty.call(target || {}, key);
}

function buildFinalOutputQualityGateMeta(result = {}) {
  const executionMeta =
    result.executionMeta && typeof result.executionMeta === "object"
      ? result.executionMeta
      : {};
  const gate =
    result.finalOutputQualityGate &&
    typeof result.finalOutputQualityGate === "object"
      ? result.finalOutputQualityGate
      : {};

  const meta = {};
  const copyExecutionMeta = (key) => {
    if (!hasOwn(executionMeta, key)) return;
    meta[key] = cloneJsonValue(executionMeta[key]);
  };

  FINAL_OUTPUT_QUALITY_GATE_META_KEYS.forEach(copyExecutionMeta);
  Object.keys(executionMeta)
    .filter(
      (key) => key.startsWith("finalOutputQualityGate") && !hasOwn(meta, key),
    )
    .sort()
    .forEach(copyExecutionMeta);

  const assignFallback = (key, value) => {
    if (hasOwn(meta, key) || value === undefined) return;
    meta[key] = cloneJsonValue(value);
  };

  assignFallback("outputCompletenessContractVersion", gate.completenessVersion);
  assignFallback(
    "semanticOutputDuplicateResolverVersion",
    gate.duplicateResolverVersion,
  );
  assignFallback("finalOutputQualityGateVersion", gate.version);
  assignFallback(
    "finalOutputQualityGateApplied",
    Object.keys(gate).length > 0 ? true : undefined,
  );
  assignFallback("finalOutputQualityGatePass", gate.pass);
  assignFallback("finalOutputQualityGateStatus", gate.status);
  assignFallback("finalOutputQualityGateFailureReasons", gate.failureReasons);
  assignFallback(
    "finalOutputQualityGateExpectedMetricCount",
    gate.expectedMetricCount,
  );
  assignFallback(
    "finalOutputQualityGateRenderedExpectedMetricCount",
    gate.renderedExpectedMetricCount,
  );

  return meta;
}

function serializeManifestMetaValue(value) {
  if (value == null) return "";
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (Array.isArray(value) || typeof value === "object") {
    return JSON.stringify(value);
  }
  return value;
}

function parseManifestBoolean(value) {
  const normalized = String(value == null ? "" : value)
    .trim()
    .toUpperCase();
  if (normalized === "TRUE") return true;
  if (normalized === "FALSE") return false;
  return Boolean(value);
}

function parseManifestJson(value, fallback = []) {
  try {
    return JSON.parse(String(value == null ? "" : value));
  } catch (_error) {
    return cloneJsonValue(fallback);
  }
}

function parseFinalOutputQualityGateMeta(meta = {}) {
  const keys = Object.keys(meta)
    .filter(
      (key) =>
        FINAL_OUTPUT_QUALITY_GATE_META_KEYS.includes(key) ||
        key.startsWith("finalOutputQualityGate"),
    )
    .sort((left, right) => {
      const leftIndex = FINAL_OUTPUT_QUALITY_GATE_META_KEYS.indexOf(left);
      const rightIndex = FINAL_OUTPUT_QUALITY_GATE_META_KEYS.indexOf(right);
      if (leftIndex >= 0 && rightIndex >= 0) {
        return leftIndex - rightIndex;
      }
      if (leftIndex >= 0) return -1;
      if (rightIndex >= 0) return 1;
      return left.localeCompare(right);
    });

  return Object.fromEntries(
    keys.map((key) => {
      const value = meta[key];
      if (FINAL_OUTPUT_QUALITY_GATE_BOOLEAN_KEYS.has(key)) {
        return [key, parseManifestBoolean(value)];
      }
      if (FINAL_OUTPUT_QUALITY_GATE_NUMBER_KEYS.has(key)) {
        return [key, Number(value || 0)];
      }
      if (FINAL_OUTPUT_QUALITY_GATE_JSON_KEYS.has(key)) {
        return [key, parseManifestJson(value, [])];
      }
      return [key, String(value == null ? "" : value)];
    }),
  );
}

function resultRows(result = {}) {
  return Array.isArray(result.rows) ? result.rows : [];
}

function resultOutputHeaders(result = {}) {
  const rows = resultRows(result);
  return [
    ...new Set(
      rows
        .flatMap((row) => Object.keys(row || {}))
        .map(normalizeText)
        .filter(Boolean),
    ),
  ];
}

function buildSectionManifest(section = {}, resolvedSheetName = "", index = 0) {
  const result = section.result || {};
  return {
    sectionIndex: index,
    sectionId: normalizeText(section.sectionId || ""),
    title: normalizeText(
      section.title || section.sectionId || `section_${index + 1}`,
    ),
    resolvedSheetName: normalizeText(
      resolvedSheetName || section.title || section.sectionId || "",
    ),
    sectionType: normalizeText(section.sectionType || ""),
    resultType: normalizeText(result.resultType || ""),
    operation: normalizeText(result.operation || ""),
    metricHeader: normalizeText(result.metric?.header || ""),
    groupHeader: normalizeText(result.groupBy?.header || ""),
    rowCount: resultRows(result).length,
    outputHeaders: resultOutputHeaders(result),
    sourceMetricHeaders: Array.isArray(result.metrics)
      ? result.metrics
          .map((item) => normalizeText(item?.header || item?.label || item))
          .filter(Boolean)
      : [],
    metricIds: collectSectionMetricIds(section),
    contractCoverageVersion: normalizeText(
      result.meta?.contractCoverageVersion ||
        section.contractCoverageVersion ||
        "",
    ),
    complete: result.meta?.complete !== false,
  };
}

function buildSummarySheetRecipeManifest({
  result = {},
  businessSections = [],
  resolvedSections = [],
} = {}) {
  const resolvedByIndex = new Map(
    (resolvedSections || []).map((entry) => [Number(entry.index), entry]),
  );

  const sections = (businessSections || []).map((section, index) =>
    buildSectionManifest(
      section,
      resolvedByIndex.get(index)?.resolvedSheetName || "",
      index,
    ),
  );

  const renderedMetricIds = Array.from(
    new Set(
      sections.flatMap((section) => section.metricIds || []).filter(Boolean),
    ),
  );

  const coverage = result.contractSummaryCoverage || {};
  const expectedMetricIds = Array.from(
    new Set(
      (coverage.expectedMetricIds || renderedMetricIds)
        .map(normalizeText)
        .filter(Boolean),
    ),
  );

  const missingMetricIds = expectedMetricIds.filter(
    (metricId) => !renderedMetricIds.includes(metricId),
  );
  const finalOutputQualityGateMeta = buildFinalOutputQualityGateMeta(result);

  return {
    version: SUMMARY_SHEET_RECIPE_MANIFEST_VERSION,
    complete: true,
    templateId: normalizeText(result.templateId || ""),
    templateTitle: normalizeText(result.title || ""),
    resultType: normalizeText(result.resultType || "businessTemplate"),
    generatedAt: new Date().toISOString(),
    sectionCount: sections.length,
    renderedMetricIds,
    expectedMetricIds,
    missingMetricIds,
    metricCoveragePass: missingMetricIds.length === 0,
    metricCoverageRate: expectedMetricIds.length
      ? renderedMetricIds.filter((metricId) =>
          expectedMetricIds.includes(metricId),
        ).length / expectedMetricIds.length
      : 1,
    contractCoverageVersion: normalizeText(coverage.version || ""),
    contractCatalogVersion: normalizeText(
      coverage.contractCatalogVersion || "",
    ),
    finalOutputQualityGateMetaVersion:
      FINAL_OUTPUT_QUALITY_GATE_MANIFEST_SERIALIZATION_VERSION,
    finalOutputQualityGateMetaPresent:
      Object.keys(finalOutputQualityGateMeta).length > 0,
    finalOutputQualityGateMeta,
    sections,
  };
}

function manifestToAoa(manifest = {}) {
  return [
    ["key", "value"],
    ["manifestVersion", manifest.version || ""],
    ["complete", manifest.complete === true ? "TRUE" : "FALSE"],
    ["templateId", manifest.templateId || ""],
    ["templateTitle", manifest.templateTitle || ""],
    ["resultType", manifest.resultType || ""],
    ["generatedAt", manifest.generatedAt || ""],
    ["sectionCount", manifest.sectionCount || 0],
    ["renderedMetricIdsJson", JSON.stringify(manifest.renderedMetricIds || [])],
    ["expectedMetricIdsJson", JSON.stringify(manifest.expectedMetricIds || [])],
    ["missingMetricIdsJson", JSON.stringify(manifest.missingMetricIds || [])],
    [
      "metricCoveragePass",
      manifest.metricCoveragePass === true ? "TRUE" : "FALSE",
    ],
    ["metricCoverageRate", manifest.metricCoverageRate ?? 0],
    ["contractCoverageVersion", manifest.contractCoverageVersion || ""],
    ["contractCatalogVersion", manifest.contractCatalogVersion || ""],
    [
      "finalOutputQualityGateMetaVersion",
      manifest.finalOutputQualityGateMetaVersion || "",
    ],
    [
      "finalOutputQualityGateMetaPresent",
      manifest.finalOutputQualityGateMetaPresent === true ? "TRUE" : "FALSE",
    ],
    ...FINAL_OUTPUT_QUALITY_GATE_META_KEYS.filter((key) =>
      hasOwn(manifest.finalOutputQualityGateMeta, key),
    ).map((key) => [
      key,
      serializeManifestMetaValue(manifest.finalOutputQualityGateMeta[key]),
    ]),
    ...Object.keys(manifest.finalOutputQualityGateMeta || {})
      .filter(
        (key) =>
          key.startsWith("finalOutputQualityGate") &&
          !FINAL_OUTPUT_QUALITY_GATE_META_KEYS.includes(key),
      )
      .sort()
      .map((key) => [
        key,
        serializeManifestMetaValue(manifest.finalOutputQualityGateMeta[key]),
      ]),
    [],
    [
      "sectionIndex",
      "sectionId",
      "title",
      "resolvedSheetName",
      "sectionType",
      "resultType",
      "operation",
      "metricHeader",
      "groupHeader",
      "rowCount",
      "outputHeadersJson",
      "sourceMetricHeadersJson",
      "metricIdsJson",
      "contractCoverageVersion",
      "complete",
    ],
    ...(manifest.sections || []).map((section) => [
      section.sectionIndex,
      section.sectionId,
      section.title,
      section.resolvedSheetName,
      section.sectionType,
      section.resultType,
      section.operation,
      section.metricHeader,
      section.groupHeader,
      section.rowCount,
      JSON.stringify(section.outputHeaders || []),
      JSON.stringify(section.sourceMetricHeaders || []),
      JSON.stringify(section.metricIds || []),
      section.contractCoverageVersion || "",
      section.complete === false ? "FALSE" : "TRUE",
    ]),
  ];
}

function ensureHiddenSheet(wb, sheetName) {
  wb.Workbook = wb.Workbook || {};
  wb.Workbook.Sheets = Array.isArray(wb.Workbook.Sheets)
    ? wb.Workbook.Sheets
    : [];

  const index = (wb.SheetNames || []).indexOf(sheetName);
  if (index < 0) return;

  while (wb.Workbook.Sheets.length <= index) {
    wb.Workbook.Sheets.push({});
  }

  wb.Workbook.Sheets[index] = {
    ...(wb.Workbook.Sheets[index] || {}),
    name: sheetName,
    Hidden: 2,
  };
}

function appendSummarySheetRecipeManifest(wb, args = {}) {
  const manifest = buildSummarySheetRecipeManifest(args);
  const XLSX = getXlsxModule();
  const ws = XLSX.utils.aoa_to_sheet(manifestToAoa(manifest));

  XLSX.utils.book_append_sheet(wb, ws, SUMMARY_SHEET_RECIPE_MANIFEST_SHEET);
  ensureHiddenSheet(wb, SUMMARY_SHEET_RECIPE_MANIFEST_SHEET);

  wb["!beebeeSummaryRecipeManifest"] = manifest;
  return manifest;
}

function parseJsonArray(value) {
  try {
    const parsed = JSON.parse(String(value || "[]"));
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function readSummarySheetRecipeManifest(workbook = {}) {
  if (workbook["!beebeeSummaryRecipeManifest"]?.version) {
    return workbook["!beebeeSummaryRecipeManifest"];
  }

  const ws = workbook.Sheets?.[SUMMARY_SHEET_RECIPE_MANIFEST_SHEET];
  if (!ws) return null;

  const XLSX = getXlsxModule();
  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    raw: false,
    defval: "",
  });

  const meta = {};
  let sectionHeaderIndex = -1;

  for (let index = 1; index < rows.length; index += 1) {
    const row = rows[index] || [];
    if (String(row[0] || "") === "sectionIndex") {
      sectionHeaderIndex = index;
      break;
    }
    if (row[0]) meta[String(row[0])] = row[1];
  }

  const sections = [];
  if (sectionHeaderIndex >= 0) {
    for (let index = sectionHeaderIndex + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      if (row.every((value) => value === "")) continue;

      sections.push({
        sectionIndex: Number(row[0] || 0),
        sectionId: String(row[1] || ""),
        title: String(row[2] || ""),
        resolvedSheetName: String(row[3] || ""),
        sectionType: String(row[4] || ""),
        resultType: String(row[5] || ""),
        operation: String(row[6] || ""),
        metricHeader: String(row[7] || ""),
        groupHeader: String(row[8] || ""),
        rowCount: Number(row[9] || 0),
        outputHeaders: parseJsonArray(row[10]),
        sourceMetricHeaders: parseJsonArray(row[11]),
        metricIds: parseJsonArray(row[12]),
        contractCoverageVersion: String(row[13] || ""),
        complete: String(row[14] || "TRUE").toUpperCase() !== "FALSE",
      });
    }
  }

  return {
    version: String(meta.manifestVersion || ""),
    complete: String(meta.complete || "").toUpperCase() === "TRUE",
    templateId: String(meta.templateId || ""),
    templateTitle: String(meta.templateTitle || ""),
    resultType: String(meta.resultType || ""),
    generatedAt: String(meta.generatedAt || ""),
    sectionCount: Number(meta.sectionCount || sections.length),
    renderedMetricIds: parseJsonArray(meta.renderedMetricIdsJson),
    expectedMetricIds: parseJsonArray(meta.expectedMetricIdsJson),
    missingMetricIds: parseJsonArray(meta.missingMetricIdsJson),
    metricCoveragePass:
      String(meta.metricCoveragePass || "").toUpperCase() === "TRUE",
    metricCoverageRate: Number(meta.metricCoverageRate || 0),
    contractCoverageVersion: String(meta.contractCoverageVersion || ""),
    contractCatalogVersion: String(meta.contractCatalogVersion || ""),
    finalOutputQualityGateMetaVersion: String(
      meta.finalOutputQualityGateMetaVersion || "",
    ),
    finalOutputQualityGateMetaPresent:
      String(meta.finalOutputQualityGateMetaPresent || "").toUpperCase() ===
      "TRUE",
    finalOutputQualityGateMeta: parseFinalOutputQualityGateMeta(meta),
    sections,
  };
}

module.exports = {
  SUMMARY_SHEET_RECIPE_MANIFEST_VERSION,
  SUMMARY_SHEET_RECIPE_MANIFEST_SHEET,
  FINAL_OUTPUT_QUALITY_GATE_MANIFEST_SERIALIZATION_VERSION,
  FINAL_OUTPUT_QUALITY_GATE_META_KEYS,
  buildFinalOutputQualityGateMeta,
  parseFinalOutputQualityGateMeta,
  serializeManifestMetaValue,
  manifestToAoa,
  buildSectionManifest,
  buildSummarySheetRecipeManifest,
  appendSummarySheetRecipeManifest,
  readSummarySheetRecipeManifest,
};
