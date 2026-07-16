const path = require("path");

const LOCAL_REGRESSION_ARTIFACT_NAMING_VERSION =
  "local_regression_artifact_naming_v1";

const OUTPUT_LABELS = Object.freeze({
  queryTables: "queryTables",
  summarySheet: "자동화시트",
  analysisReport: "분석보고서",
  ppt: "PPT",
});

function sanitizeFilePart(value = "", fallback = "regression-case") {
  const normalized = String(value || "")
    .normalize("NFKC")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^[._\s]+|[._\s]+$/g, "")
    .slice(0, 140);

  return normalized || fallback;
}

function normalizeRegressionContext(value = null) {
  if (!value || typeof value !== "object") return null;

  const caseId = String(value.caseId || "").trim();
  if (value.enabled !== true || !caseId) return null;

  return {
    enabled: true,
    source: String(value.source || "automation-regression").trim(),
    runId: String(value.runId || "").trim(),
    caseId,
    templateId: String(value.templateId || "").trim(),
    retention: String(value.retention || "latest").trim() || "latest",
  };
}

function shouldUseLocalRegressionArtifactNaming({
  regressionContext = null,
  localStorageEnabled = false,
  nodeEnv = "",
} = {}) {
  return Boolean(
    localStorageEnabled &&
      String(nodeEnv || "").trim().toLowerCase() !== "production" &&
      normalizeRegressionContext(regressionContext),
  );
}

function normalizeExtension(extension = "") {
  return String(extension || "")
    .trim()
    .replace(/^\.+/, "")
    .toLowerCase();
}

function buildLocalRegressionArtifactFileName({
  regressionContext = null,
  outputType = "",
  extension = "",
  localStorageEnabled = false,
  nodeEnv = "",
} = {}) {
  if (
    !shouldUseLocalRegressionArtifactNaming({
      regressionContext,
      localStorageEnabled,
      nodeEnv,
    })
  ) {
    return "";
  }

  const context = normalizeRegressionContext(regressionContext);
  const safeCaseId = sanitizeFilePart(context.caseId);
  const label = OUTPUT_LABELS[outputType] || sanitizeFilePart(outputType, "artifact");
  const ext = normalizeExtension(extension);

  if (!ext) {
    throw new Error("회귀 산출물 확장자가 필요합니다.");
  }

  return `${safeCaseId}__${label}.${ext}`;
}

function buildLocalRegressionQueryTablesKey({
  regressionContext = null,
  localStorageEnabled = false,
  nodeEnv = "",
} = {}) {
  const fileName = buildLocalRegressionArtifactFileName({
    regressionContext,
    outputType: "queryTables",
    extension: "json",
    localStorageEnabled,
    nodeEnv,
  });

  return fileName ? `query-tables/${fileName}` : "";
}

function isFlatLocalRegressionQueryTablesKey({
  key = "",
  localStorageEnabled = false,
  nodeEnv = "",
} = {}) {
  if (
    !localStorageEnabled ||
    String(nodeEnv || "").trim().toLowerCase() === "production"
  ) {
    return false;
  }

  const normalized = String(key || "").replace(/\\/g, "/");
  if (!normalized.startsWith("query-tables/")) return false;

  const relative = normalized.slice("query-tables/".length);
  return (
    Boolean(relative) &&
    !relative.includes("/") &&
    path.extname(relative).toLowerCase() === ".json" &&
    relative.includes("__queryTables")
  );
}

module.exports = {
  LOCAL_REGRESSION_ARTIFACT_NAMING_VERSION,
  sanitizeFilePart,
  normalizeRegressionContext,
  shouldUseLocalRegressionArtifactNaming,
  buildLocalRegressionArtifactFileName,
  buildLocalRegressionQueryTablesKey,
  isFlatLocalRegressionQueryTablesKey,
};
