const fs = require("fs");
const path = require("path");
const {
  QUERY_CANDIDATE_OBSERVATION_VERSION,
  normalizeText,
} = require("./queryCandidateObservation");

const QUERY_CANDIDATE_TRACE_VERSION = "query_candidate_trace_v1";

function normalizeBoolean(value, defaultValue = false) {
  if (value == null || value === "") return defaultValue;
  return !["0", "false", "off", "no"].includes(
    String(value).trim().toLowerCase(),
  );
}

function traceEnabled(options = {}) {
  if (typeof options.enabled === "boolean") return options.enabled;
  return normalizeBoolean(process.env.QUERY_CANDIDATE_TRACE, false);
}

function safeName(value = "trace") {
  return (
    normalizeText(value || "trace")
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .slice(0, 120) || "trace"
  );
}

function defaultTracePath({ caseId = "", runId = "" } = {}) {
  const root = path.resolve(
    process.env.QUERY_CANDIDATE_TRACE_DIR ||
      path.join(process.cwd(), "tests", "results", "query-candidate-traces"),
  );
  const fileName = `${safeName(runId || "manual")}_${safeName(
    caseId || "case",
  )}.ndjson`;
  return path.join(root, fileName);
}

function compactObservation(observation = {}) {
  return {
    observationVersion:
      observation.version || QUERY_CANDIDATE_OBSERVATION_VERSION,
    observationSha256: observation.observationSha256 || "",
    queryShape: {
      tableCount: observation.queryShape?.tableCount || 0,
      analysisEligibleCount: observation.queryShape?.analysisEligibleCount || 0,
      templateEligibleCount: observation.queryShape?.templateEligibleCount || 0,
      shapeSha256: observation.queryShape?.shapeSha256 || "",
    },
    candidateObservation: {
      candidateContractVersion:
        observation.candidateObservation?.candidateContractVersion || "",
      candidateScoringVersion:
        observation.candidateObservation?.candidateScoringVersion || "",
      candidateUiPayloadVersion:
        observation.candidateObservation?.candidateUiPayloadVersion || "",
      counts: observation.candidateObservation?.counts || {},
      idsByClass: observation.candidateObservation?.idsByClass || {},
      topOrder: observation.candidateObservation?.topOrder || [],
      uiRecommendedOrder:
        observation.candidateObservation?.uiRecommendedOrder || [],
    },
  };
}

function createQueryCandidateTraceEvent({
  stage = "observation",
  caseId = "",
  fileName = "",
  runId = "",
  status = "OK",
  observation = null,
  details = null,
  at = new Date().toISOString(),
} = {}) {
  return {
    version: QUERY_CANDIDATE_TRACE_VERSION,
    at,
    stage: normalizeText(stage),
    status: normalizeText(status),
    runId: normalizeText(runId),
    caseId: normalizeText(caseId),
    fileName: normalizeText(fileName),
    observation: observation ? compactObservation(observation) : null,
    details: details && typeof details === "object" ? details : null,
  };
}

function appendQueryCandidateTrace(event = {}, options = {}) {
  if (!traceEnabled(options)) {
    return { written: false, reason: "trace_disabled", path: "" };
  }
  const targetPath = path.resolve(
    options.path ||
      defaultTracePath({ caseId: event.caseId, runId: event.runId }),
  );
  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  fs.appendFileSync(targetPath, `${JSON.stringify(event)}\n`, "utf8");
  return { written: true, reason: "", path: targetPath };
}

module.exports = {
  QUERY_CANDIDATE_TRACE_VERSION,
  traceEnabled,
  safeName,
  defaultTracePath,
  compactObservation,
  createQueryCandidateTraceEvent,
  appendQueryCandidateTrace,
};
