"use strict";

const crypto = require("crypto");
const {
  READINESS_DECISION,
  evaluateReadinessGate,
} = require("./queryCandidatePlannerFeatureControl");

const EVIDENCE_VERSION =
  "query_candidate_planner_internal_canary_evidence_bundle_v1";
const EVIDENCE_SOURCE = "REAL_SHADOW_TRAFFIC";
const MAX_EVIDENCE_AGE_MS = 7 * 24 * 60 * 60 * 1000;
const MIN_SAMPLE_SIZE = 30;
const SHA256_RE = /^[a-f0-9]{64}$/i;

const FORBIDDEN_KEYS = new Set([
  "rows",
  "rawRows",
  "sampleValues",
  "fileName",
  "originalFileName",
  "email",
  "userId",
  "accountId",
  "tenantId",
  "queryTablesKey",
  "storageKey",
  "cacheSecret",
  "apiKey",
  "prompt",
  "rawPayload",
]);

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function sha256(value) {
  return crypto
    .createHash("sha256")
    .update(typeof value === "string" ? value : JSON.stringify(value))
    .digest("hex");
}

function findForbiddenPaths(value, path = "root", output = []) {
  if (Array.isArray(value)) {
    value.forEach((entry, index) =>
      findForbiddenPaths(entry, `${path}[${index}]`, output),
    );
    return output;
  }
  if (!isPlainObject(value)) return output;
  for (const [key, child] of Object.entries(value)) {
    if (FORBIDDEN_KEYS.has(key)) output.push(`${path}.${key}`);
    findForbiddenPaths(child, `${path}.${key}`, output);
  }
  return output;
}

function validDate(value) {
  const parsed = Date.parse(String(value || ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function validReport(report, expectedVersion) {
  return Boolean(
    isPlainObject(report) &&
      report.version === expectedVersion &&
      report.decision === "EVALUATION_PASS" &&
      report.failClosed === true &&
      report.evaluationOnly === true &&
      report.promotionAuthorized === false &&
      Number.isInteger(report.sampleSize) &&
      report.sampleSize >= MIN_SAMPLE_SIZE &&
      SHA256_RE.test(String(report.reportSha256 || "")),
  );
}

function parseEvidenceJson(raw) {
  if (!String(raw || "").trim()) {
    return { value: null, error: "CANARY_EVIDENCE_NOT_CONFIGURED" };
  }
  try {
    return { value: JSON.parse(String(raw)), error: "" };
  } catch (_error) {
    return { value: null, error: "CANARY_EVIDENCE_JSON_INVALID" };
  }
}

function validateQueryCandidatePlannerInternalCanaryEvidence(
  evidence,
  { now = Date.now } = {},
) {
  const errors = [];
  if (!isPlainObject(evidence)) {
    errors.push("EVIDENCE_OBJECT_REQUIRED");
  }
  if (evidence?.version !== EVIDENCE_VERSION) {
    errors.push("EVIDENCE_VERSION_INVALID");
  }
  if (evidence?.source !== EVIDENCE_SOURCE) {
    errors.push("REAL_SHADOW_TRAFFIC_EVIDENCE_REQUIRED");
  }
  if (evidence?.synthetic !== false) {
    errors.push("SYNTHETIC_EVIDENCE_FORBIDDEN");
  }
  if (evidence?.actualTraffic !== true) {
    errors.push("ACTUAL_TRAFFIC_EVIDENCE_REQUIRED");
  }

  const evaluatedAt = validDate(evidence?.evaluatedAt);
  const expiresAt = validDate(evidence?.expiresAt);
  const current = Number(now());
  if (evaluatedAt === null || expiresAt === null) {
    errors.push("EVIDENCE_TIME_WINDOW_INVALID");
  } else {
    if (evaluatedAt > current + 60 * 1000) {
      errors.push("EVIDENCE_FROM_FUTURE");
    }
    if (expiresAt <= current) errors.push("EVIDENCE_EXPIRED");
    if (current - evaluatedAt > MAX_EVIDENCE_AGE_MS) {
      errors.push("EVIDENCE_TOO_OLD");
    }
  }

  const readiness = evaluateReadinessGate(evidence?.readiness);
  if (!readiness.valid || evidence?.readiness?.decision !== READINESS_DECISION) {
    errors.push("PATCH13_3_READINESS_EVIDENCE_INVALID");
  }

  if (!validReport(
    evidence?.accuracy,
    "query_candidate_planner_accuracy_evaluation_report_v1",
  )) {
    errors.push("ACCURACY_EVIDENCE_INVALID");
  }
  if (!validReport(
    evidence?.operational,
    "query_candidate_planner_cost_cache_latency_evaluation_report_v1",
  )) {
    errors.push("OPERATIONAL_EVIDENCE_INVALID");
  }
  if (!validReport(
    evidence?.shadow,
    "query_candidate_planner_shadow_accuracy_evaluation_report_v1",
  )) {
    errors.push("SHADOW_EVIDENCE_INVALID");
  }

  if (evidence?.shadow?.primaryResponseUnchangedRate !== 1) {
    errors.push("PRIMARY_RESPONSE_UNCHANGED_RATE_REQUIRED");
  }
  if (evidence?.shadow?.guardrailViolationCount !== 0) {
    errors.push("GUARDRAIL_VIOLATION_PRESENT");
  }
  if (evidence?.shadow?.privacyViolationCount !== 0) {
    errors.push("PRIVACY_VIOLATION_PRESENT");
  }
  if (evidence?.operational?.pricingSource !== "APPROVED_ACTUAL") {
    errors.push("APPROVED_ACTUAL_PRICING_REQUIRED");
  }
  if (evidence?.llmPolicy?.mode !== "SEMANTIC_PROFILER_ONLY") {
    errors.push("SEMANTIC_PROFILER_ONLY_REQUIRED");
  }
  if (evidence?.llmPolicy?.plannerEscalationAllowed !== false) {
    errors.push("PLANNER_ESCALATION_MUST_BE_DISABLED");
  }

  const forbiddenPaths = findForbiddenPaths(evidence);
  if (forbiddenPaths.length > 0) errors.push("SENSITIVE_EVIDENCE_FIELD_PRESENT");

  const uniqueErrors = Object.freeze([...new Set(errors)]);
  return Object.freeze({
    version: EVIDENCE_VERSION,
    valid: uniqueErrors.length === 0,
    reason: uniqueErrors.length === 0
      ? "REAL_SHADOW_EVIDENCE_VALID"
      : uniqueErrors[0],
    errors: uniqueErrors,
    readiness: readiness.valid ? evidence.readiness : null,
    evidenceSha256: isPlainObject(evidence) ? sha256(evidence) : "",
    summary: Object.freeze({
      source: String(evidence?.source || ""),
      actualTraffic: evidence?.actualTraffic === true,
      synthetic: evidence?.synthetic === true,
      evaluatedAt: String(evidence?.evaluatedAt || ""),
      expiresAt: String(evidence?.expiresAt || ""),
      accuracySampleSize: Number(evidence?.accuracy?.sampleSize || 0),
      operationalSampleSize: Number(evidence?.operational?.sampleSize || 0),
      shadowSampleSize: Number(evidence?.shadow?.sampleSize || 0),
      rawEvidenceIncluded: false,
      forbiddenPathsIncluded: false,
    }),
    failClosed: true,
  });
}

module.exports = Object.freeze({
  EVIDENCE_VERSION,
  EVIDENCE_SOURCE,
  MAX_EVIDENCE_AGE_MS,
  MIN_SAMPLE_SIZE,
  FORBIDDEN_KEYS,
  parseEvidenceJson,
  findForbiddenPaths,
  validateQueryCandidatePlannerInternalCanaryEvidence,
});
