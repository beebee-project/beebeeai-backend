const crypto = require("crypto");
const {
  evaluateAccuracyDataset,
} = require("./queryCandidatePlannerAccuracyEvaluator");
const {
  calculateProviderCostMicrousd,
  evaluateCostCacheLatency,
} = require("./queryCandidatePlannerCostCacheLatencyEvaluator");
const {
  evaluateShadowAccuracy,
  observationsToPredictions,
} = require("./queryCandidatePlannerShadowAccuracyEvaluator");
const {
  EVIDENCE_VERSION,
  validateQueryCandidatePlannerInternalCanaryEvidence,
} = require("./queryCandidatePlannerInternalCanaryEvidence");

const BUILDER_VERSION =
  "query_candidate_planner_real_shadow_evidence_bundle_builder_v1";
const MIN_SAMPLE_SIZE = 30;
const MIN_OBSERVATIONS_PER_CASE = 3;

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function clone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, canonicalize(value[key])]),
  );
}

function sha256(value) {
  return crypto
    .createHash("sha256")
    .update(
      typeof value === "string" ? value : JSON.stringify(canonicalize(value)),
    )
    .digest("hex");
}

function text(value, maxLength = 160) {
  return String(value == null ? "" : value)
    .trim()
    .slice(0, maxLength);
}

function recordPayload(record) {
  return isPlainObject(record?.payload) ? record.payload : record;
}

function normalizeRecords(records = []) {
  return records
    .filter((record) => isPlainObject(record))
    .map((record) => ({ ...record, payload: recordPayload(record) }))
    .filter(
      (record) =>
        record.source === "REAL_SHADOW_TRAFFIC" ||
        record.payload?.source === "REAL_SHADOW_TRAFFIC",
    )
    .filter(
      (record) =>
        record.actualTraffic === true || record.payload?.actualTraffic === true,
    )
    .filter(
      (record) =>
        record.synthetic !== true && record.payload?.synthetic !== true,
    )
    .sort(
      (a, b) =>
        String(a.observedAt || a.payload?.observedAt || "").localeCompare(
          String(b.observedAt || b.payload?.observedAt || ""),
        ) || String(a.recordId || "").localeCompare(String(b.recordId || "")),
    );
}

function predictionKey(observation) {
  return sha256(observation?.shadowPrediction || {});
}

function selectConsensusObservations(executionRecords, accuracyDataset) {
  const byCase = new Map();
  for (const record of executionRecords) {
    const observation = record.payload?.shadowAccuracyObservation;
    if (!isPlainObject(observation)) continue;
    const caseId = text(observation.caseId);
    if (!byCase.has(caseId)) byCase.set(caseId, []);
    byCase.get(caseId).push(observation);
  }
  const selected = [];
  const counts = {};
  const errors = [];
  for (const item of accuracyDataset.cases || []) {
    const caseId = text(item.caseId);
    const observations = byCase.get(caseId) || [];
    counts[caseId] = observations.length;
    if (observations.length < MIN_OBSERVATIONS_PER_CASE) {
      errors.push(`INSUFFICIENT_CASE_OBSERVATIONS:${caseId}`);
      continue;
    }
    const groups = new Map();
    for (const observation of observations) {
      const key = predictionKey(observation);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(observation);
    }
    const winner = [...groups.entries()].sort(
      (left, right) =>
        right[1].length - left[1].length || left[0].localeCompare(right[0]),
    )[0];
    selected.push(
      winner[1]
        .slice()
        .sort((a, b) =>
          String(a.observationId).localeCompare(String(b.observationId)),
        )[0],
    );
  }
  return Object.freeze({
    valid: errors.length === 0,
    errors: Object.freeze(errors),
    selected: Object.freeze(selected),
    counts: Object.freeze(counts),
  });
}

function providerCost(provider, pricingPolicy) {
  const execution = {
    provider: {
      called: provider.called === true,
      modelId: text(provider.modelId),
      inputTokens: Math.max(0, Math.trunc(Number(provider.inputTokens) || 0)),
      outputTokens: Math.max(0, Math.trunc(Number(provider.outputTokens) || 0)),
      ...(Number.isInteger(provider.observedCostMicrousd) &&
      provider.observedCostMicrousd > 0
        ? { observedCostMicrousd: provider.observedCostMicrousd }
        : {}),
    },
  };
  const result = calculateProviderCostMicrousd(execution, pricingPolicy);
  return result.valid ? result.costMicrousd : 0;
}

function buildOperationalDataset(records, pricingPolicy) {
  const executions = records.filter((record) => record.kind === "EXECUTION");
  const lifecycle = records.filter((record) => record.kind === "LIFECYCLE");
  const state = new Map();
  const lifecycleEvents = [];
  const preliminary = [];

  const scenarioState = (scenarioId) => {
    if (!state.has(scenarioId)) {
      state.set(scenarioId, {
        seen: false,
        lastEvent: "",
        lastIdentity: "",
        deletedIdentity: "",
        expectedColdCostMicrousd: 0,
      });
    }
    return state.get(scenarioId);
  };

  for (const record of records) {
    const payload = record.payload || {};
    const scenarioId = text(
      record.scenarioId || payload.scenarioId || payload.caseId,
    );
    const current = scenarioState(scenarioId);
    if (record.kind === "LIFECYCLE") {
      const item = payload.lifecycle || {};
      const event = text(item.event).toUpperCase();
      const identity = text(item.uploadFingerprintSha256, 64).toLowerCase();
      lifecycleEvents.push({
        eventId:
          record.recordId ||
          sha256({ scenarioId, event, observedAt: record.observedAt }),
        scenarioId,
        event,
        cacheDisposition: text(item.cacheDisposition) || "UNKNOWN",
        invalidationAttempted: item.invalidationAttempted === true,
        invalidationSucceeded: item.invalidationSucceeded === true,
        staleCacheReused: item.staleCacheReused === true,
        priorUploadIdentitySha256:
          event === "DELETE" || event === "REUPLOAD" ? identity : "",
        newUploadIdentitySha256: "",
      });
      current.lastEvent = event;
      if (event === "DELETE" || event === "REUPLOAD")
        current.deletedIdentity = identity;
      if (identity) current.lastIdentity = identity;
      continue;
    }
    if (record.kind !== "EXECUTION") continue;
    const op = payload.operational || {};
    const provider = { ...(op.provider || {}) };
    if (provider.called === true && !text(provider.modelId)) {
      provider.modelId = text(
        op.modelIdFallback || "semantic_profiler_default",
      );
    }
    const identity = text(
      op.lifecycleHints?.uploadFingerprintSha256,
      64,
    ).toLowerCase();
    const afterDownload = current.lastEvent === "DOWNLOAD";
    const afterReupload = Boolean(
      current.deletedIdentity &&
      identity &&
      current.deletedIdentity !== identity,
    );
    let phase = "WARM";
    if (afterReupload) phase = "REUPLOAD";
    else if (afterDownload) phase = "DOWNLOAD_REUSE";
    else if (!current.seen || provider.called === true) phase = "COLD";
    if (provider.called === true) {
      current.expectedColdCostMicrousd = Math.max(
        current.expectedColdCostMicrousd,
        Number(op.expectedColdCostMicrousd) || 0,
        providerCost(provider, pricingPolicy),
      );
    }
    if (afterReupload) {
      lifecycleEvents.push({
        eventId: sha256({
          scenarioId,
          prior: current.deletedIdentity,
          next: identity,
          execution: record.recordId,
        }),
        scenarioId,
        event: "REUPLOAD",
        cacheDisposition: "NEW_IDENTITY",
        invalidationAttempted: true,
        invalidationSucceeded: true,
        staleCacheReused: false,
        priorUploadIdentitySha256: current.deletedIdentity,
        newUploadIdentitySha256: identity,
      });
      current.deletedIdentity = "";
    }
    preliminary.push({
      executionId: record.recordId,
      scenarioId,
      phase,
      status: text(op.status).toUpperCase() || "ERROR",
      latencyMs: Math.max(0, Number(op.latencyMs) || 0),
      explicitExpectedColdCostMicrousd:
        Number(op.expectedColdCostMicrousd) || 0,
      cache: {
        readAttempted: op.cache?.readAttempted === true,
        hit: op.cache?.hit === true,
        level: text(op.cache?.level).toUpperCase() || "NONE",
        writeAttempted: op.cache?.writeAttempted === true,
        writeSucceeded: op.cache?.writeSucceeded === true,
      },
      provider: {
        called: provider.called === true,
        modelId: text(provider.modelId),
        inputTokens: Math.max(0, Math.trunc(Number(provider.inputTokens) || 0)),
        outputTokens: Math.max(
          0,
          Math.trunc(Number(provider.outputTokens) || 0),
        ),
        ...(Number.isInteger(provider.observedCostMicrousd) &&
        provider.observedCostMicrousd > 0
          ? { observedCostMicrousd: provider.observedCostMicrousd }
          : {}),
      },
      lifecycleContext: {
        afterDownload,
        afterReupload,
        staleCacheReused: op.lifecycleHints?.staleCacheReused === true,
      },
    });
    current.seen = true;
    current.lastEvent = "";
    if (identity) current.lastIdentity = identity;
  }

  const executionsFinal = preliminary.map((execution) => {
    const expected = Math.max(
      execution.explicitExpectedColdCostMicrousd,
      state.get(execution.scenarioId)?.expectedColdCostMicrousd || 0,
    );
    const { explicitExpectedColdCostMicrousd: _removed, ...rest } = execution;
    return Object.freeze({
      ...rest,
      expectedColdCostMicrousd: Math.trunc(expected),
    });
  });

  return Object.freeze({
    version: "query_candidate_planner_operational_evaluation_dataset_v1",
    datasetId: `real_shadow_operational_${sha256(executionsFinal).slice(0, 16)}`,
    benchmarkMode: "REAL_SHADOW_TRAFFIC_APPROVED_ACTUAL_PRICING",
    executions: Object.freeze(executionsFinal),
    lifecycleEvents: Object.freeze(lifecycleEvents),
    guardrails: Object.freeze({
      source: "REAL_SHADOW_TRAFFIC",
      actualTraffic: true,
      synthetic: false,
      rawRowsIncluded: false,
      fileNamesIncluded: false,
      userIdentityIncluded: false,
      evaluationOnly: true,
      promotionAuthorized: false,
    }),
  });
}

function reportSummary(report, sampleSize, extra = {}) {
  return Object.freeze({
    version: report.version,
    decision: report.decision,
    failClosed: report.failClosed === true,
    evaluationOnly: report.evaluationOnly !== false,
    promotionAuthorized: false,
    sampleSize,
    reportSha256: text(report.reportSha256, 64) || sha256(report),
    ...extra,
  });
}

function blocked(reason, details = {}) {
  return Object.freeze({
    version: BUILDER_VERSION,
    decision: "EVALUATION_BLOCKED",
    reason,
    failClosed: true,
    promotionAuthorized: false,
    evidenceBundle: null,
    details: Object.freeze(details),
  });
}

function buildQueryCandidatePlannerRealShadowEvidenceBundle({
  records,
  readiness,
  accuracyDataset,
  accuracyThresholdPolicy,
  operationalThresholdPolicy,
  approvedActualPricingPolicy,
  shadowThresholdPolicy,
  evaluatedAt = new Date().toISOString(),
  expiresInHours = 24,
  now = Date.now,
} = {}) {
  const normalized = normalizeRecords(records);
  const executionRecords = normalized.filter(
    (record) => record.kind === "EXECUTION",
  );
  if (executionRecords.length < MIN_SAMPLE_SIZE) {
    return blocked("REAL_SHADOW_MINIMUM_SAMPLE_SIZE_NOT_MET", {
      actual: executionRecords.length,
      required: MIN_SAMPLE_SIZE,
    });
  }
  if (approvedActualPricingPolicy?.mode !== "APPROVED_ACTUAL") {
    return blocked("APPROVED_ACTUAL_PRICING_REQUIRED");
  }
  const consensus = selectConsensusObservations(
    executionRecords,
    accuracyDataset || { cases: [] },
  );
  if (!consensus.valid) {
    return blocked(consensus.errors[0], {
      errors: consensus.errors,
      counts: consensus.counts,
    });
  }
  const shadowDataset = Object.freeze({
    version: "query_candidate_planner_shadow_accuracy_observation_dataset_v1",
    datasetId: `real_shadow_accuracy_${sha256(consensus.selected).slice(0, 16)}`,
    capturePolicyVersion:
      "query_candidate_planner_shadow_accuracy_capture_policy_v1",
    sourceAccuracyDatasetVersion: accuracyDataset.version,
    sourceAccuracyDatasetId: accuracyDataset.datasetId,
    observations: consensus.selected,
    guardrails: Object.freeze({
      source: "REAL_SHADOW_TRAFFIC",
      actualTraffic: true,
      synthetic: false,
      evaluationOnly: true,
      productionMergeAuthorized: false,
    }),
  });
  const shadowReport = evaluateShadowAccuracy({
    shadowObservationDataset: shadowDataset,
    accuracyDataset,
    shadowThresholdPolicy,
    accuracyThresholdPolicy,
  });
  const predictions = observationsToPredictions(consensus.selected);
  const accuracyReport = evaluateAccuracyDataset({
    dataset: accuracyDataset,
    predictions,
    thresholdPolicy: accuracyThresholdPolicy,
  });
  const operationalDataset = buildOperationalDataset(
    normalized,
    approvedActualPricingPolicy,
  );
  const operationalReport = evaluateCostCacheLatency({
    dataset: operationalDataset,
    thresholdPolicy: operationalThresholdPolicy,
    pricingPolicy: approvedActualPricingPolicy,
  });
  if (accuracyReport.decision !== "EVALUATION_PASS") {
    return blocked("REAL_SHADOW_ACCURACY_EVALUATION_BLOCKED", {
      accuracyReport,
    });
  }
  if (shadowReport.decision !== "EVALUATION_PASS") {
    return blocked("REAL_SHADOW_SHADOW_EVALUATION_BLOCKED", { shadowReport });
  }
  if (operationalReport.decision !== "EVALUATION_PASS") {
    return blocked("REAL_SHADOW_OPERATIONAL_EVALUATION_BLOCKED", {
      operationalReport,
    });
  }
  const evaluatedMs = Date.parse(evaluatedAt);
  const currentMs = Number(now());
  if (!Number.isFinite(evaluatedMs) || evaluatedMs > currentMs + 60000) {
    return blocked("REAL_SHADOW_EVALUATED_AT_INVALID");
  }
  const hours = Math.max(1, Math.min(168, Number(expiresInHours) || 24));
  const expiresAt = new Date(evaluatedMs + hours * 3600000).toISOString();
  const shadowSummary = shadowReport.observationSummary || {};
  const evidenceBundle = Object.freeze({
    version: EVIDENCE_VERSION,
    source: "REAL_SHADOW_TRAFFIC",
    synthetic: false,
    actualTraffic: true,
    evaluatedAt: new Date(evaluatedMs).toISOString(),
    expiresAt,
    readiness: clone(readiness),
    accuracy: reportSummary(accuracyReport, executionRecords.length, {
      caseCount: accuracyReport.caseCount,
      observationCountsByCase: consensus.counts,
    }),
    operational: reportSummary(
      operationalReport,
      operationalReport.sample?.executions || 0,
      {
        pricingSource: "APPROVED_ACTUAL",
        pricingPolicyId: approvedActualPricingPolicy.policyId,
        executionCount: operationalReport.sample?.executions || 0,
        lifecycleEventCount: operationalReport.sample?.lifecycleEvents || 0,
      },
    ),
    shadow: reportSummary(shadowReport, executionRecords.length, {
      primaryResponseUnchangedRate: shadowSummary.primaryResponseUnchangedRate,
      guardrailViolationCount: shadowSummary.guardrailViolationCount,
      privacyViolationCount: shadowSummary.privacyViolationCount,
      completedRate: shadowSummary.completedRate,
      comparisonCoverage: shadowSummary.comparisonCoverage,
    }),
    llmPolicy: Object.freeze({
      mode: "SEMANTIC_PROFILER_ONLY",
      plannerEscalationAllowed: false,
    }),
  });
  const validation = validateQueryCandidatePlannerInternalCanaryEvidence(
    evidenceBundle,
    { now },
  );
  if (!validation.valid) {
    return blocked(validation.reason, { validation });
  }
  return Object.freeze({
    version: BUILDER_VERSION,
    decision: "EVALUATION_PASS",
    reason: "REAL_SHADOW_EVIDENCE_BUNDLE_VALID",
    failClosed: true,
    promotionAuthorized: false,
    evidenceBundle,
    evidenceSha256: validation.evidenceSha256,
    datasets: Object.freeze({
      shadow: shadowDataset,
      operational: operationalDataset,
    }),
    reports: Object.freeze({
      accuracy: accuracyReport,
      operational: operationalReport,
      shadow: shadowReport,
    }),
    validation,
  });
}

module.exports = Object.freeze({
  BUILDER_VERSION,
  MIN_SAMPLE_SIZE,
  MIN_OBSERVATIONS_PER_CASE,
  normalizeRecords,
  selectConsensusObservations,
  buildOperationalDataset,
  buildQueryCandidatePlannerRealShadowEvidenceBundle,
  sha256,
});
