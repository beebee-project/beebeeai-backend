const crypto = require("crypto");

const EVALUATOR_VERSION =
  "query_candidate_planner_cost_cache_latency_evaluator_v1";
const REPORT_VERSION =
  "query_candidate_planner_cost_cache_latency_evaluation_report_v1";
const DATASET_VERSION =
  "query_candidate_planner_operational_evaluation_dataset_v1";
const THRESHOLD_POLICY_VERSION =
  "query_candidate_planner_operational_threshold_policy_v1";
const PRICING_POLICY_VERSION = "query_candidate_planner_cost_pricing_policy_v1";

const DECISIONS = Object.freeze({
  PASS: "EVALUATION_PASS",
  BLOCKED: "EVALUATION_BLOCKED",
});

const EXECUTION_PHASES = Object.freeze([
  "COLD",
  "WARM",
  "DOWNLOAD_REUSE",
  "REUPLOAD",
]);
const EXECUTION_STATUSES = Object.freeze([
  "SUCCESS",
  "TIMEOUT",
  "ERROR",
  "BLOCKED",
]);
const CACHE_LEVELS = Object.freeze(["L1", "L2", "L3", "L4", "MISS", "NONE"]);
const HIT_LEVELS = new Set(["L1", "L2", "L3", "L4"]);
const LIFECYCLE_EVENTS = Object.freeze(["DOWNLOAD", "DELETE", "REUPLOAD"]);
const SHA256_RE = /^[a-f0-9]{64}$/i;

const SENSITIVE_KEYS = new Set([
  "rows",
  "rawRows",
  "sampleValues",
  "fileName",
  "originalFileName",
  "email",
  "userId",
  "tenantId",
  "queryTablesKey",
  "storageKey",
  "rawPayload",
  "cacheSecret",
  "encryptionKey",
  "apiKey",
]);

function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function text(value) {
  return String(value == null ? "" : value).trim();
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

function canonicalJson(value) {
  return JSON.stringify(canonicalize(value));
}

function sha256(value) {
  const serialized = typeof value === "string" ? value : canonicalJson(value);
  return crypto.createHash("sha256").update(serialized).digest("hex");
}

function freezeDeep(value) {
  if (Array.isArray(value)) {
    value.forEach(freezeDeep);
    return Object.freeze(value);
  }
  if (isPlainObject(value) && !Object.isFrozen(value)) {
    Object.values(value).forEach(freezeDeep);
    Object.freeze(value);
  }
  return value;
}

function round(value, digits = 6) {
  if (!Number.isFinite(value)) return 0;
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function safeRate(numerator, denominator) {
  if (
    !Number.isFinite(numerator) ||
    !Number.isFinite(denominator) ||
    denominator <= 0
  ) {
    return 0;
  }
  return round(numerator / denominator);
}

function mean(values) {
  const finite = values.filter(Number.isFinite);
  if (finite.length === 0) return 0;
  return round(finite.reduce((sum, value) => sum + value, 0) / finite.length);
}

function percentile(values, percentileValue) {
  const finite = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (finite.length === 0) return 0;
  const p = Math.min(1, Math.max(0, Number(percentileValue)));
  if (p === 0) return finite[0];
  const rank = Math.ceil(p * finite.length) - 1;
  return finite[Math.max(0, Math.min(finite.length - 1, rank))];
}

function findSensitivePaths(value, basePath = "$") {
  const paths = [];
  if (Array.isArray(value)) {
    value.forEach((entry, index) => {
      paths.push(...findSensitivePaths(entry, `${basePath}[${index}]`));
    });
    return paths;
  }
  if (!isPlainObject(value)) return paths;
  for (const [key, entry] of Object.entries(value)) {
    const childPath = `${basePath}.${key}`;
    if (SENSITIVE_KEYS.has(key)) paths.push(childPath);
    paths.push(...findSensitivePaths(entry, childPath));
  }
  return paths;
}

function validateBoolean(value, path, errors) {
  if (typeof value !== "boolean") errors.push(`${path} must be boolean`);
}

function validateNonNegativeInteger(
  value,
  path,
  errors,
  { optional = false } = {},
) {
  if (optional && value == null) return;
  if (!Number.isInteger(value) || value < 0) {
    errors.push(`${path} must be a non-negative integer`);
  }
}

function validatePositiveNumber(
  value,
  path,
  errors,
  { allowZero = true } = {},
) {
  const valid = Number.isFinite(value) && (allowZero ? value >= 0 : value > 0);
  if (!valid)
    errors.push(
      `${path} must be a finite ${allowZero ? "non-negative" : "positive"} number`,
    );
}

function validateOperationalEvaluationDataset(dataset) {
  const errors = [];
  if (!isPlainObject(dataset)) {
    return freezeDeep({ valid: false, errors: ["dataset must be an object"] });
  }
  if (dataset.version !== DATASET_VERSION) {
    errors.push(`dataset version must be ${DATASET_VERSION}`);
  }
  if (!text(dataset.datasetId)) errors.push("datasetId is required");
  if (!text(dataset.benchmarkMode)) errors.push("benchmarkMode is required");
  if (!Array.isArray(dataset.executions) || dataset.executions.length === 0) {
    errors.push("executions must be a non-empty array");
  }
  if (!Array.isArray(dataset.lifecycleEvents)) {
    errors.push("lifecycleEvents must be an array");
  }
  for (const path of findSensitivePaths(dataset)) {
    errors.push(`dataset contains forbidden sensitive field: ${path}`);
  }

  const executionIds = new Set();
  for (const [index, execution] of (dataset.executions || []).entries()) {
    const base = `executions[${index}]`;
    if (!isPlainObject(execution)) {
      errors.push(`${base} must be an object`);
      continue;
    }
    const executionId = text(execution.executionId);
    if (!executionId) errors.push(`${base}.executionId is required`);
    if (executionIds.has(executionId))
      errors.push(`duplicate executionId: ${executionId}`);
    executionIds.add(executionId);
    if (!text(execution.scenarioId))
      errors.push(`${base}.scenarioId is required`);
    if (!EXECUTION_PHASES.includes(execution.phase)) {
      errors.push(
        `${base}.phase must be one of ${EXECUTION_PHASES.join(", ")}`,
      );
    }
    if (!EXECUTION_STATUSES.includes(execution.status)) {
      errors.push(
        `${base}.status must be one of ${EXECUTION_STATUSES.join(", ")}`,
      );
    }
    validatePositiveNumber(execution.latencyMs, `${base}.latencyMs`, errors);
    validateNonNegativeInteger(
      execution.expectedColdCostMicrousd,
      `${base}.expectedColdCostMicrousd`,
      errors,
    );

    const cache = execution.cache;
    if (!isPlainObject(cache)) {
      errors.push(`${base}.cache is required`);
    } else {
      validateBoolean(
        cache.readAttempted,
        `${base}.cache.readAttempted`,
        errors,
      );
      validateBoolean(cache.hit, `${base}.cache.hit`, errors);
      validateBoolean(
        cache.writeAttempted,
        `${base}.cache.writeAttempted`,
        errors,
      );
      validateBoolean(
        cache.writeSucceeded,
        `${base}.cache.writeSucceeded`,
        errors,
      );
      if (!CACHE_LEVELS.includes(cache.level)) {
        errors.push(
          `${base}.cache.level must be one of ${CACHE_LEVELS.join(", ")}`,
        );
      }
      if (cache.hit === true && !HIT_LEVELS.has(cache.level)) {
        errors.push(`${base}.cache.hit requires L1-L4 level`);
      }
      if (cache.hit === false && HIT_LEVELS.has(cache.level)) {
        errors.push(`${base}.cache miss cannot use hit level`);
      }
      if (cache.hit === true && cache.readAttempted !== true) {
        errors.push(`${base}.cache.hit requires readAttempted`);
      }
      if (cache.writeSucceeded === true && cache.writeAttempted !== true) {
        errors.push(`${base}.cache.writeSucceeded requires writeAttempted`);
      }
    }

    const provider = execution.provider;
    if (!isPlainObject(provider)) {
      errors.push(`${base}.provider is required`);
    } else {
      validateBoolean(provider.called, `${base}.provider.called`, errors);
      validateNonNegativeInteger(
        provider.inputTokens,
        `${base}.provider.inputTokens`,
        errors,
      );
      validateNonNegativeInteger(
        provider.outputTokens,
        `${base}.provider.outputTokens`,
        errors,
      );
      validateNonNegativeInteger(
        provider.observedCostMicrousd,
        `${base}.provider.observedCostMicrousd`,
        errors,
        { optional: true },
      );
      if (provider.called === true && !text(provider.modelId)) {
        errors.push(
          `${base}.provider.modelId is required when provider is called`,
        );
      }
      if (provider.called !== true) {
        if (
          (provider.inputTokens || 0) !== 0 ||
          (provider.outputTokens || 0) !== 0
        ) {
          errors.push(
            `${base}.provider tokens must be zero when provider is not called`,
          );
        }
        if ((provider.observedCostMicrousd || 0) !== 0) {
          errors.push(
            `${base}.provider cost must be zero when provider is not called`,
          );
        }
      }
      if (cache?.hit === true && provider.called === true) {
        errors.push(`${base} cannot call provider on cache hit`);
      }
    }

    const lifecycle = execution.lifecycleContext;
    if (!isPlainObject(lifecycle)) {
      errors.push(`${base}.lifecycleContext is required`);
    } else {
      validateBoolean(
        lifecycle.afterDownload,
        `${base}.lifecycleContext.afterDownload`,
        errors,
      );
      validateBoolean(
        lifecycle.afterReupload,
        `${base}.lifecycleContext.afterReupload`,
        errors,
      );
      validateBoolean(
        lifecycle.staleCacheReused,
        `${base}.lifecycleContext.staleCacheReused`,
        errors,
      );
      if (
        execution.phase === "DOWNLOAD_REUSE" &&
        lifecycle.afterDownload !== true
      ) {
        errors.push(`${base} DOWNLOAD_REUSE requires afterDownload`);
      }
      if (execution.phase === "REUPLOAD" && lifecycle.afterReupload !== true) {
        errors.push(`${base} REUPLOAD requires afterReupload`);
      }
    }
  }

  const lifecycleIds = new Set();
  for (const [index, event] of (dataset.lifecycleEvents || []).entries()) {
    const base = `lifecycleEvents[${index}]`;
    if (!isPlainObject(event)) {
      errors.push(`${base} must be an object`);
      continue;
    }
    const eventId = text(event.eventId);
    if (!eventId) errors.push(`${base}.eventId is required`);
    if (lifecycleIds.has(eventId))
      errors.push(`duplicate lifecycle eventId: ${eventId}`);
    lifecycleIds.add(eventId);
    if (!text(event.scenarioId)) errors.push(`${base}.scenarioId is required`);
    if (!LIFECYCLE_EVENTS.includes(event.event)) {
      errors.push(
        `${base}.event must be one of ${LIFECYCLE_EVENTS.join(", ")}`,
      );
    }
    if (!text(event.cacheDisposition))
      errors.push(`${base}.cacheDisposition is required`);
    validateBoolean(
      event.invalidationAttempted,
      `${base}.invalidationAttempted`,
      errors,
    );
    validateBoolean(
      event.invalidationSucceeded,
      `${base}.invalidationSucceeded`,
      errors,
    );
    validateBoolean(event.staleCacheReused, `${base}.staleCacheReused`, errors);
    for (const key of [
      "priorUploadIdentitySha256",
      "newUploadIdentitySha256",
    ]) {
      const value = text(event[key]);
      if (value && !SHA256_RE.test(value)) {
        errors.push(`${base}.${key} must be SHA-256 when present`);
      }
    }
    if (event.event === "DOWNLOAD" && event.cacheDisposition !== "RETAINED") {
      errors.push(`${base} DOWNLOAD must retain cache`);
    }
    if (event.event === "DELETE") {
      if (
        event.invalidationAttempted !== true ||
        event.invalidationSucceeded !== true
      ) {
        errors.push(`${base} DELETE must successfully invalidate cache`);
      }
      if (event.cacheDisposition !== "INVALIDATED") {
        errors.push(`${base} DELETE cacheDisposition must be INVALIDATED`);
      }
    }
    if (event.event === "REUPLOAD") {
      const prior = text(event.priorUploadIdentitySha256);
      const next = text(event.newUploadIdentitySha256);
      if (!prior || !next)
        errors.push(`${base} REUPLOAD requires prior and new identity hashes`);
    }
  }

  return freezeDeep({ valid: errors.length === 0, errors });
}

function validatePricingPolicy(policy) {
  const errors = [];
  if (!isPlainObject(policy)) {
    return freezeDeep({
      valid: false,
      errors: ["pricingPolicy must be an object"],
    });
  }
  if (policy.version !== PRICING_POLICY_VERSION) {
    errors.push(`pricing policy version must be ${PRICING_POLICY_VERSION}`);
  }
  if (!text(policy.policyId)) errors.push("pricing policyId is required");
  if (policy.currency !== "USD") errors.push("pricing currency must be USD");
  if (!text(policy.mode)) errors.push("pricing mode is required");
  if (
    !isPlainObject(policy.models) ||
    Object.keys(policy.models).length === 0
  ) {
    errors.push("pricing models must be a non-empty object");
  }
  for (const [modelId, model] of Object.entries(policy.models || {})) {
    if (!text(modelId) || !isPlainObject(model)) {
      errors.push("pricing model entry must be an object");
      continue;
    }
    validateNonNegativeInteger(
      model.inputMicrousdPerMillionTokens,
      `models.${modelId}.inputMicrousdPerMillionTokens`,
      errors,
    );
    validateNonNegativeInteger(
      model.outputMicrousdPerMillionTokens,
      `models.${modelId}.outputMicrousdPerMillionTokens`,
      errors,
    );
  }
  for (const path of findSensitivePaths(policy)) {
    errors.push(`pricing policy contains forbidden sensitive field: ${path}`);
  }
  return freezeDeep({ valid: errors.length === 0, errors });
}

function validateThresholdPolicy(policy) {
  const errors = [];
  if (!isPlainObject(policy)) {
    return freezeDeep({
      valid: false,
      errors: ["thresholdPolicy must be an object"],
    });
  }
  if (policy.version !== THRESHOLD_POLICY_VERSION) {
    errors.push(`threshold policy version must be ${THRESHOLD_POLICY_VERSION}`);
  }
  if (!text(policy.policyId)) errors.push("threshold policyId is required");
  const minimum = policy.minimumSampleSize;
  if (!isPlainObject(minimum)) {
    errors.push("minimumSampleSize is required");
  } else {
    for (const key of [
      "executions",
      "coldExecutions",
      "warmExecutions",
      "lifecycleEvents",
    ]) {
      validateNonNegativeInteger(
        minimum[key],
        `minimumSampleSize.${key}`,
        errors,
      );
    }
  }
  validateNonNegativeInteger(
    policy.monthlyProjectionExecutions,
    "monthlyProjectionExecutions",
    errors,
  );
  const thresholds = policy.thresholds;
  if (!isPlainObject(thresholds)) {
    errors.push("thresholds are required");
  } else {
    const required = [
      "cacheHitRateMin",
      "warmCacheHitRateMin",
      "downloadReuseCacheHitRateMin",
      "providerCallRateMax",
      "warmProviderCallRateMax",
      "reuploadProviderCallRateMin",
      "p95LatencyMsMax",
      "p99LatencyMsMax",
      "warmP95LatencyMsMax",
      "cacheHitP95LatencyMsMax",
      "timeoutRateMax",
      "errorRateMax",
      "averageCostMicrousdMax",
      "providerCallAverageCostMicrousdMax",
      "warmAverageCostMicrousdMax",
      "monthlyProjectedCostMicrousdMax",
      "cacheCostAvoidanceRateMin",
      "deleteInvalidationCoverageMin",
      "downloadRetentionAccuracyMin",
      "reuploadIdentitySeparationAccuracyMin",
      "staleCacheReuseViolationCountMax",
    ];
    for (const key of required) {
      validatePositiveNumber(thresholds[key], `thresholds.${key}`, errors);
    }
  }
  return freezeDeep({ valid: errors.length === 0, errors });
}

function calculateProviderCostMicrousd(execution, pricingPolicy) {
  const provider = execution.provider || {};
  if (provider.called !== true) {
    return Object.freeze({
      costMicrousd: 0,
      source: "NO_PROVIDER_CALL",
      valid: true,
    });
  }
  if (
    Number.isInteger(provider.observedCostMicrousd) &&
    provider.observedCostMicrousd >= 0
  ) {
    return Object.freeze({
      costMicrousd: provider.observedCostMicrousd,
      source: "OBSERVED_COST",
      valid: true,
    });
  }
  const model = pricingPolicy.models?.[provider.modelId];
  if (!model) {
    return Object.freeze({
      costMicrousd: 0,
      source: "MISSING_MODEL_PRICE",
      valid: false,
      error: `pricing unavailable for model ${text(provider.modelId) || "UNKNOWN"}`,
    });
  }
  const inputCost =
    (provider.inputTokens * model.inputMicrousdPerMillionTokens) / 1_000_000;
  const outputCost =
    (provider.outputTokens * model.outputMicrousdPerMillionTokens) / 1_000_000;
  return Object.freeze({
    costMicrousd: Math.ceil(inputCost + outputCost),
    source: "TOKEN_PRICING_POLICY",
    valid: true,
  });
}

function phaseExecutions(executions, phases) {
  const allowed = new Set(Array.isArray(phases) ? phases : [phases]);
  return executions.filter((entry) => allowed.has(entry.phase));
}

function latencySummary(executions) {
  const values = executions
    .filter((entry) => entry.status === "SUCCESS")
    .map((entry) => entry.latencyMs);
  return Object.freeze({
    sampleCount: values.length,
    averageMs: mean(values),
    p50Ms: percentile(values, 0.5),
    p95Ms: percentile(values, 0.95),
    p99Ms: percentile(values, 0.99),
    maxMs: values.length ? Math.max(...values) : 0,
  });
}

function thresholdResult(metric, operator, actual, threshold) {
  const passed = operator === ">=" ? actual >= threshold : actual <= threshold;
  return Object.freeze({
    metric,
    operator,
    actual: round(actual),
    threshold,
    passed,
  });
}

function blockedReport({ dataset, thresholdPolicy, pricingPolicy, errors }) {
  return freezeDeep({
    version: REPORT_VERSION,
    evaluatorVersion: EVALUATOR_VERSION,
    decision: DECISIONS.BLOCKED,
    failClosed: true,
    validationErrors: [...errors],
    dataset: Object.freeze({
      datasetId: text(dataset?.datasetId),
      datasetSha256: isPlainObject(dataset) ? sha256(dataset) : "",
      benchmarkMode: text(dataset?.benchmarkMode),
    }),
    policies: Object.freeze({
      thresholdPolicyId: text(thresholdPolicy?.policyId),
      pricingPolicyId: text(pricingPolicy?.policyId),
    }),
    thresholdResults: Object.freeze([]),
    privacy: Object.freeze({
      rawRowsIncluded: false,
      rawFileNamesIncluded: false,
      userIdentityIncluded: false,
      storageKeysIncluded: false,
    }),
    guardrails: Object.freeze({
      evaluationOnly: true,
      routeWired: false,
      controllerWired: false,
      providerCallsExecutedByEvaluator: 0,
      promotionAuthorized: false,
      productionCandidateMergeApplied: false,
      productionReadyAssignment: false,
      railwayEnvironmentChanged: false,
    }),
  });
}

function evaluateCostCacheLatency({
  dataset,
  thresholdPolicy,
  pricingPolicy,
} = {}) {
  const datasetValidation = validateOperationalEvaluationDataset(dataset);
  const thresholdValidation = validateThresholdPolicy(thresholdPolicy);
  const pricingValidation = validatePricingPolicy(pricingPolicy);
  const validationErrors = [
    ...datasetValidation.errors,
    ...thresholdValidation.errors,
    ...pricingValidation.errors,
  ];
  if (validationErrors.length > 0) {
    return blockedReport({
      dataset,
      thresholdPolicy,
      pricingPolicy,
      errors: validationErrors,
    });
  }

  const executions = clone(dataset.executions);
  const lifecycleEvents = clone(dataset.lifecycleEvents);
  const costRecords = [];
  for (const execution of executions) {
    const cost = calculateProviderCostMicrousd(execution, pricingPolicy);
    if (!cost.valid)
      validationErrors.push(`${execution.executionId}: ${cost.error}`);
    costRecords.push({ executionId: execution.executionId, ...cost });
  }
  if (validationErrors.length > 0) {
    return blockedReport({
      dataset,
      thresholdPolicy,
      pricingPolicy,
      errors: validationErrors,
    });
  }

  const cold = phaseExecutions(executions, "COLD");
  const warm = phaseExecutions(executions, ["WARM", "DOWNLOAD_REUSE"]);
  const downloadReuse = phaseExecutions(executions, "DOWNLOAD_REUSE");
  const reupload = phaseExecutions(executions, "REUPLOAD");
  const cacheEligible = executions.filter(
    (entry) => entry.cache.readAttempted === true,
  );
  const cacheHits = cacheEligible.filter((entry) => entry.cache.hit === true);
  const cacheMisses = cacheEligible.filter((entry) => entry.cache.hit !== true);
  const providerCalls = executions.filter(
    (entry) => entry.provider.called === true,
  );
  const warmProviderCalls = warm.filter(
    (entry) => entry.provider.called === true,
  );
  const reuploadProviderCalls = reupload.filter(
    (entry) => entry.provider.called === true,
  );
  const successes = executions.filter((entry) => entry.status === "SUCCESS");
  const timeouts = executions.filter((entry) => entry.status === "TIMEOUT");
  const errors = executions.filter((entry) => entry.status === "ERROR");

  const costByExecutionId = new Map(
    costRecords.map((entry) => [entry.executionId, entry]),
  );
  const costs = executions.map(
    (entry) => costByExecutionId.get(entry.executionId).costMicrousd,
  );
  const providerCosts = providerCalls.map(
    (entry) => costByExecutionId.get(entry.executionId).costMicrousd,
  );
  const warmCosts = warm.map(
    (entry) => costByExecutionId.get(entry.executionId).costMicrousd,
  );
  const totalCostMicrousd = costs.reduce((sum, value) => sum + value, 0);
  const avoidedCostMicrousd = executions
    .filter(
      (entry) => entry.provider.called !== true && entry.cache.hit === true,
    )
    .reduce((sum, entry) => sum + entry.expectedColdCostMicrousd, 0);
  const totalPotentialCostMicrousd = totalCostMicrousd + avoidedCostMicrousd;
  const averageCostMicrousd = mean(costs);
  const monthlyProjectedCostMicrousd = Math.ceil(
    averageCostMicrousd * thresholdPolicy.monthlyProjectionExecutions,
  );

  const cacheLevelCounts = Object.fromEntries(
    ["L1", "L2", "L3", "L4"].map((level) => [
      level,
      cacheHits.filter((entry) => entry.cache.level === level).length,
    ]),
  );
  const cacheLevelRates = Object.fromEntries(
    Object.entries(cacheLevelCounts).map(([level, count]) => [
      level,
      safeRate(count, cacheHits.length),
    ]),
  );

  const deleteEvents = lifecycleEvents.filter(
    (event) => event.event === "DELETE",
  );
  const downloadEvents = lifecycleEvents.filter(
    (event) => event.event === "DOWNLOAD",
  );
  const reuploadEvents = lifecycleEvents.filter(
    (event) => event.event === "REUPLOAD",
  );
  const deleteInvalidated = deleteEvents.filter(
    (event) =>
      event.invalidationAttempted &&
      event.invalidationSucceeded &&
      event.cacheDisposition === "INVALIDATED",
  );
  const downloadRetained = downloadEvents.filter(
    (event) => event.cacheDisposition === "RETAINED",
  );
  const reuploadSeparated = reuploadEvents.filter((event) => {
    const prior = text(event.priorUploadIdentitySha256);
    const next = text(event.newUploadIdentitySha256);
    return prior && next && prior !== next && event.staleCacheReused !== true;
  });
  const staleCacheReuseViolationCount =
    executions.filter(
      (entry) => entry.lifecycleContext.staleCacheReused === true,
    ).length +
    lifecycleEvents.filter((event) => event.staleCacheReused === true).length;

  const cacheHitLatency = latencySummary(cacheHits);
  const overallLatency = latencySummary(executions);
  const coldLatency = latencySummary(cold);
  const warmLatency = latencySummary(warm);
  const reuploadLatency = latencySummary(reupload);

  const sample = Object.freeze({
    executions: executions.length,
    successfulExecutions: successes.length,
    coldExecutions: cold.length,
    warmExecutions: warm.length,
    downloadReuseExecutions: downloadReuse.length,
    reuploadExecutions: reupload.length,
    lifecycleEvents: lifecycleEvents.length,
  });

  const sampleResults = [
    thresholdResult(
      "sample.executions",
      ">=",
      sample.executions,
      thresholdPolicy.minimumSampleSize.executions,
    ),
    thresholdResult(
      "sample.coldExecutions",
      ">=",
      sample.coldExecutions,
      thresholdPolicy.minimumSampleSize.coldExecutions,
    ),
    thresholdResult(
      "sample.warmExecutions",
      ">=",
      sample.warmExecutions,
      thresholdPolicy.minimumSampleSize.warmExecutions,
    ),
    thresholdResult(
      "sample.lifecycleEvents",
      ">=",
      sample.lifecycleEvents,
      thresholdPolicy.minimumSampleSize.lifecycleEvents,
    ),
  ];

  const metrics = Object.freeze({
    cost: Object.freeze({
      currency: pricingPolicy.currency,
      totalMicrousd: totalCostMicrousd,
      averagePerExecutionMicrousd: averageCostMicrousd,
      averagePerProviderCallMicrousd: mean(providerCosts),
      warmAverageMicrousd: mean(warmCosts),
      p95PerExecutionMicrousd: percentile(costs, 0.95),
      avoidedByCacheMicrousd: avoidedCostMicrousd,
      cacheCostAvoidanceRate: safeRate(
        avoidedCostMicrousd,
        totalPotentialCostMicrousd,
      ),
      monthlyProjectionExecutions: thresholdPolicy.monthlyProjectionExecutions,
      monthlyProjectedCostMicrousd,
      sourceBreakdown: Object.freeze({
        observedCostCount: costRecords.filter(
          (entry) => entry.source === "OBSERVED_COST",
        ).length,
        tokenPricingPolicyCount: costRecords.filter(
          (entry) => entry.source === "TOKEN_PRICING_POLICY",
        ).length,
        noProviderCallCount: costRecords.filter(
          (entry) => entry.source === "NO_PROVIDER_CALL",
        ).length,
      }),
    }),
    cache: Object.freeze({
      eligibleReadCount: cacheEligible.length,
      hitCount: cacheHits.length,
      missCount: cacheMisses.length,
      hitRate: safeRate(cacheHits.length, cacheEligible.length),
      warmHitRate: safeRate(
        warm.filter((entry) => entry.cache.hit === true).length,
        warm.length,
      ),
      downloadReuseHitRate: safeRate(
        downloadReuse.filter((entry) => entry.cache.hit === true).length,
        downloadReuse.length,
      ),
      levelCounts: Object.freeze(cacheLevelCounts),
      levelRates: Object.freeze(cacheLevelRates),
    }),
    provider: Object.freeze({
      callCount: providerCalls.length,
      callRate: safeRate(providerCalls.length, executions.length),
      warmCallRate: safeRate(warmProviderCalls.length, warm.length),
      reuploadCallRate: safeRate(reuploadProviderCalls.length, reupload.length),
      inputTokens: providerCalls.reduce(
        (sum, entry) => sum + entry.provider.inputTokens,
        0,
      ),
      outputTokens: providerCalls.reduce(
        (sum, entry) => sum + entry.provider.outputTokens,
        0,
      ),
    }),
    latency: Object.freeze({
      overall: overallLatency,
      cold: coldLatency,
      warm: warmLatency,
      cacheHit: cacheHitLatency,
      reupload: reuploadLatency,
    }),
    reliability: Object.freeze({
      successRate: safeRate(successes.length, executions.length),
      timeoutCount: timeouts.length,
      timeoutRate: safeRate(timeouts.length, executions.length),
      errorCount: errors.length,
      errorRate: safeRate(errors.length, executions.length),
    }),
    lifecycle: Object.freeze({
      downloadEventCount: downloadEvents.length,
      deleteEventCount: deleteEvents.length,
      reuploadEventCount: reuploadEvents.length,
      downloadRetentionAccuracy: safeRate(
        downloadRetained.length,
        downloadEvents.length,
      ),
      deleteInvalidationCoverage: safeRate(
        deleteInvalidated.length,
        deleteEvents.length,
      ),
      reuploadIdentitySeparationAccuracy: safeRate(
        reuploadSeparated.length,
        reuploadEvents.length,
      ),
      staleCacheReuseViolationCount,
    }),
  });

  const t = thresholdPolicy.thresholds;
  const metricResults = [
    thresholdResult(
      "cache.hitRate",
      ">=",
      metrics.cache.hitRate,
      t.cacheHitRateMin,
    ),
    thresholdResult(
      "cache.warmHitRate",
      ">=",
      metrics.cache.warmHitRate,
      t.warmCacheHitRateMin,
    ),
    thresholdResult(
      "cache.downloadReuseHitRate",
      ">=",
      metrics.cache.downloadReuseHitRate,
      t.downloadReuseCacheHitRateMin,
    ),
    thresholdResult(
      "provider.callRate",
      "<=",
      metrics.provider.callRate,
      t.providerCallRateMax,
    ),
    thresholdResult(
      "provider.warmCallRate",
      "<=",
      metrics.provider.warmCallRate,
      t.warmProviderCallRateMax,
    ),
    thresholdResult(
      "provider.reuploadCallRate",
      ">=",
      metrics.provider.reuploadCallRate,
      t.reuploadProviderCallRateMin,
    ),
    thresholdResult(
      "latency.overall.p95Ms",
      "<=",
      metrics.latency.overall.p95Ms,
      t.p95LatencyMsMax,
    ),
    thresholdResult(
      "latency.overall.p99Ms",
      "<=",
      metrics.latency.overall.p99Ms,
      t.p99LatencyMsMax,
    ),
    thresholdResult(
      "latency.warm.p95Ms",
      "<=",
      metrics.latency.warm.p95Ms,
      t.warmP95LatencyMsMax,
    ),
    thresholdResult(
      "latency.cacheHit.p95Ms",
      "<=",
      metrics.latency.cacheHit.p95Ms,
      t.cacheHitP95LatencyMsMax,
    ),
    thresholdResult(
      "reliability.timeoutRate",
      "<=",
      metrics.reliability.timeoutRate,
      t.timeoutRateMax,
    ),
    thresholdResult(
      "reliability.errorRate",
      "<=",
      metrics.reliability.errorRate,
      t.errorRateMax,
    ),
    thresholdResult(
      "cost.averagePerExecutionMicrousd",
      "<=",
      metrics.cost.averagePerExecutionMicrousd,
      t.averageCostMicrousdMax,
    ),
    thresholdResult(
      "cost.averagePerProviderCallMicrousd",
      "<=",
      metrics.cost.averagePerProviderCallMicrousd,
      t.providerCallAverageCostMicrousdMax,
    ),
    thresholdResult(
      "cost.warmAverageMicrousd",
      "<=",
      metrics.cost.warmAverageMicrousd,
      t.warmAverageCostMicrousdMax,
    ),
    thresholdResult(
      "cost.monthlyProjectedCostMicrousd",
      "<=",
      metrics.cost.monthlyProjectedCostMicrousd,
      t.monthlyProjectedCostMicrousdMax,
    ),
    thresholdResult(
      "cost.cacheCostAvoidanceRate",
      ">=",
      metrics.cost.cacheCostAvoidanceRate,
      t.cacheCostAvoidanceRateMin,
    ),
    thresholdResult(
      "lifecycle.deleteInvalidationCoverage",
      ">=",
      metrics.lifecycle.deleteInvalidationCoverage,
      t.deleteInvalidationCoverageMin,
    ),
    thresholdResult(
      "lifecycle.downloadRetentionAccuracy",
      ">=",
      metrics.lifecycle.downloadRetentionAccuracy,
      t.downloadRetentionAccuracyMin,
    ),
    thresholdResult(
      "lifecycle.reuploadIdentitySeparationAccuracy",
      ">=",
      metrics.lifecycle.reuploadIdentitySeparationAccuracy,
      t.reuploadIdentitySeparationAccuracyMin,
    ),
    thresholdResult(
      "lifecycle.staleCacheReuseViolationCount",
      "<=",
      metrics.lifecycle.staleCacheReuseViolationCount,
      t.staleCacheReuseViolationCountMax,
    ),
  ];
  const thresholdResults = Object.freeze([...sampleResults, ...metricResults]);
  const decision = thresholdResults.every((entry) => entry.passed)
    ? DECISIONS.PASS
    : DECISIONS.BLOCKED;

  return freezeDeep({
    version: REPORT_VERSION,
    evaluatorVersion: EVALUATOR_VERSION,
    decision,
    failClosed: true,
    validationErrors: Object.freeze([]),
    dataset: Object.freeze({
      datasetId: dataset.datasetId,
      datasetSha256: sha256(dataset),
      benchmarkMode: dataset.benchmarkMode,
      executionCount: executions.length,
      lifecycleEventCount: lifecycleEvents.length,
    }),
    policies: Object.freeze({
      thresholdPolicyId: thresholdPolicy.policyId,
      thresholdPolicySha256: sha256(thresholdPolicy),
      pricingPolicyId: pricingPolicy.policyId,
      pricingPolicySha256: sha256(pricingPolicy),
      pricingMode: pricingPolicy.mode,
      currency: pricingPolicy.currency,
    }),
    sample,
    metrics,
    thresholdResults,
    failedThresholds: Object.freeze(
      thresholdResults
        .filter((entry) => !entry.passed)
        .map((entry) => entry.metric),
    ),
    privacy: Object.freeze({
      rawRowsIncluded: false,
      rawFileNamesIncluded: false,
      userIdentityIncluded: false,
      storageKeysIncluded: false,
      cacheSecretsIncluded: false,
    }),
    guardrails: Object.freeze({
      evaluationOnly: true,
      routeWired: false,
      controllerWired: false,
      providerCallsExecutedByEvaluator: 0,
      pricingPolicyIsInvoice: false,
      promotionAuthorized: false,
      productionCandidateMergeApplied: false,
      productionReadyAssignment: false,
      railwayEnvironmentChanged: false,
    }),
  });
}

module.exports = Object.freeze({
  EVALUATOR_VERSION,
  REPORT_VERSION,
  DATASET_VERSION,
  THRESHOLD_POLICY_VERSION,
  PRICING_POLICY_VERSION,
  DECISIONS,
  EXECUTION_PHASES,
  EXECUTION_STATUSES,
  CACHE_LEVELS,
  LIFECYCLE_EVENTS,
  canonicalJson,
  sha256,
  percentile,
  validateOperationalEvaluationDataset,
  validatePricingPolicy,
  validateThresholdPolicy,
  calculateProviderCostMicrousd,
  evaluateCostCacheLatency,
});
