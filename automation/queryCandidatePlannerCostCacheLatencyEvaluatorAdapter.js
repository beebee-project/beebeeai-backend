const path = require("path");

const PREFERRED_EXPORTS = Object.freeze([
  "evaluateQueryCandidatePlannerCostCacheLatency",
  "evaluateQueryCandidatePlannerOperationalPerformance",
  "evaluateQueryCandidatePlannerOperationalDataset",
  "evaluateCostCacheLatency",
  "evaluateOperationalPerformance",
  "evaluateOperationalDataset",
  "evaluate",
]);

function evaluatorFunctions(moduleValue) {
  if (typeof moduleValue === "function") {
    return [{ name: "<module-function>", fn: moduleValue }];
  }
  const entries = Object.entries(moduleValue || {}).filter(
    ([, value]) => typeof value === "function",
  );

  const preferred = [];
  for (const name of PREFERRED_EXPORTS) {
    const match = entries.find(([key]) => key === name);
    if (match) preferred.push({ name: match[0], fn: match[1] });
  }

  const patterned = entries
    .filter(
      ([name]) =>
        !preferred.some((entry) => entry.name === name) &&
        /evaluat/i.test(name) &&
        /(cost|cache|latency|operational)/i.test(name),
    )
    .map(([name, fn]) => ({ name, fn }));

  return preferred.concat(patterned);
}

function reportLike(value) {
  return (
    value &&
    typeof value === "object" &&
    value.version ===
      "query_candidate_planner_cost_cache_latency_evaluation_report_v1"
  );
}

function invocationShapes({ dataset, pricingPolicy, thresholdPolicy }) {
  return [
    {
      label: "object-dataset-pricing-threshold",
      args: [{ dataset, pricingPolicy, thresholdPolicy }],
    },
    {
      label: "object-observations-pricing-policy",
      args: [
        {
          observations: dataset,
          pricingPolicy,
          policy: thresholdPolicy,
        },
      ],
    },
    {
      label: "object-evaluation-dataset",
      args: [
        {
          evaluationDataset: dataset,
          pricingPolicy,
          thresholdPolicy,
        },
      ],
    },
    {
      label: "positional-dataset-pricing-threshold",
      args: [dataset, pricingPolicy, thresholdPolicy],
    },
  ];
}

async function invokeExistingCostCacheLatencyEvaluator({
  dataset,
  pricingPolicy,
  thresholdPolicy,
  evaluatorModule = null,
} = {}) {
  const moduleValue =
    evaluatorModule ||
    require(
      path.resolve(
        "automation/queryCandidatePlannerCostCacheLatencyEvaluator.js",
      ),
    );

  const candidates = evaluatorFunctions(moduleValue);
  if (!candidates.length) {
    const keys =
      typeof moduleValue === "object"
        ? Object.keys(moduleValue || {}).sort()
        : [];
    const error = new Error(
      `No compatible evaluator export found. exports=${keys.join(",")}`,
    );
    error.code = "EVALUATOR_EXPORT_UNRESOLVED";
    throw error;
  }

  const failures = [];
  for (const candidate of candidates) {
    for (const shape of invocationShapes({
      dataset,
      pricingPolicy,
      thresholdPolicy,
    })) {
      try {
        const result = await Promise.resolve(candidate.fn(...shape.args));
        if (reportLike(result)) {
          return Object.freeze({
            exportName: candidate.name,
            invocationShape: shape.label,
            report: result,
          });
        }
        failures.push(`${candidate.name}/${shape.label}: non-report result`);
      } catch (error) {
        failures.push(
          `${candidate.name}/${shape.label}: ${error?.code || error?.message || error}`,
        );
      }
    }
  }

  const error = new Error(
    `Existing evaluator could not be invoked compatibly. ${failures
      .slice(0, 8)
      .join(" | ")}`,
  );
  error.code = "EVALUATOR_INVOCATION_UNRESOLVED";
  throw error;
}

module.exports = Object.freeze({
  PREFERRED_EXPORTS,
  evaluatorFunctions,
  reportLike,
  invocationShapes,
  invokeExistingCostCacheLatencyEvaluator,
});
