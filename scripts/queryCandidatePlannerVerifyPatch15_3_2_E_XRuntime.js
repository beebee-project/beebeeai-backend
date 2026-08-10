function bool(value, fallback = false) {
  if (value == null || String(value).trim() === "") return fallback;
  const v = String(value).trim().toLowerCase();
  if (["1", "true", "yes", "on"].includes(v)) return true;
  if (["0", "false", "no", "off"].includes(v)) return false;
  return fallback;
}

const env = process.env;
const evidenceRequested = bool(
  env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED,
  false,
);
const evidenceKillSwitch = bool(
  env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH,
  true,
);
const internalCanary = bool(
  env.QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED,
  false,
);
const production = bool(env.QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED, false);
const productionRoute = bool(
  env.QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED,
  false,
);
const promotionGate = bool(
  env.QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED,
  false,
);
const audience = String(
  env.QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE || "BLOCKED",
)
  .trim()
  .toUpperCase();
const rollout = Number(
  env.QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT || 0,
);

const errors = [];
if (evidenceRequested)
  errors.push("PATCH_E_X_EVIDENCE_COLLECTOR_MUST_BE_DISABLED");
if (!evidenceKillSwitch)
  errors.push("PATCH_E_X_EVIDENCE_KILL_SWITCH_MUST_BE_ENGAGED");
if (internalCanary)
  errors.push("PATCH_E_X_INTERNAL_CANARY_MUST_REMAIN_DISABLED");
if (production) errors.push("PATCH_E_X_PRODUCTION_MUST_REMAIN_DISABLED");
if (productionRoute)
  errors.push("PATCH_E_X_PRODUCTION_ROUTE_MUST_REMAIN_DISABLED");
if (promotionGate) errors.push("PATCH_E_X_PROMOTION_GATE_MUST_REMAIN_DISABLED");
if (audience !== "BLOCKED")
  errors.push("PATCH_E_X_PROMOTION_AUDIENCE_MUST_REMAIN_BLOCKED");
if (!Number.isFinite(rollout) || rollout !== 0)
  errors.push("PATCH_E_X_ROLLOUT_PERCENT_MUST_BE_ZERO");

if (errors.length) {
  for (const error of errors) console.error(`BLOCKED ${error}`);
  process.exit(1);
}

console.log("PASS patch 15.3.2-E-X runtime restore verification");
console.log("PATCH_15_3_2_E_STATUS EXCLUDED_NA");
console.log("EVIDENCE_COLLECTOR_ENABLED false");
console.log("EVIDENCE_KILL_SWITCH true");
console.log("INTERNAL_CANARY_ENABLED false");
console.log("PRODUCTION_ENABLED false");
console.log("PROMOTION_GATE_ENABLED false");
console.log("PROMOTION_AUDIENCE_MODE BLOCKED");
console.log("PROMOTION_ROLLOUT_PERCENT 0");
console.log("READY_FOR_PATCH_15_3_2_F true");
console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
