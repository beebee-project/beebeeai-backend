# Patch 15.3.2-D — Limited Evidence Collector Activation

## Phase
PHASE 15.3-B — Secure Evidence Collection, step 2 of 3.

Sequence:
1. 15.3.2-C Encryption Secret & Safe Deployment — CLOSED prerequisite
2. 15.3.2-D Limited Evidence Collector Activation — this patch
3. 15.3.2-E Real Shadow Observation Collection — next

**Do not start PHASE 15.3-C Evaluation yet.** Patch E must collect the real encrypted evidence first.

## Safety contract
Patch D activates only the already-wired real-shadow evidence collector. It does not enable Internal Canary, production merge, production route, production ready assignment, promotion gate, or rollout.

Limited activation requires:
- Patch C secure runtime baseline still passes;
- exactly one allowlisted internal subject hash;
- finalized 10-case registry;
- evidence secret and pinned registry SHA remain valid;
- TTL is 1–7 days;
- max records is 30–5000;
- production/canary flags stay fail-closed;
- new collector kill switch is active before activation;
- activation is performed by changing only the evidence enable and kill-switch variables.

## New/updated files
- UPDATED `automation/queryCandidatePlannerRealShadowEvidenceConfig.js`
  - adds `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH`
  - `enabled` becomes effective state: requested enabled AND kill switch released
- NEW `automation/queryCandidatePlannerRealShadowLimitedActivation.js`
- NEW preflight/runtime verifier CLIs
- NEW private-output guard
- Patch D smoke tests

## Step 1 — deploy Patch D with collector still OFF
Set Railway:
```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=1
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_TTL_DAYS=7
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_MAX_RECORDS=5000
```
Keep all production/canary safety variables unchanged.

Redeploy, then run inside the Railway service environment:
```powershell
node .\scripts\queryCandidatePlannerVerifyRealShadowLimitedActivationPreflight.js
```
Expected:
```text
PASS patch 15.3.2-D limited collector activation preflight
REGISTRY_CASES 10
ALLOWLIST_ENTRIES 1
TTL_DAYS 7
MAX_RECORDS 5000
COLLECTOR_ENABLED false
COLLECTOR_KILL_SWITCH true
READY_FOR_LIMITED_ACTIVATION true
PRODUCTION_PROMOTION_AUTHORIZED false
```

## Step 2 — activate the limited collector
Change only:
```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=1
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=0
```
Redeploy.

Do not alter:
```text
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED=0
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH=1
QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE=BLOCKED
QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH=1
```

Then verify:
```powershell
node .\scripts\queryCandidatePlannerVerifyRealShadowLimitedActivationRuntime.js
```
Expected:
```text
PASS patch 15.3.2-D limited collector runtime verification
REGISTRY_CASES 10
ALLOWLIST_ENTRIES 1
TTL_DAYS 7
MAX_RECORDS 5000
COLLECTOR_ENABLED true
COLLECTOR_KILL_SWITCH false
READY_FOR_PATCH_15_3_2_E true
INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false
PRODUCTION_PROMOTION_AUTHORIZED false
```

## Immediate rollback
Fastest rollback:
```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=1
```
Redeploy. Effective collector state becomes OFF even if `...EVIDENCE_ENABLED=1` remains.

Full rollback:
```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=1
```

## Exit gate
Patch D closes only when deployed runtime verification returns `READY_FOR_PATCH_15_3_2_E true`. Then Patch E may collect actual encrypted observations. PHASE 15.3-C remains BLOCKED until Patch E is completed.
