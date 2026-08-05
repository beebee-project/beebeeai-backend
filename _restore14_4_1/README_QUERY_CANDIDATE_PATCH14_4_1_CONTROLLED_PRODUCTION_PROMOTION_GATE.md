# Query Candidate Planner Patch 14.4.1

## Controlled Production Promotion Gate — Default BLOCKED

Patch 14.4.1 implements the standalone authorization decision required by the Patch 14.4 Controlled Production Merge Adapter. The gate evaluates Production safety controls, Patch 13.3 readiness evidence, an internal allowlist, and deterministic rollout assignment. It is **not imported by a route or controller**, so it cannot change the live `/analysis-candidates` response.

## Patch scope

New files only:

```text
automation/queryCandidatePlannerControlledProductionPromotionGate.js
automation/queryCandidatePlannerControlledProductionPromotionGate.schema.json
tests/queryCandidatePatch14_4_1*.js
PATCH_VALIDATION_PATCH14_4_1.json
PATCH_MANIFEST_PATCH14_4_1.json
README_QUERY_CANDIDATE_PATCH14_4_1_CONTROLLED_PRODUCTION_PROMOTION_GATE.md
```

No existing route, controller, Feature Control, Merge Adapter, Shadow boundary, cache lifecycle, or Internal Preview file is replaced.

## Decision order

```text
1. Operation and Merge Adapter version validation
2. Promotion Gate environment validation
3. Promotion Gate enabled flag
4. Patch 14.0 Feature Control
   - invalid Feature Control environment
   - Global Kill Switch
   - Planner feature flag
   - Production Kill Switch
   - Production feature flag
   - Production candidate merge flag
5. Patch 13.3 readiness evidence
6. Audience mode
7. Allowlist match
8. Deterministic rollout selection
```

Any failed check produces an immutable `BLOCK` decision with `failClosed=true`.

## Default behavior

```text
Promotion Gate enabled                         false
Audience mode                                  BLOCKED
Allowlist                                      empty
Rollout percentage                             0
Production Kill Switch                         unchanged; default ON
Route/controller wiring                        none
Primary API response mutation                  none
Production READY assignment                    none
Production route change                        none
```

Applying this patch does not activate Production merging.

## Audience modes

### `BLOCKED`

The gate remains blocked even when the gate flag is enabled. This is the default mode.

### `ALLOWLIST`

Only a request with a matching privacy-safe `subjectSha256` is allowed. The environment allowlist must contain only 64-character SHA-256 values. Raw email addresses, user IDs, tenant IDs, and filenames are rejected as invalid configuration.

This mode is intended for Patch 15.3 Internal Allowlist Canary.

### `ROLLOUT`

Allowlisted subjects are always selected so internal canary users remain available during rollout. Other subjects are assigned deterministically with:

```text
SHA256_MOD_10000_V1
```

The same subject and rollout salt always receive the same bucket from 0 to 9999. A rollout percentage of 1 selects buckets below 100, 5 selects buckets below 500, and 100 selects all valid subjects.

This mode is intended for Patch 15.4 staged rollout.

## Environment variables

Do **not** set these in Railway during Patch 14.4.1 QA.

```text
QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED
QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE
QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256
QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT
QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_SALT
```

Accepted audience modes:

```text
BLOCKED
ALLOWLIST
ROLLOUT
```

Strict configuration rules:

```text
Gate enabled                 0/1 or false/true only
Allowlist entries            comma/semicolon/space-separated SHA-256 only
Rollout percentage           integer 0 through 100
Rollout salt                 16 to 200 characters when ROLLOUT > 0
```

Invalid configuration is fail-closed.

## Privacy boundary

The decision does not expose:

```text
raw subjectSha256
raw allowlist values
rollout salt
email address
user ID
tenant ID
filename
```

It exposes only a second-level `subjectTagSha256`, allowlist count/match state, rollout percentage, and deterministic bucket.

## Patch 14.4 decision contract

An allowed decision matches the exact contract already required by the Merge Adapter:

```text
version        query_candidate_planner_controlled_production_promotion_gate_decision_v1
allowed        true
decision       ALLOW
operation      PRODUCTION_CANDIDATE_MERGE
failClosed     true
adapterVersion query_candidate_planner_controlled_production_merge_adapter_v1
```

The Merge Adapter can validate this decision, but remains unwired from all live routes. The QA integration uses `apply=false`, resulting in `DRY_RUN_READY` rather than a merged Production response.

## Apply

From the backend root:

```powershell
Get-FileHash `
  .\query_candidate_patch14_4_1_controlled_production_promotion_gate.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch14_4_1_controlled_production_promotion_gate.zip `
  -DestinationPath . `
  -Force
```

The ZIP has repository-root paths. It must create the module directly under `automation`, not under an extra package directory.

## Syntax check

```powershell
node --check `
  .\automation\queryCandidatePlannerControlledProductionPromotionGate.js

Get-ChildItem .\tests\queryCandidatePatch14_4_1*.js |
  ForEach-Object {
    node --check $_.FullName
  }
```

## Patch 14.4.1 QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch14_4_1*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    node $_.FullName
  }
```

Expected PASS count: **20/20**.

## Predecessor compatibility

```powershell
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3_1CumulativeIntegrityConvergenceSmokeTest.js
node .\tests\queryCandidatePatch14_4ManifestSmokeTest.js
```

Patch 14.4.1 adds files only, so predecessor files and manifests remain unchanged.

## Non-goals

Patch 14.4.1 does not:

- import the gate from a route or controller;
- merge a Shadow result into an HTTP response;
- change the candidate popup or Internal Preview page;
- assign `READY` or any Production-equivalent status;
- activate an allowlist or rollout in Railway;
- call an external Provider;
- define Phase 15 accuracy, cost, cache-hit, or latency thresholds.

The first real gate activation remains Patch 15.3 after Patch 15.0–15.2 evaluation gates pass.
