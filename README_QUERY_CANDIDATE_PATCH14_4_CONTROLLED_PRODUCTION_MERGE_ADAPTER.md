# Query Candidate Planner Patch 14.4

## Controlled Production Merge Adapter — Default OFF

Patch 14.4 adds a standalone adapter that can project the ranked Shadow Planner result into the existing `/analysis-candidates` candidate response contract. The adapter is **not imported by any route or controller**, does not modify the current Primary response, and cannot authorize a Production merge by itself.

## Prerequisite repository cleanup

The Patch 14.3.1 backup folder was committed and pushed before it was deleted locally. Commit and push the deletion before applying Patch 14.4.

```powershell
Remove-Item .\_patch14_3_1 -Recurse -Force -ErrorAction SilentlyContinue

# Add once to .gitignore
Add-Content .\.gitignore "`n_patch*/`n.patch_backups/"

git add -u .\.patch_backups
git add .\.gitignore
git commit -m "remove temporary patch backups"
git push
```

Confirm that the working tree is clean:

```powershell
git status
```

## Patch scope

New files only:

```text
automation/queryCandidatePlannerControlledProductionMergeAdapter.js
automation/queryCandidatePlannerControlledProductionMergeAdapter.schema.json
tests/queryCandidatePatch14_4*.js
PATCH_VALIDATION_PATCH14_4.json
PATCH_MANIFEST_PATCH14_4.json
README_QUERY_CANDIDATE_PATCH14_4_CONTROLLED_PRODUCTION_MERGE_ADAPTER.md
```

No existing route, controller, Shadow boundary, cache lifecycle module, or Internal Preview file is replaced.

## Adapter flow

```text
Primary candidate payload
+ ranked Shadow Planner resolution
+ Patch 14.0 Feature Control decision
+ Patch 13.3 readiness evidence
+ future Patch 14.4.1 Promotion Gate decision

→ candidate-contract projection
→ authorization evaluation
→ default BLOCKED
→ optional merged copy only when every guard is valid
```

The merged object is a newly allocated copy. The source Primary payload remains unchanged.

## Default behavior

```text
Feature flag default                         OFF
Production feature default                  OFF
Production candidate merge flag default     OFF
Production kill switch default              ON
Patch 14.4.1 gate installed                 NO
Route/controller wiring                     NO
Primary response mutation                   NO
Production READY assignment                 NO
Production route change                     NO
```

Even if an operator manually enables all Patch 14.0 Production flags, Patch 14.4 still blocks without the exact Promotion Gate decision contract that will be implemented in Patch 14.4.1.

Expected reason:

```text
PROMOTION_GATE_DECISION_REQUIRED
```

## Candidate contract projection

The adapter projects the Shadow ranking into these existing response fields:

```text
topCandidates
candidateUiPayload.recommendedCandidates
analysisRecipeCandidates
businessTemplateCandidates
multiSourceCandidates
categoryCandidates
dashboardCandidates
```

Unknown Shadow candidate fields are not copied. Raw rows, tenant identifiers, and other arbitrary payload fields are excluded. `READY`, `PRODUCTION_READY`, and `PROMOTED` status values are not projected, so this patch cannot assign Production readiness.

## Promotion authorization contract

A future Patch 14.4.1 decision must match all of the following:

```text
version        query_candidate_planner_controlled_production_promotion_gate_decision_v1
allowed        true
decision       ALLOW
operation      PRODUCTION_CANDIDATE_MERGE
failClosed     true
adapterVersion query_candidate_planner_controlled_production_merge_adapter_v1
```

Patch 14.4 does not create this decision and does not expose an API that accepts it.

## Apply

From the backend root:

```powershell
Get-FileHash `
  .\query_candidate_patch14_4_controlled_production_merge_adapter.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch14_4_controlled_production_merge_adapter.zip `
  -DestinationPath . `
  -Force
```

## Syntax check

```powershell
node --check .\automation\queryCandidatePlannerControlledProductionMergeAdapter.js

Get-ChildItem .\tests\queryCandidatePatch14_4*.js |
  ForEach-Object { node --check $_.FullName }
```

## Patch 14.4 QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch14_4*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object { node $_.FullName }
```

Expected PASS count: **15/15**.

## Compatibility checks

Patch 14.4 adds new files only, so predecessor manifests should remain unchanged.

```powershell
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3_1CumulativeIntegrityConvergenceSmokeTest.js
```

## Environment variables

Do not change Railway environment variables for Patch 14.4. The adapter remains unused and the Production kill switch remains ON.

## Non-goals

Patch 14.4 does not:

- merge any result into the live HTTP response;
- add or change an API route;
- alter the existing candidate popup;
- assign `READY` or any equivalent Production status;
- implement allowlists, rollout percentages, or deterministic sampling;
- activate the Promotion Gate;
- call an external Provider.

Allowlist, rollout, threshold evaluation, and the first gate implementation belong to Patch 14.4.1 and Phase 15.
