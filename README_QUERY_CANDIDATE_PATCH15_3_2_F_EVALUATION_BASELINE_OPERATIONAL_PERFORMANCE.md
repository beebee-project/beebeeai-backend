# Patch 15.3.2-F — Evaluation Baseline & Operational Performance Evaluation

## Phase

PHASE 15.3-C — Evaluation & Promotion Evidence, step 1.

Predecessor: Patch 15.3.2-E Real Shadow Observation Collection.

Patch E must be finalized first. Its frozen summary `to` timestamp is the upper bound of the Patch F export. The collector remains OFF with its kill switch ON during this patch.

## Purpose

Patch F turns the completed Patch E real-shadow window into a reproducible operational evaluation baseline.

Flow:

```text
Patch E finalized summary
→ frozen Observation export
→ operator-approved actual provider pricing
→ existing Patch 15.3.2 Evidence Bundle Builder
→ existing Patch 15.1 Cost/Cache/Latency Evaluator
→ actual Operational Report
→ Patch F SHA-256 evaluation baseline
```

This patch does NOT modify the Patch 15.1 evaluator. It does NOT enable Internal Canary, Promotion Gate, Production Merge, Production Ready assignment, or Production Route.

## Patch F baseline contract

Required:

```text
Patch E READY_FOR_PATCH_15_3_2_F    true
Collection protocol                 complete
Execution sample                    >= 50
Lifecycle sample                    >= 20
Case coverage                       >= 10
Privacy violations                  0
Guardrail violations                0

Pricing mode                        APPROVED_ACTUAL
Pricing rate unit                   MICROUSD_PER_MILLION_TOKENS
Operator approval                   true
Input/output rates                  > 0
Production billing authority        false

Patch 15.1 Operational decision     EVALUATION_PASS
Operational sample                  >= 50
Evaluation only                     true
Promotion authorized                false
Production candidate merge          false
Production ready assignment         false
```

## New files

```text
automation/queryCandidatePlannerRealShadowEvaluationBaseline.js
automation/queryCandidatePlannerRealShadowEvaluationBaseline.schema.json
evaluation/queryCandidatePlannerRealShadowEvaluationBaselinePolicy.v1.json
scripts/queryCandidatePlannerPrepareApprovedActualPricingPolicy.js
scripts/queryCandidatePlannerVerifyRealShadowEvaluationPreflight.js
scripts/queryCandidatePlannerBuildRealShadowEvaluationBaseline.js
scripts/queryCandidatePlannerAssertRealShadowEvaluationPrivateOutputsUntracked.js
tests/queryCandidatePatch15_3_2_F*.js
PATCH_MANIFEST_PATCH15_3_2_F.json
PATCH_VALIDATION_PATCH15_3_2_F.json
```

The change is additive-only.

---

## 1. Apply

From the backend repository root:

```powershell
$ErrorActionPreference = "Stop"

Get-FileHash `
  .\query_candidate_patch15_3_2_F_evaluation_baseline_operational_performance.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_3_2_F_evaluation_baseline_operational_performance.zip `
  -DestinationPath . `
  -Force
```

## 2. Syntax check

```powershell
node --check `
  .\automation\queryCandidatePlannerRealShadowEvaluationBaseline.js

Get-ChildItem .\scripts\queryCandidatePlanner*RealShadow*Evaluation*.js |
  ForEach-Object { node --check $_.FullName }

Get-ChildItem .\tests\queryCandidatePatch15_3_2_F*.js |
  ForEach-Object { node --check $_.FullName }
```

## 3. Patch F QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch15_3_2_F*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    Write-Host "RUN $($_.Name)"
    node $_.FullName
    if ($LASTEXITCODE -ne 0) {
      throw "Patch 15.3.2-F smoke failed: $($_.Name)"
    }
  }
```

Expected: all Patch F smoke tests PASS.

## 4. Verify predecessor and protected evaluator

Patch F intentionally leaves the existing evaluator untouched.

```powershell
$expectedEvaluatorHash = `
  "67A0FF4D5AC83103D78C9172AA4CC072C008D195A22E99B3366F8440B9D8658C"

$actualEvaluatorHash = (
  Get-FileHash `
    .\automation\queryCandidatePlannerCostCacheLatencyEvaluator.js `
    -Algorithm SHA256
).Hash

if ($actualEvaluatorHash -ne $expectedEvaluatorHash) {
  throw "Protected Patch 15.1 evaluator drift detected"
}

node .\tests\queryCandidatePatch15_3_2_EManifestSmokeTest.js
```

If your Patch E package uses a differently named manifest smoke, run the Patch E source-integrity and manifest smoke that were installed with E.

## 5. Keep collector and production frozen

Railway must remain:

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=1

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

Do not change the existing evidence secret, registry JSON/SHA, allowlist, TTL, or MAX_RECORDS before export.

## 6. Prepare private paths

```powershell
$collectionSummary = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowObservationCollection.summary.private.json"

$privateDir = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerPatch15_3_2_F.private"

New-Item -ItemType Directory -Force -Path $privateDir | Out-Null

$recordsFile = Join-Path `
  $privateDir `
  "real-shadow-evidence-records.private.json"

$pricingFile = Join-Path `
  $privateDir `
  "queryCandidatePlannerApprovedActualPricingPolicy.private.json"

$bundleDir = Join-Path `
  $privateDir `
  "real-shadow-evidence-output.private"

$baselineFile = Join-Path `
  $privateDir `
  "queryCandidatePlannerRealShadowEvaluationBaseline.private.json"
```

## 7. Read the frozen Patch E window

```powershell
$summary = Get-Content $collectionSummary -Raw | ConvertFrom-Json

if (-not $summary.readyForPatch15_3_2_F) {
  throw "Patch E is not ready for F"
}

$from = [string]$summary.from
$to = [string]$summary.to

if ([string]::IsNullOrWhiteSpace($to)) {
  throw "Patch E export upper bound is missing"
}
```

If your finalized summary nests the window fields, use the exact `from` and `to` values printed/stored by Patch E. Never replace `to` with the current time: the finalized Patch E `to` is the immutable upper bound.

## 8. Prepare APPROVED_ACTUAL pricing

The old template is deliberately `DRAFT_NOT_APPROVED`; do not use it directly.

Confirm the provider/model price that is approved for this evaluation, then convert the price to:

```text
MICROUSD_PER_MILLION_TOKENS
```

Prepare the private policy:

```powershell
node `
  .\scripts\queryCandidatePlannerPrepareApprovedActualPricingPolicy.js `
  --policy-id "approved_actual_YYYY_MM_DD_v1" `
  --effective-at "YYYY-MM-DDTHH:mm:ss.fffZ" `
  --model-rate "semantic_profiler_default:<input-microusd-per-million>:<output-microusd-per-million>" `
  --approve true `
  --output $pricingFile
```

Do not invent or reuse stale rates. Add another `--model-rate` when the frozen records contain another model ID.

## 9. Preflight

Run with Railway variables available:

```powershell
railway run node `
  .\scripts\queryCandidatePlannerVerifyRealShadowEvaluationPreflight.js `
  --collection-summary $collectionSummary `
  --pricing $pricingFile
```

Expected:

```text
PASS patch 15.3.2-F evaluation preflight
EXECUTIONS >=50
LIFECYCLE >=20
CASES >=10
PRICING_MODE APPROVED_ACTUAL
COLLECTOR_FROZEN true
INTERNAL_CANARY_ENABLED false
PRODUCTION_PROMOTION_AUTHORIZED false
READY_FOR_FROZEN_EXPORT true
```

## 10. Frozen Observation export

```powershell
railway run node `
  .\scripts\queryCandidatePlannerExportRealShadowEvidence.js `
  --from $from `
  --to $to `
  --output $recordsFile
```

The exported file is private internal evaluation evidence and must not be committed.

## 11. Build the existing Evidence Bundle and run actual operational evaluation

Locate your valid Patch 13.3 readiness evidence:

```powershell
$readinessFile = `
  ".\candidate-planner-live-cache-parity-readiness.json"

if (-not (Test-Path $readinessFile)) {
  throw "Patch 13.3 readiness evidence is missing"
}
```

Build:

```powershell
node `
  .\scripts\queryCandidatePlannerBuildRealShadowEvidenceBundle.js `
  --records $recordsFile `
  --readiness $readinessFile `
  --pricing $pricingFile `
  --expires-hours 24 `
  --output-dir $bundleDir
```

Expected output includes:

```text
queryCandidatePlannerRealShadowOperationalReport.json
queryCandidatePlannerRealShadowOperationalDataset.json
queryCandidatePlannerRealShadowAccuracyReport.json
queryCandidatePlannerRealShadowEvaluationReport.json
queryCandidatePlannerInternalCanaryEvidenceBundle.json
```

The existing builder invokes the existing Patch 15.1 evaluator. The evaluator itself performs no provider calls.

## 12. Build the Patch F baseline

```powershell
$operationalReport = Join-Path `
  $bundleDir `
  "queryCandidatePlannerRealShadowOperationalReport.json"

node `
  .\scripts\queryCandidatePlannerBuildRealShadowEvaluationBaseline.js `
  --collection-summary $collectionSummary `
  --records $recordsFile `
  --pricing $pricingFile `
  --operational-report $operationalReport `
  --output $baselineFile
```

Expected:

```text
PASS patch 15.3.2-F evaluation baseline built
DECISION EVALUATION_BASELINE_PASS
BASELINE_SHA256 <64 hex>
EXECUTIONS >=50
LIFECYCLE >=20
CASES >=10
OPERATIONAL_DECISION EVALUATION_PASS
PRICING_MODE APPROVED_ACTUAL
EVALUATION_ONLY true
PRODUCTION_PROMOTION_AUTHORIZED false
PRIVATE_OUTPUT_DO_NOT_COMMIT true
```

## 13. Inspect operating metrics

```powershell
$baseline = Get-Content $baselineFile -Raw | ConvertFrom-Json

$baseline.operational.metrics |
  Format-List

Write-Host "BASELINE SHA256 $($baseline.baselineSha256)"
```

Evaluate the actual metrics using the existing Patch 15.1 threshold contract. Patch F does not silently relax those thresholds.

Important metrics include:

```text
Overall / Warm / Download cache hit
Provider call rate / Warm provider call rate / Reupload provider call rate
Latency p50 / p95 / p99
Timeout rate / Error rate
Average cost / Provider-call cost / Monthly projected cost
Cache cost avoidance
Download retention
Delete invalidation
Reupload identity separation
Stale cache reuse violations
```

If the existing Operational Report returns `EVALUATION_BLOCKED`, Patch F baseline generation must remain BLOCKED. Do not edit the report or lower thresholds to force PASS.

## 14. Private-output guard

```powershell
node `
  .\scripts\queryCandidatePlannerAssertRealShadowEvaluationPrivateOutputsUntracked.js

if ($LASTEXITCODE -ne 0) {
  throw "Patch F private-output guard failed"
}
```

## Patch F completion condition

```text
Patch E final summary                PASS
Collector after E                    OFF / kill-switch ON
Frozen export                        PASS
Actual pricing                       APPROVED_ACTUAL
Operational evaluation               EVALUATION_PASS
Execution sample                     >=50
Lifecycle sample                     >=20
Case coverage                        >=10
Privacy violations                   0
Guardrail violations                 0
Baseline SHA-256                     fixed
Private outputs                      untracked
Internal Canary                      OFF
Production Promotion                 BLOCKED
Production Merge / Route             unchanged
```

A Patch F `EVALUATION_BASELINE_PASS` is evaluation evidence only. It does not authorize production promotion.
