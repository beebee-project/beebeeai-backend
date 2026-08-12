# Patch 15.3.2-F.1 — E-X Compatibility & Canonical Evaluation Input

## Purpose

Patch 15.3.2-F originally expected a finalized Patch E Real Shadow collection summary and a frozen Observation export.

The current canonical repository has no Patch E collection summary. Patch F.1 therefore provides a fail-closed compatibility path that does **not** recreate Patch E, does **not** enable the collector, and does **not** fabricate a collection window.

The compatibility baseline combines:

1. Patch 15.1 canonical deterministic Cost/Cache/Latency dataset.
2. Operator-approved actual provider pricing.
3. Patch 13.3 actual Live Provider Cache-Hit parity readiness evidence.
4. The existing Patch 15.1 evaluator and threshold policy.

## Critical methodology boundary

This is:

```text
CANONICAL_BENCHMARK_WITH_APPROVED_ACTUAL_PRICING
```

It is **not** current production traffic telemetry.

Interpretation:

```text
Provider pricing                 APPROVED_ACTUAL
Live CALLED → CACHE_HIT parity   ACTUAL evidence
Latency/cache distribution       Patch 15.1 canonical deterministic benchmark
Lifecycle distribution           Patch 15.1 canonical deterministic benchmark
Actual production telemetry      false
Canary promotion evidence        false
Production promotion evidence    false
```

Patch 15.1's source dataset contains synthetic benchmark `observedCostMicrousd` values. The Patch 15.1 evaluator gives observed cost precedence over token-based external pricing. F.1 therefore creates a derived private copy and removes `observedCostMicrousd` only from provider-called benchmark rows. All token, latency, cache and lifecycle fields are preserved.

The original Patch 15.1 dataset is never modified.

## Safety contract

```text
Patch E collector activation             NO
Patch E synthetic summary                NO
Invented from/to window                  NO
Provider call by F.1                     0
Internal Canary                          OFF
Production merge                         OFF
Production READY                         OFF
Production route                         unchanged
Promotion authorization                  false
Threshold auto-relaxation                NO
Patch 15.1 evaluator modification        NO
Patch 15.1 threshold modification        NO
```

If approved actual pricing causes the historical synthetic cost thresholds to fail, F.1 records that fact as threshold recalibration evidence. It does not lower the thresholds automatically.

## Apply

From backend root:

```powershell
$ErrorActionPreference = "Stop"

$patchZip = Join-Path `
  (Get-Location) `
  "query_candidate_patch15_3_2_F_1_e_x_compatibility_canonical_evaluation_input.zip"

Get-FileHash $patchZip -Algorithm SHA256

Expand-Archive `
  $patchZip `
  -DestinationPath . `
  -Force

Write-Host "PASS Patch 15.3.2-F.1 installed"
```

## Syntax

```powershell
Get-ChildItem `
  .\automation\queryCandidatePlanner*CanonicalEvaluation*.js,
  .\automation\queryCandidatePlannerCostCacheLatencyEvaluatorAdapter.js,
  .\scripts\queryCandidatePlanner*Canonical*.js |
  ForEach-Object {
    node --check $_.FullName
    if ($LASTEXITCODE -ne 0) {
      throw "Syntax failed: $($_.Name)"
    }
  }
```

## Patch F.1 QA

```powershell
$tests = @(
  Get-ChildItem `
    .\tests\queryCandidatePatch15_3_2_F_1*SmokeTest.js |
  Sort-Object Name
)

foreach ($test in $tests) {
  Write-Host "RUN $($test.Name)"
  node $test.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Patch 15.3.2-F.1 smoke failed: $($test.Name)"
  }
}

Write-Host "PASS Patch 15.3.2-F.1 smoke $($tests.Count)/$($tests.Count)"
```

## Private paths

The F-1 pricing file created previously is reused:

```powershell
$privateDir = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerPatch15_3_2_F.private"

$pricingFile = Join-Path `
  $privateDir `
  "queryCandidatePlannerApprovedActualPricingPolicy.private.json"

$canonicalInput = Join-Path `
  $privateDir `
  "queryCandidatePlannerCanonicalEvaluationInput.private.json"

$operationalReport = Join-Path `
  $privateDir `
  "queryCandidatePlannerCanonicalOperationalReport.private.json"

$baseline = Join-Path `
  $privateDir `
  "queryCandidatePlannerCanonicalEvaluationBaseline.private.json"
```

## Resolve canonical files

```powershell
$dataset = `
  ".\evaluation\queryCandidatePlannerOperationalEvaluationDataset.v1.json"

$thresholdPolicy = `
  ".\evaluation\queryCandidatePlannerOperationalThresholdPolicy.v1.json"

$readinessCandidates = @(
  ".\candidate-planner-live-cache-parity-readiness.json",
  ".\tests\fixtures\query-candidate-planner-shadow\call_required_group_avg_time_count\candidate-planner-live-cache-parity-readiness.json"
)

$readiness = $readinessCandidates |
  Where-Object { Test-Path $_ } |
  Select-Object -First 1

if (-not $readiness) {
  throw "Patch 13.3 live cache parity readiness evidence is missing"
}

foreach ($required in @(
  $dataset,
  $thresholdPolicy,
  $pricingFile,
  $readiness
)) {
  if (-not (Test-Path $required)) {
    throw "Required F.1 input missing: $required"
  }
}
```

## Inspect current evaluator

This does not execute the evaluator or a Provider call.

```powershell
node `
  .\scripts\queryCandidatePlannerInspectCostCacheLatencyEvaluator.js
```

Required:

```text
PASS cost/cache/latency evaluator inspection
WORKTREE_EQUALS_HEAD true
PROVIDER_CALLS_EXECUTED 0
```

## Prepare canonical compatibility input

```powershell
node `
  .\scripts\queryCandidatePlannerPrepareCanonicalEvaluationInput.js `
  --dataset $dataset `
  --pricing $pricingFile `
  --readiness $readiness `
  --output $canonicalInput

if ($LASTEXITCODE -ne 0) {
  throw "Canonical evaluation input preparation failed"
}
```

Required:

```text
PASS patch 15.3.2-F.1 canonical evaluation input prepared
MODE CANONICAL_BENCHMARK_WITH_APPROVED_ACTUAL_PRICING
EXECUTIONS >=25
LIFECYCLE >=15
APPROVED_ACTUAL_PRICING true
ACTUAL_LIVE_PROVIDER_PARITY_EVIDENCE true
ACTUAL_OPERATIONAL_TELEMETRY false
PATCH_E_SUMMARY_USED false
PROVIDER_CALLS_EXECUTED_BY_PREPARATION 0
PRODUCTION_PROMOTION_AUTHORIZED false
```

## Run existing Patch 15.1 evaluator

```powershell
node `
  .\scripts\queryCandidatePlannerRunCanonicalCostCacheLatencyEvaluation.js `
  --input $canonicalInput `
  --pricing $pricingFile `
  --readiness $readiness `
  --threshold-policy $thresholdPolicy `
  --report-output $operationalReport `
  --baseline-output $baseline

if ($LASTEXITCODE -ne 0) {
  throw "Canonical Cost/Cache/Latency evaluation failed"
}
```

Important possible outcomes:

```text
OPERATIONAL_DECISION EVALUATION_PASS
BASELINE_DECISION CANONICAL_EVALUATION_BASELINE_PASS
```

or:

```text
OPERATIONAL_DECISION EVALUATION_BLOCKED
BASELINE_DECISION CANONICAL_EVALUATION_BASELINE_COST_RECALIBRATION_REQUIRED
COST_THRESHOLD_RECALIBRATION_REQUIRED true
```

The second result is not patched around. It means approved actual pricing is materially different from the historical synthetic cost assumptions and the cost threshold requires a separate evidence-based recalibration step.

If the evaluator export changed after Patch 15.3.2 and the dynamic adapter cannot resolve it, the script fails closed with:

```text
BLOCKED EVALUATOR_EXPORT_UNRESOLVED
```

or:

```text
BLOCKED EVALUATOR_INVOCATION_UNRESOLVED
```

In that case run the evaluator inspection command and use the printed `EXPORTS` line to build a tiny invocation compatibility hotfix. Do not modify the evaluator itself.

## Inspect outputs

```powershell
$reportJson = Get-Content $operationalReport -Raw | ConvertFrom-Json
$baselineJson = Get-Content $baseline -Raw | ConvertFrom-Json

Write-Host "OPERATIONAL DECISION $($reportJson.decision)"
Write-Host "BASELINE DECISION $($baselineJson.decision)"
Write-Host "BASELINE SHA256 $($baselineJson.baselineSha256)"
Write-Host `
  "COST RECALIBRATION $($baselineJson.operationalEvaluation.costThresholdRecalibrationRequired)"
Write-Host `
  "NON-COST FAILURES $($baselineJson.operationalEvaluation.nonCostFailureCount)"
```

## Private-output guard

```powershell
node `
  .\scripts\queryCandidatePlannerAssertCanonicalEvaluationPrivateOutputsUntracked.js
```

## Completion interpretation

F.1 completes compatibility evaluation input when:

```text
Patch E summary                        NOT REQUIRED
Canonical Patch 15.1 dataset           VERIFIED
Approved actual pricing                VERIFIED
Actual Patch 13.3 live parity          VERIFIED
Derived input                          DETERMINISTIC
Source dataset mutation                false
Existing evaluator                     reused
Provider calls by F.1                  0
Actual operational telemetry           false
Production promotion authorized        false
Private output staging                 false
```

A PASS baseline is an evaluation reference only.

A cost recalibration result is also a valid F.1 finding; it requires a later explicit cost-threshold recalibration patch rather than silent threshold relaxation.
