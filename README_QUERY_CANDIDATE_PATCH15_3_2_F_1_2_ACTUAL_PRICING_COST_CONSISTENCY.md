# Patch 15.3.2-F.1.2 — Actual Pricing Cache-Avoidance Cost Consistency

## Finding

Patch 15.3.2-F.1 produced:

```text
cost.totalMicrousd                    48900
cost.avoidedByCacheMicrousd            3960
cost.cacheCostAvoidanceRate         0.074915

cache.hitRate                           0.60
```

The absolute Provider cost was repriced with `APPROVED_ACTUAL`, but cache avoided cost still came from the historical synthetic `expectedColdCostMicrousd` scale.

That mixed two cost scales.

## Canonical dataset evidence

Each scenario contains one COLD execution with provider token counts and three cache-hit executions (two WARM plus one DOWNLOAD_REUSE).

With approved Terra pricing:

```text
scenario_01 COLD 1000 input / 200 output → 4400 microusd
scenario_02 COLD 1050 input / 210 output → 4620 microusd
scenario_03 COLD 1100 input / 220 output → 4840 microusd
scenario_04 COLD 1150 input / 230 output → 5060 microusd
scenario_05 COLD 1200 input / 240 output → 5280 microusd
```

Each COLD-equivalent cost is avoided three times:

```text
avoidedByCacheMicrousd
= 3 × (4400 + 4620 + 4840 + 5060 + 5280)
= 72600
```

The existing actual-pricing Provider cost is:

```text
48900 microusd
```

Therefore:

```text
cacheCostAvoidanceRate
= 72600 / (48900 + 72600)
= 0.597531
```

The existing threshold remains:

```text
cacheCostAvoidanceRateMin = 0.59
```

So the corrected result should PASS without relaxing the threshold.

## Patch boundary

F.1.2 does not modify:

```text
automation/queryCandidatePlannerCostCacheLatencyEvaluator.js
evaluation/queryCandidatePlannerOperationalThresholdPolicy.v1.json
Production routes
Promotion gate
Collector
```

It creates a private derived input and reprices `expectedColdCostMicrousd` using each scenario's COLD provider token counts and the existing APPROVED_ACTUAL pricing policy.

The source F.1 canonical input remains unchanged.

## Expected re-evaluation

Expected:

```text
cost.averagePerExecutionMicrousd        BLOCKED
cost.averagePerProviderCallMicrousd     BLOCKED
cost.monthlyProjectedCostMicrousd       BLOCKED

cost.cacheCostAvoidanceRate             PASS
actual                                  0.597531
threshold                               0.59

NON_COST_FAILURE_COUNT                  0
ABSOLUTE_COST_FAILURE_COUNT             3
```

This means cache avoided-cost consistency is fixed. The remaining three absolute cost ceilings can then be recalibrated separately.

## Apply

```powershell
$ErrorActionPreference = "Stop"

$patchZip = Join-Path `
  (Get-Location) `
  "query_candidate_patch15_3_2_F_1_2_actual_pricing_cost_consistency.zip"

Get-FileHash $patchZip -Algorithm SHA256

Expand-Archive `
  $patchZip `
  -DestinationPath . `
  -Force

Write-Host "PASS Patch 15.3.2-F.1.2 installed"
```

## Smoke tests

```powershell
$tests = @(
  Get-ChildItem `
    .\tests\queryCandidatePatch15_3_2_F_1_2*SmokeTest.js |
  Sort-Object Name
)

foreach ($test in $tests) {
  node $test.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Patch 15.3.2-F.1.2 smoke failed: $($test.Name)"
  }
}
```

## Reuse private F variables

```powershell
$privateDir = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerPatch15_3_2_F.private"

$pricingFile = Join-Path `
  $privateDir `
  "queryCandidatePlannerApprovedActualPricingPolicy.private.json"

$readiness = Join-Path `
  $privateDir `
  "queryCandidatePlannerPatch13_3HistoricalReadinessEvidence.private.json"

$canonicalInput = Join-Path `
  $privateDir `
  "queryCandidatePlannerCanonicalEvaluationInput.private.json"

$consistentInput = Join-Path `
  $privateDir `
  "queryCandidatePlannerActualPricingConsistentEvaluationInput.private.json"

$consistentReport = Join-Path `
  $privateDir `
  "queryCandidatePlannerActualPricingConsistentOperationalReport.private.json"

$consistentAssessment = Join-Path `
  $privateDir `
  "queryCandidatePlannerActualPricingCostConsistencyAssessment.private.json"

$consistentBaseline = Join-Path `
  $privateDir `
  "queryCandidatePlannerActualPricingConsistentEvaluationBaseline.private.json"

$thresholdPolicy = `
  ".\evaluation\queryCandidatePlannerOperationalThresholdPolicy.v1.json"
```

## Prepare repriced cache-avoidance input

```powershell
node `
  .\scripts\queryCandidatePlannerPrepareActualPricingConsistentEvaluationInput.js `
  --input $canonicalInput `
  --pricing $pricingFile `
  --output $consistentInput

if ($LASTEXITCODE -ne 0) {
  throw "Actual-pricing consistency input preparation failed"
}
```

Required:

```text
SCENARIOS 5
REPRICED_EXECUTIONS 25
REPRICED_CACHE_HITS 15
PROVIDER_COST_MICROUSD 48900
AVOIDED_BY_CACHE_MICROUSD 72600
PREDICTED_CACHE_COST_AVOIDANCE_RATE 0.597531
THRESHOLD_POLICY_MODIFIED false
EVALUATOR_MODIFIED false
PROVIDER_CALLS_EXECUTED 0
```

## Re-evaluate with unchanged threshold policy

```powershell
node `
  .\scripts\queryCandidatePlannerRunActualPricingCostConsistencyReevaluation.js `
  --input $consistentInput `
  --pricing $pricingFile `
  --readiness $readiness `
  --threshold-policy $thresholdPolicy `
  --report-output $consistentReport `
  --assessment-output $consistentAssessment `
  --baseline-output $consistentBaseline

if ($LASTEXITCODE -ne 0) {
  throw "Actual-pricing cost consistency re-evaluation failed"
}
```

Expected:

```text
OPERATIONAL_DECISION EVALUATION_BLOCKED

ASSESSMENT_DECISION
CACHE_AVOIDANCE_PRICING_CONSISTENCY_PASS_ABSOLUTE_COST_RECALIBRATION_REQUIRED

ABSOLUTE_COST_FAILURE_COUNT 3
CACHE_COST_AVOIDANCE_PASSED true
CACHE_COST_AVOIDANCE_ACTUAL 0.597531
CACHE_COST_AVOIDANCE_THRESHOLD 0.59
NON_COST_FAILURE_COUNT 0

THRESHOLD_POLICY_MODIFIED false
EVALUATOR_MODIFIED false
PROVIDER_CALLS_EXECUTED_BY_EVALUATOR 0
PRODUCTION_PROMOTION_AUTHORIZED false
```

Do not change the original threshold file during this patch.
