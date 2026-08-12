# Patch 15.3.2-F.1.1 — Historical Patch 13.3 Readiness Evidence Recovery

## Why this patch exists

Patch 15.3.2-F.1 requires Patch 13.3 Live Provider Cache-Hit Parity readiness evidence.

The current repository search returned:

```text
READINESS_FILE_COUNT 0
```

The historical Patch 13.3 run was already completed successfully. Re-running that live test would create a new paid Provider call and would no longer represent the exact historical evidence used by Patch 13.3.

F.1.1 therefore restores a **sanitized historical evidence capsule** from the previously verified Patch 13.3 evidence.

It does not claim to be new telemetry.

## Historical facts preserved

```text
source version               query_candidate_planner_live_cache_parity_readiness_evidence_v1
model                        gpt-5.6-terra

origin status                SHADOW_COMPLETED
origin invocation            CALLED
origin provider calls        1

replay status                SHADOW_COMPLETED
replay invocation            CACHE_HIT
replay provider calls        0
planner cache source         L3_SEMANTIC
reentry cache source         L4_REENTRY

parity valid                 true
observed provider calls      1
encrypted persistent files   3
plaintext persistent files   0

readiness eligible           true
promotion allowed            false
production route auto-wire   false
```

Original integrity references preserved in the capsule:

```text
Parity Audit SHA-256
9f231f6354d70a92b0461b930bf876fc0c723bc5c62eac8f26092ac28a54b5b2

Replay Audit SHA-256
77380cab79603663ec5cbed085c02e96af3274b0c84ae78343272078b9e77d66

Readiness Gate SHA-256
12fe722248ff2403a334ffbe735f97eec7cc52de7099a118c4144fd16d3e7823
```

## Deliberately excluded

The restored capsule excludes fields that Patch F.1 does not need:

```text
responseId
token usage values
raw rows
sample values
```

This keeps the recovery input minimal while preserving the historical security and parity facts needed by the F.1 validator.

## Safety

```text
Provider call by recovery       0
Collector activation            0
Current operational telemetry   false
Canary activation               false
Production promotion            false
Production route change         false
```

## Apply

From the backend repository root:

```powershell
$ErrorActionPreference = "Stop"

$patchZip = Join-Path `
  (Get-Location) `
  "query_candidate_patch15_3_2_F_1_1_historical_readiness_recovery.zip"

Get-FileHash `
  $patchZip `
  -Algorithm SHA256

Expand-Archive `
  $patchZip `
  -DestinationPath . `
  -Force

Write-Host "PASS Patch 15.3.2-F.1.1 installed"
```

## Syntax and smoke

```powershell
node --check `
  .\automation\queryCandidatePlannerHistoricalReadinessEvidenceRecovery.js

Get-ChildItem `
  .\scripts\queryCandidatePlanner*HistoricalReadiness*.js |
  ForEach-Object {
    node --check $_.FullName
    if ($LASTEXITCODE -ne 0) {
      throw "Syntax failed: $($_.Name)"
    }
  }

$tests = @(
  Get-ChildItem `
    .\tests\queryCandidatePatch15_3_2_F_1_1*SmokeTest.js |
  Sort-Object Name
)

foreach ($test in $tests) {
  node $test.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Smoke failed: $($test.Name)"
  }
}
```

## Restore to private F directory

Reuse the existing private directory:

```powershell
$privateDir = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerPatch15_3_2_F.private"

New-Item `
  -ItemType Directory `
  -Force `
  -Path $privateDir |
  Out-Null

$readiness = Join-Path `
  $privateDir `
  "queryCandidatePlannerPatch13_3HistoricalReadinessEvidence.private.json"
```

Create the capsule:

```powershell
node `
  .\scripts\queryCandidatePlannerRestoreHistoricalReadinessEvidence.js `
  --output $readiness

if ($LASTEXITCODE -ne 0) {
  throw "Historical readiness recovery failed"
}
```

No Provider API call occurs.

## Verify F.1 compatibility

```powershell
node `
  .\scripts\queryCandidatePlannerVerifyHistoricalReadinessF1Compatibility.js `
  --readiness $readiness

if ($LASTEXITCODE -ne 0) {
  throw "Historical readiness is not compatible with Patch F.1"
}
```

Required:

```text
PASS historical readiness is Patch 15.3.2-F.1 compatible
ORIGIN_PROVIDER_CALLS 1
REPLAY_PROVIDER_CALLS 0
PLANNER_CACHE_SOURCE L3_SEMANTIC
REENTRY_CACHE_SOURCE L4_REENTRY
PARITY_VALID true
ENCRYPTED_PERSISTENT_FILE_COUNT 3
PLAINTEXT_PERSISTENT_FILE_COUNT 0
READINESS_ELIGIBLE true
PROVIDER_CALLS_EXECUTED_BY_VERIFICATION 0
PRODUCTION_PROMOTION_AUTHORIZED false
```

## Continue Patch F.1

After compatibility PASS, use this `$readiness` path instead of searching for the missing fixture file.

```powershell
node `
  .\scripts\queryCandidatePlannerPrepareCanonicalEvaluationInput.js `
  --dataset $dataset `
  --pricing $pricingFile `
  --readiness $readiness `
  --output $canonicalInput
```

Then run the existing F.1 Cost/Cache/Latency evaluation.

## Private-output guard

```powershell
node `
  .\scripts\queryCandidatePlannerAssertHistoricalReadinessPrivateOutputUntracked.js
```

The recovered evidence file is private evaluation input and should not be committed.
