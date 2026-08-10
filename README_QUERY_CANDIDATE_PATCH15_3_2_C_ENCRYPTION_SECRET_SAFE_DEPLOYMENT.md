# Patch 15.3.2-C — Encryption Secret & Safe Deployment

## Phase

PHASE 15.3-B — Secure Evidence Collection, step 1 of 3.

Sequence:

1. 15.3.2-C Encryption Secret & Safe Deployment — this patch
2. 15.3.2-D Limited Evidence Collector Activation
3. 15.3.2-E Real Shadow Observation Collection

Patch C does **not** activate the collector. It prepares and verifies the secret, finalized real-shadow registry, internal subject allowlist, encryption round-trip, and production safety state before Patch D is allowed.

## Prerequisite

PHASE 15.3-A must have been finalized successfully. The following private outputs are expected locally and must not be committed:

- `queryCandidatePlannerRealShadowEvidenceFoundation.summary.private.json`
- `queryCandidatePlannerRealShadowCaseRegistry.private.json`
- `queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt`

The foundation summary must contain:

- `decision = REAL_SHADOW_EVIDENCE_FOUNDATION_PASS`
- `readyForPatch15_3_2_C = true`
- full case coverage
- expected-rejection evidence coverage
- `productionPromotionAuthorized = false`

## Security contract

Patch C requires all of the following before it returns `READY_FOR_PATCH_15_3_2_D true`:

- evidence secret is exactly 64 Base64URL characters generated from 48 random bytes;
- the raw secret is never printed by the new secret-generation CLI;
- the secret is stored only in a `.private.env` file and Railway secret variable;
- collector remains disabled (`QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0`);
- finalized registry JSON matches the Phase 15.3-A registry SHA-256;
- deployed runtime registry JSON matches `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_SHA256`;
- internal subject allowlist is non-empty;
- AES-256-GCM encryption round-trip succeeds;
- the wrong secret cannot decrypt the self-test envelope;
- production merge, route, ready-assignment, promotion gate, rollout, and internal canary remain blocked/off;
- no operation in this patch authorizes production promotion.

## New files

- `automation/queryCandidatePlannerRealShadowSecureDeployment.js`
- `scripts/queryCandidatePlannerCreateRealShadowEvidenceSecretPrivate.js`
- `scripts/queryCandidatePlannerVerifyRealShadowSecureDeployment.js`
- `scripts/queryCandidatePlannerVerifyRealShadowSecureRuntime.js`
- `scripts/queryCandidatePlannerAssertRealShadowSecureDeploymentPrivateOutputsUntracked.js`
- Patch 15.3.2-C smoke tests

The patch is additive-only. It does not modify routes or collector runtime behavior.

## 1. Generate the secret without logging it

Run:

```powershell
$secretFile = Join-Path (Get-Location) "queryCandidatePlannerRealShadowEvidenceSecret.private.env"

node .\scripts\queryCandidatePlannerCreateRealShadowEvidenceSecretPrivate.js --output $secretFile

if ($LASTEXITCODE -ne 0) {
  throw "Secret generation failed"
}
```

Expected output includes only the secret SHA-256, not the raw secret:

```text
PASS real shadow evidence secret private file created
SECRET_SHA256 <64 hex>
ENTROPY_BYTES 48
SECRET_FORMAT BASE64URL_64
RAW_SECRET_LOGGED false
COLLECTOR_ENABLED_BY_THIS_OPERATION false
PRIVATE_OUTPUT_DO_NOT_COMMIT true
```

The generator fails if the file already exists. `--force` is required for an intentional rotation.

## 2. Load the secret into the current PowerShell session without printing it

```powershell
$secretLine = Get-Content $secretFile | Where-Object {
  $_ -like "QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET=*" -and
  $_ -notlike "*SECRET_SHA256=*"
} | Select-Object -First 1

if (-not $secretLine) {
  throw "Secret line missing"
}

$evidenceSecret = $secretLine.Substring($secretLine.IndexOf("=") + 1)

if ($evidenceSecret -cnotmatch "^[A-Za-z0-9_-]{64}$") {
  throw "Secret format invalid"
}

$env:QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET = $evidenceSecret
$env:QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED = "0"
```

Do not use `Write-Host $evidenceSecret` and do not paste the secret into chat or logs.

## 3. Load the finalized registry and registry SHA-256

```powershell
$registryFile = Join-Path (Get-Location) "queryCandidatePlannerRealShadowCaseRegistry.private.json"
$foundationSummaryFile = Join-Path (Get-Location) "queryCandidatePlannerRealShadowEvidenceFoundation.summary.private.json"

$registryObject = Get-Content $registryFile -Raw | ConvertFrom-Json
$registryJson = $registryObject | ConvertTo-Json -Depth 30 -Compress
$foundationSummary = Get-Content $foundationSummaryFile -Raw | ConvertFrom-Json

$env:QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON = $registryJson
$env:QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_SHA256 = [string]$foundationSummary.registrySha256
```

## 4. Internal allowlist

`QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256` must contain at least one SHA-256 internal subject hash. Reuse the same immutable internal account/tenant subject hashing contract already used by Patch 15.3. Do not put an email address, name, JWT, or raw account identifier in the allowlist variable.

Example generation for an immutable internal account ID:

```powershell
$subjectHash = node .\scripts\queryCandidatePlannerCanarySubjectHash.js "<immutable-internal-account-id>"

if ($LASTEXITCODE -ne 0) {
  throw "Subject hash generation failed"
}

if ($subjectHash -cnotmatch "^[a-fA-F0-9]{64}$") {
  throw "Subject hash invalid"
}

$env:QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256 = $subjectHash
```

Do not share the raw account identifier.

## 5. Local secure-deployment preflight

```powershell
$verifyArgs = @(
  ".\scripts\queryCandidatePlannerVerifyRealShadowSecureDeployment.js",
  "--foundation-summary", $foundationSummaryFile,
  "--registry", $registryFile
)

& node @verifyArgs

if ($LASTEXITCODE -ne 0) {
  throw "Patch 15.3.2-C secure deployment preflight failed"
}
```

Expected:

```text
PASS patch 15.3.2-C secure deployment preflight
SECRET_SHA256 <64 hex>
REGISTRY_SHA256 <64 hex>
ALLOWLIST_ENTRIES >=1
ENCRYPTION_ROUND_TRIP true
WRONG_SECRET_REJECTED true
COLLECTOR_ENABLED false
READY_FOR_PATCH_15_3_2_D true
RAW_SECRET_LOGGED false
PRODUCTION_PROMOTION_AUTHORIZED false
```

## 6. Railway variables for Patch C

Configure these on the `beebeeai-backend` service. Keep the collector disabled:

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET=<secret from private env file>
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON=<finalized registry JSON>
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_SHA256=<foundation registry SHA-256>
QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256=<internal subject hash(es)>
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_TTL_DAYS=7
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_MAX_RECORDS=5000
```

Do not change the existing fail-closed production/canary state during Patch C:

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

Redeploy after changing Railway variables.

## 7. Deployed runtime verification

After the patch is deployed, run the runtime verifier inside the deployed service environment (for example through a Railway shell/CLI session where the service environment variables are present):

```powershell
node .\scripts\queryCandidatePlannerVerifyRealShadowSecureRuntime.js
```

Expected:

```text
PASS patch 15.3.2-C secure runtime verification
REGISTRY_CASES 10
ALLOWLIST_ENTRIES >=1
ENCRYPTION_ROUND_TRIP true
WRONG_SECRET_REJECTED true
COLLECTOR_ENABLED false
READY_FOR_PATCH_15_3_2_D true
RAW_SECRET_LOGGED false
PRODUCTION_PROMOTION_AUTHORIZED false
```

Patch D must not start unless both local preflight and deployed runtime verification return `READY_FOR_PATCH_15_3_2_D true`.

## 8. Private output guard

```powershell
node .\scripts\queryCandidatePlannerAssertRealShadowSecureDeploymentPrivateOutputsUntracked.js

if ($LASTEXITCODE -ne 0) {
  throw "Secure-deployment private output guard failed"
}
```

Expected:

```text
PASS no secure-deployment private outputs staged
```

## Exit gate

PHASE 15.3-B remains in Patch C until all are true:

```text
PHASE_15_3_A_FOUNDATION_PASS=true
SECRET_FORMAT_VALID=true
ENCRYPTION_ROUND_TRIP=true
WRONG_SECRET_REJECTED=true
FINAL_REGISTRY_HASH_MATCH=true
RUNTIME_REGISTRY_HASH_MATCH=true
INTERNAL_ALLOWLIST_CONFIGURED=true
COLLECTOR_ENABLED=false
INTERNAL_CANARY_ENABLED=false
PRODUCTION_PROMOTION_AUTHORIZED=false
PRIVATE_OUTPUTS_UNTRACKED=true
READY_FOR_PATCH_15_3_2_D=true
```

Only then proceed to Patch 15.3.2-D — Limited Evidence Collector Activation.
