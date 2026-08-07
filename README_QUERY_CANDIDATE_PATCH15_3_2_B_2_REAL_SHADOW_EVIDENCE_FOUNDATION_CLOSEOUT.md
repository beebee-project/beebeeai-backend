# Patch 15.3.2-B.2 — Real Shadow Evidence Foundation Closeout Gate

## Phase

**PHASE 15.3-A — Real Shadow Evidence Foundation**의 종료 패치다.

Patch 15.3.2-B/B.1에서 실제 업로드 가능한 Source Catalog와 source-bound v2 Fingerprint Ledger를 만들었지만, `seed_unstructured_unsupported`는 제품상 후보군을 생성하지 않는 것이 정상인 Case다. 이 패치는 해당 Case를 일반적인 "후보 생성 실패"로 취급하지 않고, Patch 15.0 Accuracy Dataset의 `unsupported.expectedRejected=true` 계약에 맞는 **Expected Rejection Evidence**로 명시적으로 검증한다.

이 패치는 Production runtime, API route, UI, Feature Flag, Railway 변수, Evidence Collector 상태를 변경하지 않는다.

## Decision

`seed_unstructured_unsupported`의 Foundation Evidence 인정 조건:

- Patch 15.0에서 `unsupported.expectedRejected=true`
- Source Catalog에 결합된 실제 업로드 원본
- 실제 플랫폼 실행에서 생성된 Request Fingerprint 64-hex
- 실제 플랫폼 실행에서 생성된 Upload Fingerprint 64-hex
- Request/Upload Fingerprint가 서로 다름
- Observation status가 `COMPLETED` 또는 `COMPLETED_SAFE`
- Observation reason이 비어 있지 않음
- `shadow.accepted === 0`
- Capture source가 `INTERNAL_PREVIEW` 또는 `API_SHADOW_OBSERVATION`
- raw candidate payload, raw file content, raw identity를 Attestation에 저장하지 않음

따라서 "생성 가능한 업무 템플릿 후보가 없습니다." UI는 이 Case에서 자동 실패 조건이 아니다. 실제 Shadow가 안전하게 완료되고 후보를 0개 수락한 증거가 있어야 한다.

## Safety state

- Evidence Collector: **OFF 유지**
- Internal Allowlist Canary: **OFF 유지**
- Production Merge: **OFF 유지**
- Production Promotion: **BLOCKED 유지**
- Railway 변수 자동 변경: 없음
- JWT / Internal Preview Token 저장: 없음
- 실제 원본 파일 포함: 없음
- Private Ledger/Catalog/Attestation 포함: 없음

## Added components

- `automation/queryCandidatePlannerRealShadowEvidenceFoundation.js`
  - Expected Rejection Attestation 생성/검증
  - 10/10 Source Catalog + 10/10 v2 Ledger + Expected Rejection Evidence를 하나의 Foundation Gate로 평가
  - `READY_FOR_PATCH_15_3_2_C` 판정
- `scripts/queryCandidatePlannerRecordRealShadowExpectedRejection.js`
  - Expected-Rejection Case의 Fingerprint 기록과 Attestation을 한 번에 생성
  - Case ID/Fingerprint/Source binding 오류 fail-closed
- `scripts/queryCandidatePlannerFinalizeRealShadowEvidenceFoundation.js`
  - Real Shadow Registry와 PHASE 15.3-A Foundation Summary를 동시에 최종화
- `scripts/queryCandidatePlannerAssertRealShadowFoundationPrivateOutputsUntracked.js`
  - 기존 Private 파일 + Expected Rejection Attestation + Foundation Summary Git staging 차단

Patch 15.3.2-B.1 파일은 수정하지 않는다. 따라서 B.1 Source Integrity와 Manifest 계약을 그대로 유지한다.

---

# 1. Apply

```powershell
$ErrorActionPreference = "Stop"

Get-FileHash `
  .\query_candidate_patch15_3_2_B_2_real_shadow_evidence_foundation_closeout.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_3_2_B_2_real_shadow_evidence_foundation_closeout.zip `
  -DestinationPath . `
  -Force
```

# 2. Syntax check

```powershell
node --check `
  .\automation\queryCandidatePlannerRealShadowEvidenceFoundation.js

node --check `
  .\scripts\queryCandidatePlannerRecordRealShadowExpectedRejection.js

node --check `
  .\scripts\queryCandidatePlannerFinalizeRealShadowEvidenceFoundation.js

node --check `
  .\scripts\queryCandidatePlannerAssertRealShadowFoundationPrivateOutputsUntracked.js
```

# 3. Patch 15.3.2-B.2 QA

```powershell
$tests = Get-ChildItem `
  .\tests\queryCandidatePatch15_3_2_B_2*SmokeTest.js |
  Sort-Object Name

foreach ($test in $tests) {
  Write-Host "RUN $($test.Name)"
  node $test.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Patch 15.3.2-B.2 QA failed: $($test.Name)"
  }
}

Write-Host "PASS Patch 15.3.2-B.2 QA $($tests.Count)/$($tests.Count)"
```

Expected: **14/14 PASS**.

# 4. Predecessor regression

B.1:

```powershell
$tests = Get-ChildItem `
  .\tests\queryCandidatePatch15_3_2_B_1*SmokeTest.js |
  Sort-Object Name

foreach ($test in $tests) {
  node $test.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Patch 15.3.2-B.1 regression failed: $($test.Name)"
  }
}
```

Expected: **17/17 PASS**.

B:

```powershell
$tests = Get-ChildItem `
  .\tests\queryCandidatePatch15_3_2_B*SmokeTest.js |
  Where-Object {
    $_.Name -notlike "queryCandidatePatch15_3_2_B_1*" -and
    $_.Name -notlike "queryCandidatePatch15_3_2_B_2*"
  } |
  Sort-Object Name

foreach ($test in $tests) {
  node $test.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Patch 15.3.2-B regression failed: $($test.Name)"
  }
}
```

Expected: **12/12 PASS**.

---

# 5. Current 9/10 migration — final unsupported Case

현재 v2 Ledger가 다음 상태라면:

```text
READY 9 cases
PENDING seed_unstructured_unsupported
PROGRESS 9/10
REMAINING 1
COMPLETE false
LEGACY_LEDGER_ACCEPTED false
```

기존 9개를 다시 수집하지 않는다. 마지막 Expected-Rejection Case만 실제 플랫폼에서 다시 실행한다.

## 5.1 Variables

```powershell
$caseId = "seed_unstructured_unsupported"

$sourceFile = Join-Path `
  (Get-Location) `
  ".local_uploads\서식중심_양식형_쿼리화_테스트.xlsx"

$sourceCatalog = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowUploadableSourceCatalog.private.json"

$ledger = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowFingerprintLedger.v2.private.json"

$attestation = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowExpectedRejectionAttestation.private.json"
```

`$sourceFile`은 Source Catalog에 해당 Case로 실제 결합된 파일과 정확히 동일해야 한다.

## 5.2 Capture

Internal Preview Baseline을 먼저 확보하고, 그 다음 실제 BeeBee AI 플랫폼에 위 파일을 업로드한다. UI에서 다음 메시지가 나오는 것은 이 Case에서는 정상일 수 있다.

```text
생성 가능한 업무 템플릿 후보가 없습니다.
```

반드시 Baseline 이후의 **fresh Observation**을 `$fresh`로 선택한다.

Fingerprint와 Outcome metadata를 확인한다. Fingerprint 원문은 출력하지 않는다.

```powershell
$requestFingerprint = `
  [string]$fresh.requestFingerprintSha256

$uploadFingerprint = `
  [string]$fresh.cacheLifecycle.identity.uploadFingerprintSha256

$observationStatus = `
  [string]$fresh.status

$observationReason = `
  [string]$fresh.reason

$shadowAccepted = `
  [int]$fresh.shadow.accepted

Write-Host "OBSERVED AT       $($fresh.observedAt)"
Write-Host "STATUS            $observationStatus"
Write-Host "REASON            $observationReason"
Write-Host "SHADOW ACCEPTED   $shadowAccepted"
Write-Host "REQUEST FP LENGTH $($requestFingerprint.Length)"
Write-Host "UPLOAD FP LENGTH  $($uploadFingerprint.Length)"
```

Expected example:

```text
STATUS            COMPLETED_SAFE
REASON            SKIPPED
SHADOW ACCEPTED   0
REQUEST FP LENGTH 64
UPLOAD FP LENGTH  64
```

`FAILED_SAFE`, `TIMEOUT_SAFE`, 64자리 미만/초과, `shadowAccepted > 0`이면 기록하지 않는다.

## 5.3 Record expected rejection + attestation

```powershell
node `
  .\scripts\queryCandidatePlannerRecordRealShadowExpectedRejection.js `
  --ledger $ledger `
  --source-catalog $sourceCatalog `
  --source-file $sourceFile `
  --case-id $caseId `
  --request-fingerprint $requestFingerprint `
  --upload-fingerprint $uploadFingerprint `
  --capture-source "INTERNAL_PREVIEW" `
  --observation-status $observationStatus `
  --observation-reason $observationReason `
  --shadow-accepted $shadowAccepted `
  --observed-at "$($fresh.observedAt)" `
  --attestation-output $attestation

if ($LASTEXITCODE -ne 0) {
  throw "Expected rejection evidence record failed"
}
```

Expected:

```text
PASS recorded expected-rejection case=seed_unstructured_unsupported
PROGRESS 10/10
REMAINING 0
EXPECTED_REJECTION_VERIFIED true
SHADOW_ACCEPTED 0
RAW_FINGERPRINTS_LOGGED false
SOURCE_PATH_LOGGED false
RAW_FILE_CONTENT_LOGGED false
```

## 5.4 Ledger progress

```powershell
node `
  .\scripts\queryCandidatePlannerShowRealShadowRegistryProgress.js `
  --ledger $ledger `
  --source-catalog $sourceCatalog

if ($LASTEXITCODE -ne 0) {
  throw "Registry progress failed"
}
```

Required:

```text
PROGRESS 10/10
REMAINING 0
COMPLETE true
LEGACY_LEDGER_ACCEPTED false
```

---

# 6. PHASE 15.3-A Finalization Gate

10/10 이후 한 번만 실행한다.

```powershell
node `
  .\scripts\queryCandidatePlannerFinalizeRealShadowEvidenceFoundation.js `
  --ledger $ledger `
  --source-catalog $sourceCatalog `
  --expected-rejection-attestation $attestation `
  --registry-output ".\queryCandidatePlannerRealShadowCaseRegistry.private.json" `
  --railway-output ".\queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt" `
  --registry-summary-output ".\queryCandidatePlannerRealShadowCaseRegistry.summary.private.json" `
  --foundation-summary-output ".\queryCandidatePlannerRealShadowEvidenceFoundation.summary.private.json"

if ($LASTEXITCODE -ne 0) {
  throw "PHASE 15.3-A finalization failed"
}
```

Required output:

```text
PASS phase 15.3-A real shadow evidence foundation finalized
EXPECTED_REJECTION_CASES 1
EXPECTED_REJECTION_EVIDENCE 1
READY_FOR_PATCH_15_3_2_C true
COLLECTOR_ENABLED_BY_THIS_OPERATION false
INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false
PRODUCTION_PROMOTION_AUTHORIZED false
PRIVATE_OUTPUT_DO_NOT_COMMIT true
```

# 7. Private output guard

```powershell
node `
  .\scripts\queryCandidatePlannerAssertRealShadowFoundationPrivateOutputsUntracked.js

if ($LASTEXITCODE -ne 0) {
  throw "Phase 15.3-A private output staging violation"
}
```

Expected:

```text
PASS no phase 15.3-A private outputs staged
```

# 8. Phase exit criteria

PHASE 15.3-A 완료 선언은 아래가 모두 참일 때만 가능하다.

```text
Uploadable Source Catalog              10/10 COMPLETE
Source-bound v2 Fingerprint Ledger     10/10 COMPLETE
Expected Rejection Evidence            1/1 VERIFIED
Real Shadow Case Registry              FINALIZED
Private Evidence Git isolation         PASS
Legacy Ledger accepted                 false
Evidence Collector enabled             false
Internal Canary enabled                false
Production Promotion authorized        false
READY_FOR_PATCH_15_3_2_C               true
```

그 다음 단계는 **Patch 15.3.2-C — Encryption Secret & Safe Deployment**다.
