# Patch 15.3.2-B — Real Shadow 10-Case Registry Finalization

Patch 15.3.2-B는 Patch 15.0 Accuracy Dataset의 10개 case를 실제 내부 요청에서 관찰한 두 fingerprint와 연결해 최종 Real Shadow Case Registry를 만든다.

이 패치는 실제 fingerprint를 임의로 만들지 않는다. 실제 내부 계정으로 각 기준 파일을 업로드하고 후보군 조회까지 수행한 뒤, Internal Preview 또는 API Shadow Observation에서 확인한 값을 기록해야 한다.

## 안전 경계

패치 적용만으로 다음 상태는 바뀌지 않는다.

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
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

결과적으로:

```text
Evidence Collector                 OFF
Internal Canary                    OFF
Production Merge                   OFF
Production Route 변경              없음
사용자 응답                         PRIMARY 유지
```

## 포함 파일

```text
automation/queryCandidatePlannerRealShadowRegistryFinalization.js
scripts/queryCandidatePlannerScaffoldRealShadowFingerprintLedger.js
scripts/queryCandidatePlannerRecordRealShadowFingerprint.js
scripts/queryCandidatePlannerShowRealShadowRegistryProgress.js
scripts/queryCandidatePlannerFinalizeRealShadowCaseRegistry.js
scripts/queryCandidatePlannerAssertRealShadowPrivateOutputsUntracked.js
evaluation/queryCandidatePlannerRealShadowFingerprintLedger.template.json
```

QA와 정합성 검사를 위한 테스트 파일도 함께 포함한다.

## Registry 대상 10개 case

Patch 15.0 Accuracy Dataset의 caseId와 정확히 일치해야 한다.

```text
1. hardcase_two_tables_one_sheet_waste
2. real_world_event_applicant_workshop
3. seed_attendance_conditional
4. seed_sales_ready
5. template_course_evaluation_report
6. seed_unstructured_unsupported
7. ambiguous_mixed_columns_review
8. inventory_stock_movement
9. project_task_tracker
10. expense_claim_review
```

추가 case는 허용하지 않으며 하나라도 누락되면 최종 Registry 생성이 차단된다.

## 저장 가능한 정보

```text
caseId
scenarioId
requestFingerprintSha256
uploadFingerprintSha256
captureSource
capturedAt
expectedColdCostMicrousd
modelId
짧은 operatorNote
```

## 저장 금지 정보

다음 필드가 Ledger에 들어오면 Fail-closed 처리한다.

```text
이메일·이름
MongoDB user _id
Google ID
Tenant·Organization ID
파일명·원본 파일명
queryTablesKey·storageKey
원본 행·샘플 값
JWT·Bearer Token
암호화 Secret
```

## fingerprint 출처

각 case를 내부 Allowlist 계정으로 실제 실행한 뒤 아래 값을 확인한다.

```text
requestFingerprintSha256
cacheLifecycle.identity.uploadFingerprintSha256
```

Lifecycle Observation에서는 upload fingerprint가 다음 경로로 보일 수 있다.

```text
identity.uploadFingerprintSha256
```

허용되는 `captureSource` 값은 다음 두 개뿐이다.

```text
API_SHADOW_OBSERVATION
INTERNAL_PREVIEW
```

Patch 15.2 합성 Observation Dataset의 fingerprint나 임의 SHA-256은 사용할 수 없다.

---

# 1. 적용

```powershell
$ErrorActionPreference = "Stop"

Get-FileHash `
  .\query_candidate_patch15_3_2_B_real_shadow_registry_finalization.zip `
  -Algorithm SHA256
```

제공된 ZIP SHA-256과 일치하면 적용한다.

```powershell
Expand-Archive `
  .\query_candidate_patch15_3_2_B_real_shadow_registry_finalization.zip `
  -DestinationPath . `
  -Force
```

# 2. 설치 확인

```powershell
Test-Path `
  .\automation\queryCandidatePlannerRealShadowRegistryFinalization.js

Test-Path `
  .\scripts\queryCandidatePlannerScaffoldRealShadowFingerprintLedger.js

Test-Path `
  .\scripts\queryCandidatePlannerRecordRealShadowFingerprint.js

Test-Path `
  .\scripts\queryCandidatePlannerFinalizeRealShadowCaseRegistry.js

Test-Path `
  .\evaluation\queryCandidatePlannerRealShadowFingerprintLedger.template.json
```

모두 `True`여야 한다.

# 3. 문법 검사

```powershell
node --check `
  .\automation\queryCandidatePlannerRealShadowRegistryFinalization.js

Get-ChildItem `
  .\scripts\queryCandidatePlanner*RealShadow*.js |
  ForEach-Object {
    node --check $_.FullName

    if ($LASTEXITCODE -ne 0) {
      throw "Syntax check failed: $($_.Name)"
    }
  }
```

# 4. Patch 15.3.2-B QA

```powershell
Get-ChildItem `
  .\tests\queryCandidatePatch15_3_2_B*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    Write-Host "RUN $($_.Name)"

    node $_.FullName

    if ($LASTEXITCODE -ne 0) {
      throw "Patch 15.3.2-B test failed: $($_.Name)"
    }
  }

Write-Host "PASS Patch 15.3.2-B QA 12/12"
```

검사 범위:

```text
10-case Scaffold
실제 fingerprint 기록
미등록 case 차단
잘못된 fingerprint 차단
중복 fingerprint 차단
개인정보 필드 차단
불완전 Registry 차단
10/10 Finalization
결과 결정성
Private Output 스테이징 차단
Source Integrity
Manifest
```

# 5. 선행 정합성 검사

```powershell
$tests = @(
  ".\tests\queryCandidatePatch15_3_2_ASourceIntegritySmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2_AManifestSmokeTest.js",
  ".\tests\queryCandidatePatch15_3SourceIntegritySmokeTest.js",
  ".\tests\queryCandidatePatch15_3ManifestSmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2SourceIntegritySmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2ManifestSmokeTest.js",
  ".\tests\queryCandidatePatch15_3_1PredecessorIntegrityRepairSmokeTest.js"
)

foreach ($test in $tests) {
  Write-Host "RUN $test"
  node $test

  if ($LASTEXITCODE -ne 0) {
    throw "Predecessor integrity failed: $test"
  }
}

Write-Host "PASS Patch 15.3.2-B and predecessor integrity"
```

---

# 6. Private Ledger 생성

Patch 코드의 QA와 커밋을 먼저 끝낸 다음 실제 fingerprint 작업을 시작한다.

```powershell
node `
  .\scripts\queryCandidatePlannerScaffoldRealShadowFingerprintLedger.js `
  --registry-id "internal_real_shadow_2026_08_v1" `
  --output ".\queryCandidatePlannerRealShadowFingerprintLedger.private.json"
```

예상 출력:

```text
PASS fingerprint ledger scaffold cases=10
PRIVATE_OUTPUT_DO_NOT_COMMIT true
```

Private 파일이 `git status`에 나타나지 않도록 로컬 저장소 전용 exclude에 추가할 수 있다.

```powershell
@"
queryCandidatePlannerRealShadowFingerprintLedger.private.json
queryCandidatePlannerRealShadowCaseRegistry.private.json
queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt
queryCandidatePlannerRealShadowCaseRegistry.summary.private.json
queryCandidatePlannerRealShadowEvidenceSecret.private.txt
"@ | Add-Content .\.git\info\exclude
```

`.git/info/exclude`는 로컬 전용이며 원격 저장소에 커밋되지 않는다.

# 7. 실제 10개 요청 수행

각 case에 대응하는 기준 엑셀을 내부 Allowlist 계정으로 다음 순서로 실행한다.

```text
업로드
→ 후보군 조회
→ Internal Preview 또는 API Shadow Observation 확인
→ request fingerprint 기록
→ upload fingerprint 기록
```

첫 Registry에는 최초 업로드 identity의 upload fingerprint를 사용한다. 삭제 후 재업로드 fingerprint는 Patch 15.3.2-E 수집 단계에서 별도 관찰한다.

## 기록 명령

실제 값만 로컬 PowerShell에 입력한다. 원본 사용자 ID나 파일명은 입력하지 않는다.

```powershell
node `
  .\scripts\queryCandidatePlannerRecordRealShadowFingerprint.js `
  --ledger ".\queryCandidatePlannerRealShadowFingerprintLedger.private.json" `
  --case-id "hardcase_two_tables_one_sheet_waste" `
  --request-fingerprint "실제_64자리_request_SHA256" `
  --upload-fingerprint "실제_64자리_upload_SHA256" `
  --capture-source "INTERNAL_PREVIEW" `
  --captured-at "$([DateTime]::UtcNow.ToString('o'))"
```

API Shadow Observation에서 확인한 경우:

```powershell
--capture-source "API_SHADOW_OBSERVATION"
```

성공 시 fingerprint 원문은 다시 출력하지 않는다.

```text
PASS recorded case=<caseId>
PROGRESS 1/10
REMAINING 9
RAW_FINGERPRINTS_LOGGED false
```

같은 명령에서 `--case-id`와 실제 두 fingerprint만 바꿔 10개 case를 모두 기록한다.

# 8. 진행률 확인

```powershell
node `
  .\scripts\queryCandidatePlannerShowRealShadowRegistryProgress.js `
  --ledger ".\queryCandidatePlannerRealShadowFingerprintLedger.private.json"
```

완료 전 예시:

```text
READY hardcase_two_tables_one_sheet_waste
PENDING real_world_event_applicant_workshop
...
PROGRESS 1/10
REMAINING 9
COMPLETE false
```

완료 기준:

```text
PROGRESS 10/10
REMAINING 0
COMPLETE true
```

# 9. 최종 Registry 생성

10개 모두 완료된 뒤 실행한다.

```powershell
node `
  .\scripts\queryCandidatePlannerFinalizeRealShadowCaseRegistry.js `
  --ledger ".\queryCandidatePlannerRealShadowFingerprintLedger.private.json" `
  --output ".\queryCandidatePlannerRealShadowCaseRegistry.private.json" `
  --railway-output ".\queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt" `
  --summary-output ".\queryCandidatePlannerRealShadowCaseRegistry.summary.private.json"
```

정상 결과:

```text
PASS real shadow registry finalized sha256=<64자리> cases=10
PRIVATE_OUTPUT_DO_NOT_COMMIT true
COLLECTOR_ENABLED_BY_THIS_OPERATION false
INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false
PRODUCTION_PROMOTION_AUTHORIZED false
```

생성 파일:

```text
queryCandidatePlannerRealShadowCaseRegistry.private.json
queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt
queryCandidatePlannerRealShadowCaseRegistry.summary.private.json
```

Railway용 파일에는 다음 한 줄이 들어 있다.

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON=<한 줄 JSON>
```

Patch 15.3.2-C 전에는 Railway에 등록하지 않는다.

# 10. Private 출력 스테이징 차단 검사

```powershell
node `
  .\scripts\queryCandidatePlannerAssertRealShadowPrivateOutputsUntracked.js

if ($LASTEXITCODE -ne 0) {
  throw "Private Real Shadow output is staged"
}
```

정상 결과:

```text
PASS no real shadow private outputs staged
```

# 완료 기준

```text
Patch 15.3.2-B QA                    12/12 PASS
Patch 15.3.2-A 정합성                PASS
Patch 15.3·15.3.2 정합성             PASS
기준 case                            10/10
Request fingerprint                  10개·중복 없음
Upload fingerprint                   10개·중복 없음
Capture source                       모두 실제 관찰 출처
actualTraffic                        모두 true
synthetic                            모두 false
Private Registry 생성                PASS
Registry runtime contract            PASS
Collector                            계속 OFF
Internal Canary                      계속 OFF
Production Promotion                 승인 안 함
```

Patch 15.3.2-B의 실제 운영 완료는 코드 설치만으로 성립하지 않는다. 실제 내부 요청 10건에서 fingerprint를 수집해 `PROGRESS 10/10`과 Registry Finalization PASS까지 확인해야 한다.
