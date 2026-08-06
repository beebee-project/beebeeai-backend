# Patch 15.3.2-A — Real Shadow Case Registry & Encryption Secret Preparation

## 목적

Patch 15.3.2-A는 Patch 15.3.2의 실제 Shadow Evidence Collector를 켜기 전에 필요한 두 운영 자산을 안전하게 준비한다.

1. Patch 15.0 Accuracy Dataset 10개 case와 실제 내부 요청 fingerprint를 연결한 Case Registry
2. Real Shadow Evidence 전용 48-byte 무작위 암호화 Secret

이 패치는 Route, Collector, Production Merge, Internal Canary를 변경하지 않는다.

```text
Evidence Collector                 계속 OFF
Internal Canary                    계속 OFF
Production Merge                   계속 OFF
Production Route                   변경 없음
일반·내부 사용자 응답              기존 Primary 유지
```

## 추가 파일

```text
automation/queryCandidatePlannerRealShadowPreparation.js
scripts/queryCandidatePlannerGenerateRealShadowEvidenceSecret.js
scripts/queryCandidatePlannerScaffoldRealShadowCaseRegistry.js
scripts/queryCandidatePlannerPrepareRealShadowCaseRegistry.js
scripts/queryCandidatePlannerVerifyRealShadowPreparation.js
evaluation/queryCandidatePlannerRealShadowCaseRegistry.draft.json
tests/queryCandidatePatch15_3_2_A*.js
PATCH_MANIFEST_PATCH15_3_2_A.json
PATCH_VALIDATION_PATCH15_3_2_A.json
```

기존 Production 소스는 교체하지 않는다.

## Patch 15.0 기준 10개 case

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

Registry Builder는 위 10개가 모두 존재하고, 추가 case가 없으며, fingerprint가 중복되지 않아야 PASS한다.

## 적용

```powershell
$ErrorActionPreference = "Stop"

Get-FileHash `
  .\query_candidate_patch15_3_2_A_real_shadow_preparation.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_3_2_A_real_shadow_preparation.zip `
  -DestinationPath . `
  -Force
```

## 설치 확인

```powershell
Test-Path `
  .\automation\queryCandidatePlannerRealShadowPreparation.js

Test-Path `
  .\scripts\queryCandidatePlannerGenerateRealShadowEvidenceSecret.js

Test-Path `
  .\scripts\queryCandidatePlannerPrepareRealShadowCaseRegistry.js

Test-Path `
  .\evaluation\queryCandidatePlannerRealShadowCaseRegistry.draft.json
```

모두 `True`여야 한다.

## 문법 검사

```powershell
node --check `
  .\automation\queryCandidatePlannerRealShadowPreparation.js

Get-ChildItem `
  .\scripts\queryCandidatePlanner*RealShadow*.js |
  ForEach-Object {
    node --check $_.FullName

    if ($LASTEXITCODE -ne 0) {
      throw "Syntax check failed: $($_.Name)"
    }
  }
```

## Patch 15.3.2-A QA

`tests` 폴더가 `.gitignore` 대상이어도 로컬 실행은 가능하다.

```powershell
Get-ChildItem `
  .\tests\queryCandidatePatch15_3_2_A*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    Write-Host "RUN $($_.Name)"
    node $_.FullName

    if ($LASTEXITCODE -ne 0) {
      throw "Patch 15.3.2-A test failed: $($_.Name)"
    }
  }
```

예상 결과는 9개 PASS다.

## 1단계 — Collector OFF 상태 유지

Fingerprint 발견 중에는 아래 값을 유지한다.

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

## 2단계 — 실제 fingerprint 발견

Patch 15.0 기준 파일을 내부 Allowlist 계정으로 각각 한 번 업로드하고 후보군 조회까지 실행한다.

Internal Preview 또는 API Shadow Observation에서 아래 두 값을 기록한다.

```text
requestFingerprintSha256
cacheLifecycle.identity.uploadFingerprintSha256
```

Lifecycle Observation 화면에서는 다음 경로로 보일 수 있다.

```text
identity.uploadFingerprintSha256
```

각 값은 소문자 또는 대문자 64자리 SHA-256이어야 한다.

주의:

- Patch 15.2 합성 Observation Dataset의 fingerprint를 사용하지 않는다.
- 파일명, 이메일, MongoDB `_id`, queryTablesKey를 Registry에 넣지 않는다.
- 동일 fingerprint를 둘 이상의 case에 등록하지 않는다.
- 최초 업로드의 upload fingerprint를 등록한다. 재업로드 identity는 실제 수집 단계에서 별도로 관찰된다.

## 3단계 — Registry Draft 작성

기본 Draft:

```text
evaluation/queryCandidatePlannerRealShadowCaseRegistry.draft.json
```

직접 편집하기 전에 작업용 사본을 만든다.

```powershell
Copy-Item `
  .\evaluation\queryCandidatePlannerRealShadowCaseRegistry.draft.json `
  .\queryCandidatePlannerRealShadowCaseRegistry.private.draft.json
```

각 case의 두 항목을 실제 값으로 채운다.

```json
{
  "requestFingerprintSha256": "실제 64자리 request fingerprint",
  "uploadFingerprintSha256": "실제 64자리 upload fingerprint"
}
```

`expectedColdCostMicrousd`는 Patch 15.3.2-A에서는 0으로 유지할 수 있다. 실제 승인 가격은 Evidence Bundle 작성 단계에서 별도로 적용한다.

원본 Accuracy Dataset에서 Draft를 다시 만들려면:

```powershell
node `
  .\scripts\queryCandidatePlannerScaffoldRealShadowCaseRegistry.js `
  --registry-id "internal_real_shadow_2026_08_v1" `
  --output ".\queryCandidatePlannerRealShadowCaseRegistry.private.draft.json"
```

## 4단계 — Registry 생성

```powershell
node `
  .\scripts\queryCandidatePlannerPrepareRealShadowCaseRegistry.js `
  --draft ".\queryCandidatePlannerRealShadowCaseRegistry.private.draft.json" `
  --output ".\queryCandidatePlannerRealShadowCaseRegistry.private.json" `
  --railway-output ".\queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt" `
  --summary-output ".\queryCandidatePlannerRealShadowCaseRegistry.summary.private.json"
```

PASS 시:

```text
PASS real shadow registry sha256=<64자리> cases=10
```

생성 파일:

```text
queryCandidatePlannerRealShadowCaseRegistry.private.json
queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt
queryCandidatePlannerRealShadowCaseRegistry.summary.private.json
```

이 세 파일은 내부 운영 자료이며 공개 저장소에 커밋하지 않는다.

## 5단계 — 전용 암호화 Secret 생성

Registry가 PASS한 후 생성한다.

```powershell
node `
  .\scripts\queryCandidatePlannerGenerateRealShadowEvidenceSecret.js
```

출력 형식:

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET=<64자리 Base64URL>
SECRET_SHA256=<64자리 SHA-256>
```

첫 번째 줄의 값을 Railway Secret 변수에 복사한다.

금지:

```text
JWT_SECRET 재사용
FILE_ENCRYPTION_SECRET 재사용
QUERY_JSON_SECRET 재사용
Git 커밋
채팅·이메일 공유
수집 기간 중 임의 변경
```

Secret 원문은 이후 확인 로그에 출력하지 않고 `SECRET_SHA256`만 비교한다.

## 6단계 — 로컬 준비 검증

현재 PowerShell 세션에만 Secret과 기존 Allowlist SHA-256을 넣는다.

```powershell
$env:QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET = `
  "방금 생성한 64자리 Secret"

$env:QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256 = `
  "기존 내부 계정 64자리 subject SHA-256"
```

Production 안전값도 명시한다.

```powershell
$env:QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH = "1"
$env:QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE = "BLOCKED"
$env:QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH = "1"
```

검증:

```powershell
node `
  .\scripts\queryCandidatePlannerVerifyRealShadowPreparation.js `
  --registry ".\queryCandidatePlannerRealShadowCaseRegistry.private.json"
```

정상 출력:

```text
PASS real shadow preparation registrySha256=<64자리>
SECRET_SHA256 <64자리>
ALLOWLIST_ENTRIES 1
CASE_COUNT 10
COLLECTOR_ENABLED_BY_THIS_OPERATION false
INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false
PRODUCTION_PROMOTION_AUTHORIZED false
```

이 검사는 Collector를 켜지 않고, 실제 활성화 시 Configuration이 유효할지만 메모리에서 검증한다.

## 7단계 — Railway 변수 등록

검증 통과 후 아래 변수만 등록한다.

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET=<생성한 Secret>
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_TTL_DAYS=7
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_MAX_RECORDS=5000
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON=<railway.private.txt의 = 뒤 한 줄 JSON>
```

등록 후에도 `EVIDENCE_ENABLED=0`을 유지한다. Patch 15.3.2-B에서 배포 로그와 Runtime Configuration을 확인한 뒤 Collector만 제한적으로 ON한다.

## Git 적용

신규 Production 파일은 없고 도구·Draft·테스트만 추가된다.

```powershell
git add `
  .\automation\queryCandidatePlannerRealShadowPreparation.js `
  .\scripts\queryCandidatePlannerGenerateRealShadowEvidenceSecret.js `
  .\scripts\queryCandidatePlannerScaffoldRealShadowCaseRegistry.js `
  .\scripts\queryCandidatePlannerPrepareRealShadowCaseRegistry.js `
  .\scripts\queryCandidatePlannerVerifyRealShadowPreparation.js `
  .\evaluation\queryCandidatePlannerRealShadowCaseRegistry.draft.json `
  .\PATCH_MANIFEST_PATCH15_3_2_A.json `
  .\PATCH_VALIDATION_PATCH15_3_2_A.json `
  .\README_QUERY_CANDIDATE_PATCH15_3_2_A_REAL_SHADOW_PREPARATION.md
```

`tests`가 `.gitignore` 대상이면 강제 추가한다.

```powershell
git add -f `
  .\tests\queryCandidatePatch15_3_2_A*SmokeTest.js
```

다음 private 파일은 추가하지 않는다.

```text
*.private.json
*.private.txt
Secret 원문
```

## 완료 기준

```text
Patch 15.3.2-A QA                         9/9 PASS
Accuracy Dataset case coverage            10/10
Request fingerprint                       10개·중복 없음
Upload fingerprint                        10개·중복 없음
Registry runtime contract                 PASS
Secret                                    48 random bytes / Base64URL 64자
Collector                                 OFF
Internal Canary                           OFF
Production Promotion                      금지
```
