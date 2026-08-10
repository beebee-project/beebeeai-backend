# Patch 15.3.2-E — Real Shadow Observation Collection

## 목적

Patch 15.3.2-E는 Patch 15.3.2-D에서 제한 활성화한 암호화 Evidence Collector를 이용해 실제 내부 Allowlist 트래픽의 Observation을 평가 가능한 고정 세트로 수집한다.

이 단계는 평가 자체를 수행하지 않는다. 다음 Patch 15.3.2-F/G가 사용할 실제 Evidence를 확보하고 수집 구간을 고정하는 단계다.

## 선행 Gate

다음 Runtime Gate가 이미 PASS해야 한다.

```text
PASS patch 15.3.2-D limited collector runtime verification
REGISTRY_CASES 10
ALLOWLIST_ENTRIES 1
COLLECTOR_ENABLED true
COLLECTOR_KILL_SWITCH false
READY_FOR_PATCH_15_3_2_E true
INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false
PRODUCTION_PROMOTION_AUTHORIZED false
```

## 수집 중 유지해야 하는 Railway 상태

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=1
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_TTL_DAYS=7
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_MAX_RECORDS=5000

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

Internal Canary와 Production Promotion은 이 패치에서 활성화하지 않는다.

## 수집 완료 계약

기존 Evidence Bundle Builder의 최소 계약은 실제 Execution 30건 이상, 10개 case 각각 3건 이상이다.

Patch E는 이후 Operational Evaluation까지 한 번에 진행할 수 있도록 더 강한 실제 QA 프로토콜을 요구한다.

10개 case 각각:

```text
1. 최초 업로드 후 후보 조회        COLD
2. 같은 업로드 후보 재조회         WARM 1
3. 같은 업로드 후보 재조회         WARM 2
4. 다운로드 1회
5. 같은 업로드 후보 재조회         DOWNLOAD_REUSE
6. 파일 삭제
7. 동일 파일 재업로드
8. 후보 조회                       REUPLOAD
```

case당 요구 Evidence:

```text
Execution                  >= 5
DOWNLOAD lifecycle         >= 1
DELETE lifecycle           >= 1
Distinct upload identity   >= 2
```

10개 case 전체 목표:

```text
Execution                  >= 50
Lifecycle                  >= 20
Registry case coverage     10/10
Privacy violation          0
Guardrail violation        0
```

## 중요: 수집 Window 시작 전 정리

기존 테스트 업로드가 남아 있다면 수집 Window를 시작하기 전에 삭제한다. Window 시작 후 사전 정리를 수행하면 첫 실행의 Lifecycle 순서가 Evaluation에 섞일 수 있다.

## Patch 설치

ZIP은 repository root에 바로 풀리는 flat 구조다.

```powershell
Expand-Archive `
  .\query_candidate_patch15_3_2_E_real_shadow_observation_collection.zip `
  -DestinationPath . `
  -Force
```

## 신규 QA

```powershell
$tests = @(
    Get-ChildItem ".\tests\queryCandidatePatch15_3_2_E*SmokeTest.js" |
    Sort-Object Name
)

if ($tests.Count -ne 20) {
    throw "Expected exactly 20 Patch E smoke tests"
}

foreach ($test in $tests) {
    Write-Host "RUN $($test.Name)"
    & node $test.FullName
    if ($LASTEXITCODE -ne 0) {
        throw "Patch E QA failed: $($test.Name)"
    }
}

Write-Host "PASS Patch 15.3.2-E QA 20/20"
```

## 수집 Window 시작

Railway 환경변수와 D Runtime Gate를 사용하므로 `railway run`으로 실행한다.

```powershell
railway run node .\scripts\queryCandidatePlannerStartRealShadowObservationCollection.js --output .\queryCandidatePlannerRealShadowObservationCollectionWindow.private.json
```

정상:

```text
PASS patch 15.3.2-E observation collection window started
COLLECTOR_ENABLED true
COLLECTOR_KILL_SWITCH false
INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false
PRODUCTION_PROMOTION_AUTHORIZED false
PRIVATE_OUTPUT_DO_NOT_COMMIT true
```

Window 파일은 다시 만들지 않는다. PowerShell 세션이 종료되어도 같은 파일을 계속 사용한다.

## 실제 10개 Source

```text
hardcase_two_tables_one_sheet_waste
  .local_uploads\생활폐기물_한시트_표2개.xlsx

real_world_event_applicant_workshop
  .local_uploads\real_world_event_applicant_workshop.csv

seed_attendance_conditional
  .local_uploads\attendance_check_report.csv

seed_sales_ready
  .local_uploads\real_world_sales_multi_branch_monthly.csv

template_course_evaluation_report
  .local_uploads\course_evaluation_report.csv

seed_unstructured_unsupported
  .local_uploads\서식중심_양식형_쿼리화_테스트.xlsx

ambiguous_mixed_columns_review
  .local_uploads\교차표_기간컬럼_쿼리화_테스트.xlsx

inventory_stock_movement
  .local_uploads\inventory_stock_status.csv

project_task_tracker
  .local_uploads\task_issue_tracking_report.csv

expense_claim_review
  .local_uploads\travel_expense_settlement_analysis.csv
```

`seed_unstructured_unsupported`에서 후보 없음 결과는 정상 Expected Rejection이다. 해당 Case도 5회의 실제 Shadow 실행 Observation을 수집한다.

## Pilot Case 권장

10개 전체를 반복하기 전에 첫 Case 하나만 8단계 프로토콜로 실행하고 Progress를 확인한다.

```powershell
railway run node .\scripts\queryCandidatePlannerShowRealShadowObservationCollectionProgress.js --window .\queryCandidatePlannerRealShadowObservationCollectionWindow.private.json
```

첫 Case의 정상 목표:

```text
READY hardcase_two_tables_one_sheet_waste EXEC=5/5 DOWNLOAD=1/1 DELETE=1/1 IDENTITIES=2/2
EXECUTIONS 5/50
LIFECYCLE 2/20
READY_FOR_PATCH_15_3_2_F false
```

Pilot Case가 READY가 되기 전에는 나머지 9개를 진행하지 않는다.

다운로드 후 `DOWNLOAD=0`이면 현재 UI의 다른 다운로드 경로를 반복하지 말고 원인을 먼저 확인한다. 가능한 경우 업로드 원본 다운로드 경로를 우선 사용한다.

## 전체 Progress 확인

```powershell
railway run node .\scripts\queryCandidatePlannerShowRealShadowObservationCollectionProgress.js --window .\queryCandidatePlannerRealShadowObservationCollectionWindow.private.json
```

최종 목표:

```text
READY <10 case 모두>
EXECUTIONS 50/50
LIFECYCLE 20/20
BUILDER_MINIMUM_READY true
COLLECTION_PROTOCOL_COMPLETE true
PRIVACY_VIOLATIONS 0
GUARDRAIL_VIOLATIONS 0
READY_FOR_PATCH_15_3_2_F true
RAW_RECORDS_LOGGED false
FINGERPRINTS_LOGGED false
```

Record 원문과 Fingerprint는 Progress stdout에 출력하지 않는다.

## Collection Finalize

Progress가 READY일 때만 실행한다.

```powershell
railway run node .\scripts\queryCandidatePlannerFinalizeRealShadowObservationCollection.js --window .\queryCandidatePlannerRealShadowObservationCollectionWindow.private.json --output .\queryCandidatePlannerRealShadowObservationCollection.summary.private.json
```

정상:

```text
PASS patch 15.3.2-E real shadow observation collection finalized
EXECUTIONS >=50
LIFECYCLE >=20
READY_FOR_PATCH_15_3_2_F true
COLLECTOR_ENABLED_BY_THIS_OPERATION false
INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false
PRODUCTION_PROMOTION_AUTHORIZED false
RAW_RECORDS_LOGGED false
FINGERPRINTS_LOGGED false
PRIVATE_OUTPUT_DO_NOT_COMMIT true
```

Summary의 `to` 시각이 Patch 15.3.2-F Export의 상한선이 된다.

## Finalize 직후 Collector Freeze

Collection Finalize 후 추가 Observation이 평가 Dataset에 섞이지 않도록 Railway에서 Collector를 다시 fail-closed로 고정한다.

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=1
```

Secret, Registry JSON/SHA, Allowlist, TTL, MAX_RECORDS는 삭제하지 않는다. Patch F Export에 필요하다.

## Private Output Guard

```powershell
& node ".\scripts\queryCandidatePlannerAssertRealShadowObservationCollectionPrivateOutputsUntracked.js"

if ($LASTEXITCODE -ne 0) {
    throw "Patch E private-output guard failed"
}
```

정상:

```text
PASS no observation-collection private outputs staged
```

## Patch E 완료 조건

```text
Patch D Runtime                     PASS
Patch E QA                          20/20 PASS
Actual Execution                    >=50
Actual Lifecycle                    >=20
10-case protocol                    COMPLETE
Privacy violations                  0
Guardrail violations                0
READY_FOR_PATCH_15_3_2_F            true
Private outputs                     untracked
Collector after finalization        OFF / kill-switch ON
Internal Canary                     OFF
Production Promotion                BLOCKED
```

이후 PHASE 15.3-C의 Patch 15.3.2-F — Observation Export & Actual Pricing Policy로 이동한다.
