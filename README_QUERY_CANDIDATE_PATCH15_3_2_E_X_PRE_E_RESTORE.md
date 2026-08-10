# Patch 15.3.2-E-X — Real Shadow Observation Collection Exclusion & Pre-E Restore

## 결정

Patch 15.3.2-E — Real Shadow Observation Collection을 활성 15.3 로드맵에서 완전히 배제한다.

정확한 상태 표현:

```text
15.3.2-D Limited Evidence Collector Activation  CLOSED
15.3.2-E Real Shadow Observation Collection     EXCLUDED / N/A
15.3.2-F Evaluation Baseline & Operational Assessment  NEXT
15.3.2-G Evaluation Evidence Bundle & Internal Canary Readiness
15.3.3 Internal Production Exposure
15.3.4 Production Promotion Gate + 15.4 Readiness
15.4 Controlled Production Rollout
```

`15.3.2-E`는 `DEFERRED`, `BLOCKED`, `SKIPPED`가 아니다. 후속 패치의 prerequisite도 아니다.

## 이 교정 패치가 하는 일

- Patch E에서 추가한 Observation Collection automation/scripts/tests/docs를 제거한다.
- repository root의 Patch E collection window/summary private 파일을 제거한다.
- Patch E 문제 진단을 위해 임시 추가했던 `[real-shadow-evidence]`, `[real-shadow-subject]` 로그를 제거하고 알려진 wrapper를 원래 fire-and-forget 호출로 복원한다.
- Patch 15.3.2-D 구현은 유지한다.
- 새로운 로드맵 override를 추가한다.
- Production, Internal Canary, Promotion을 활성화하지 않는다.

## 중요한 구조 변경

15.3.2-F/G는 더 이상 Patch E 실제 Observation을 요구하지 않는다.
실제 Production traffic 검증 책임은 15.3.3 Internal Production Exposure와 15.3.4 Production Promotion Gate로 이동한다.

단, Patch E에서 발견된 `COLLECTOR_SUBJECT_NOT_ALLOWLISTED`는 제품 기능 실패로 승격하지 않지만, 15.3.3이 같은 internal subject contract를 사용할 수 있으므로 **15.3.3 직전 independent allowlist readiness gate**에서 별도로 검증한다.

## 적용

ZIP은 flat 구조다. repository root에 바로 풀고 실행한다.

```powershell
Expand-Archive `
  .\query_candidate_patch15_3_2_E_X_pre_e_restore_and_exclusion.zip `
  -DestinationPath . `
  -Force

& .\apply_patch15_3_2_E_X_pre_e_restore.ps1
```

정상:

```text
PASS Patch 15.3.2-E-X pre-E source restore applied
PATCH_15_3_2_E_STATUS EXCLUDED_NA
NEXT_PATCH 15.3.2-F
RUNTIME_ACTION_REQUIRED EVIDENCE_ENABLED=0 EVIDENCE_KILL_SWITCH=1
```

## QA

```powershell
$tests = @(
  Get-ChildItem ".\tests\queryCandidatePatch15_3_2_E_X*SmokeTest.js" |
  Sort-Object Name
)

if ($tests.Count -ne 7) {
  throw "Expected exactly 7 Patch 15.3.2-E-X smoke tests"
}

foreach ($test in $tests) {
  Write-Host "RUN $($test.Name)"
  & node $test.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Patch 15.3.2-E-X QA failed: $($test.Name)"
  }
}

& node .\scripts\queryCandidatePlannerVerifyPatch15_3_2_E_XPreERestore.js
if ($LASTEXITCODE -ne 0) {
  throw "Patch 15.3.2-E-X final verifier failed"
}

Write-Host "PASS Patch 15.3.2-E-X QA 7/7 + verifier"
```

## Railway runtime restore

Source restore 후 Railway에서:

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_KILL_SWITCH=1
```

으로 변경 후 Deploy한다.

다음은 유지한다.

```text
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED=0
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH=1
QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE=BLOCKED
QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH=1
```

Evidence Secret/Registry/Allowlist 변수는 삭제할 필요가 없다. Collector만 fail-closed로 닫는다.

Railway Deploy 완료 후 새 runtime gate를 실행한다.

```powershell
railway run node .\scripts\queryCandidatePlannerVerifyPatch15_3_2_E_XRuntime.js
```

정상:

```text
PASS patch 15.3.2-E-X runtime restore verification
PATCH_15_3_2_E_STATUS EXCLUDED_NA
EVIDENCE_COLLECTOR_ENABLED false
EVIDENCE_KILL_SWITCH true
INTERNAL_CANARY_ENABLED false
PRODUCTION_ENABLED false
PROMOTION_GATE_ENABLED false
PROMOTION_AUDIENCE_MODE BLOCKED
PROMOTION_ROLLOUT_PERCENT 0
READY_FOR_PATCH_15_3_2_F true
PRODUCTION_PROMOTION_AUTHORIZED false
```

## 완료 상태

```text
15.3.2-D  CLOSED
15.3.2-E  EXCLUDED / N/A
15.3.2-E-X CLOSED
NEXT      15.3.2-F Evaluation Baseline & Operational Assessment
```
