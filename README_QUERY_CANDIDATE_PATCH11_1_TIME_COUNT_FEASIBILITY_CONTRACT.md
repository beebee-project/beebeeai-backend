# Patch 11.1 — Time Count Feasibility Contract

## 목적

Patch 11 표본 감사에서 `sheet1t1_timecount_신청일자`가 다음 사유로 `UNSUPPORTED` 처리됐습니다.

```text
GENERIC_OPERATION_NOT_SUPPORTED
OPERAND_BINDING_NOT_PASSED
```

그러나 `time_count`는 기간 열별 행 건수를 계산하는 기존 summarySheet operation입니다. Patch 11.1은 Feasibility Gate에 `time_count/timecount` 실행 계약을 추가합니다.

## 정책 버전

```text
deterministic_candidate_feasibility_policy_v1_1
```

출력 문서·item·execution plan 계약 버전은 그대로 유지합니다.

```text
query_candidate_feasibility_resolution_v1
query_candidate_feasibility_item_v1
query_candidate_execution_plan_v1
```

## Time Count 계약

```text
operation: time_count 또는 timecount
필수 operand: period 1개
measure operand: 불필요
source: 단일하게 확정된 source scope
output: summarysheet
```

period columnId 확인 순서:

1. 기존 `operandBinding.operands`의 period 결속
2. `requiredRoles.period`의 PASS columnId
3. source 범위에서 단 하나로 확정되는 matched period/date column

두 번째·세 번째 경로는 기존 Resolver가 `time_count` operand를 명시적으로 만들지 않은 경우에만 Feasibility 전용 fallback으로 사용합니다.

## 보수적 경계

```text
period columnId 1개  -> READY 가능
period columnId 없음 -> UNSUPPORTED
period columnId 2개 이상 -> UNSUPPORTED
measure 누락 -> 차단하지 않음
```

모호한 기간 열을 임의 선택하지 않습니다.

## 변경 파일

```text
automation/queryCandidateFeasibilityGate.js
automation/queryCandidateFeasibilityGate.schema.json

README_QUERY_CANDIDATE_PATCH11_DETERMINISTIC_FEASIBILITY_GATE.md
PATCH_VALIDATION_PATCH11.json
PATCH_MANIFEST_PATCH11.json

tests/queryCandidateFeasibilityGateTimeCountSmokeTest.js
tests/queryCandidateFeasibilityGateTimeCountAliasSmokeTest.js
tests/queryCandidateFeasibilityGateTimeCountMissingPeriodSmokeTest.js
tests/queryCandidateFeasibilityGateTimeCountAmbiguousPeriodSmokeTest.js

tests/queryCandidatePatch11SourceIntegritySmokeTest.js

tests/queryCandidatePatch11_1SourceIntegritySmokeTest.js
tests/queryCandidatePatch11_1ManifestSmokeTest.js
PATCH_VALIDATION_PATCH11_1.json
PATCH_MANIFEST_PATCH11_1.json
```

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch11_1_time_count_feasibility_contract.zip `
  -DestinationPath . `
  -Force
```

## 신규 검증

```powershell
node --check .\automation\queryCandidateFeasibilityGate.js

node .\tests\queryCandidateFeasibilityGateTimeCountSmokeTest.js
node .\tests\queryCandidateFeasibilityGateTimeCountAliasSmokeTest.js
node .\tests\queryCandidateFeasibilityGateTimeCountMissingPeriodSmokeTest.js
node .\tests\queryCandidateFeasibilityGateTimeCountAmbiguousPeriodSmokeTest.js

node .\tests\queryCandidatePatch11_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch11_1ManifestSmokeTest.js
```

## Patch 11 누적 검증

```powershell
node .\tests\queryCandidateFeasibilityGateSmokeTest.js
node .\tests\queryCandidateFeasibilityGateDeclaredExecutorSmokeTest.js
node .\tests\queryCandidateFeasibilityGateGenericOperandSmokeTest.js
node .\tests\queryCandidateFeasibilityGateStructuralReviewSmokeTest.js
node .\tests\queryCandidateFeasibilityGateUnsupportedSmokeTest.js
node .\tests\queryCandidateFeasibilityGateMissingOperandSmokeTest.js
node .\tests\queryCandidateFeasibilityGateSelectedOnlySmokeTest.js
node .\tests\queryCandidateFeasibilityGateRankIndependenceSmokeTest.js
node .\tests\queryCandidateFeasibilityGatePrivacyBoundarySmokeTest.js
node .\tests\queryCandidateFeasibilityGateSchemaSmokeTest.js
node .\tests\queryCandidateFeasibilityGateBaselineSmokeTest.js

node .\tests\queryCandidatePatch11SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch11ManifestSmokeTest.js
```

## 기준선 재작성

정책 버전과 행사 신청자 Feasibility 결과가 변경되므로 기존 기준선을 다시 작성합니다.

```powershell
node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=write

node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=compare
```

예상 핵심 변화:

```text
real_world_event_applicant_workshop
READY       13 -> 14
UNSUPPORTED  1 -> 0
```

전체 예상:

```text
READY       51 -> 52
REVIEW       4 -> 4
UNSUPPORTED  1 -> 0
```

정책 버전 확인:

```powershell
(Get-Content `
  .\tests\fixtures\query-candidate-baseline\candidate-feasibility-resolution-index.json `
  -Raw | ConvertFrom-Json).policyVersion
```

기대값:

```text
deterministic_candidate_feasibility_policy_v1_1
```

## 재감사

```powershell
node .\tests\queryCandidateFeasibilityGateSampleAudit.js `
  --limit=20
```

확인할 후보:

```text
sheet1t1_timecount_신청일자
status    = READY
operation = timecount
source    = Sheet1#T1
columns   = Sheet1#T1.column_3
blocking  = 없음
```

## 비변경 사항

```text
Production route 변경 없음
OpenAI 호출 없음
Ranker 의존 없음
원본 후보 상태 변경 없음
Patch 9 family disposition 변경 없음
원본 행·sample value 저장 없음
```
