# Patch 12.1.1.1 — Legacy Time Count Compatibility Test Hash Fix

## 목적

Patch 12.1.1의 Resolver 격리 자체는 실제 사용자 저장소에서 정상 확인됐습니다.

```text
queryCandidateResolverCapture --mode=compare
PASS 6/6
real_world_event_applicant_workshop differences=0
resolutionSha256=ae97854a2e8a5fd16bcd85189f23190f2447cf8ece5414f3175a6da18ea8e257
```

하지만 합성 `time_count` 호환성 테스트는 전체 `candidate-resolution` SHA를
exact 값으로 고정하고 있었습니다. 전체 resolution SHA는 테스트 대상 후보뿐
아니라 같은 seed fixture의 다른 후보, retrieval/capability source hash와
메타데이터에도 의존합니다. 따라서 Resolver 의미 동작이 동일해도 저장소의
누적 fixture 상태에 따라 전체 SHA가 달라질 수 있습니다.

이 패치는 production 코드를 변경하지 않고 테스트 경계만 바로잡습니다.

## 변경 내용

### 제거한 검증

```text
전체 resolutionSha256 exact equality
16688690c40ee7b7838c2e41cd63740564d91a018ba5de5faa9b85992561f946
```

### 유지한 검증

일반 `time_count` 후보 전체 item SHA는 계속 고정합니다.

```text
resolutionItemSha256
7b3616e2e2f6c377bcf76582fdc96f7a18200e3d86f7d7f7536292358bde81bc
```

### 추가한 의미 검증

다음 필드를 명시적으로 확인합니다.

```text
plannerReentry=false
candidateId=sales_table_timecount_거래일자
recipeId=time_count
result=RESOLVED
previousRetrievalResult=DEFERRED
bindingStatus=INFERRED
bindingSource=IDENTIFIER_INFERENCE
sourceTableIds=[sales_table]
matchedColumnIds=[sales_table.column_1]
period role=PASS
operandBinding=NOT_APPLICABLE
operandBinding.reasonCode=NO_EXPLICIT_RECIPE_OPERANDS
required capabilities 모두 PASS
terminalPriorResult=false
semanticReassessmentPerformed=true
```

### Semantic Compatibility Fingerprint

Fixture 주변 source hash와 다른 후보 상태를 제외하고 위 의미 계약만 직렬화해
별도의 SHA를 고정합니다.

```text
semanticCompatibilityFingerprintSha256
fcd34372b8d1316284e51e816bd0c95ad7c1625f7cd2be7ae0605c3fb8c1503a
```

전체 `resolutionSha256`는 exact 값 대신 64자리 소문자 SHA-256 형식만
확인합니다.

## 불변 경계

```text
production code 변경 없음
Resolver 정책 변경 없음
Planner Re-entry 계약 변경 없음
Shadow fixture 변경 없음
기존 baseline 재작성 없음
production route 변경 없음
평문 persistence 변경 없음
```

## 적용

저장소 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch12_1_1_1_legacy_time_count_compatibility_test_hash_fix.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check `
  .\tests\queryCandidatePlannerResolverLegacyTimeCountCompatibilitySmokeTest.js

node --check `
  .\tests\queryCandidatePatch12_1_1_1SourceIntegritySmokeTest.js

node --check `
  .\tests\queryCandidatePatch12_1_1_1ManifestSmokeTest.js
```

## 패치 테스트

```powershell
node `
  .\tests\queryCandidatePlannerResolverLegacyTimeCountCompatibilitySmokeTest.js

node `
  .\tests\queryCandidatePatch12_1_1SourceIntegritySmokeTest.js

node `
  .\tests\queryCandidatePatch12_1_1ManifestSmokeTest.js

node `
  .\tests\queryCandidatePatch12_1_1_1SourceIntegritySmokeTest.js

node `
  .\tests\queryCandidatePatch12_1_1_1ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate planner resolver legacy time_count compatibility smoke
PASS query candidate patch12.1.1 source integrity smoke
PASS query candidate patch12.1.1 manifest smoke
PASS query candidate patch12.1.1.1 source integrity smoke
PASS query candidate patch12.1.1.1 manifest smoke
```

## Shadow 회귀

```powershell
node .\tests\queryCandidatePlannerShadowCapture.js `
  --mode=compare

node .\tests\queryCandidatePlannerShadowBaselineSmokeTest.js
node .\tests\queryCandidatePlannerShadowSampleAudit.js
```

기대값:

```text
PASS 1/1
accepted=2
resolved=2
ready=2
ranked=2
status=SHADOW_COMPLETED
productionCandidateMerge=false
productionRouteChanged=false
```

## 전체 기존 계층 회귀

기준선을 다시 쓰지 않고 비교만 실행합니다.

```powershell
node .\tests\queryCandidatePlannerCapture.js `
  --mode=compare

node .\tests\queryCandidateRankerCapture.js `
  --mode=compare

node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=compare

node .\tests\queryCandidateFamilyResolverCapture.js `
  --mode=compare

node .\tests\queryCandidateResolverCapture.js `
  --mode=compare
```

모두 `PASS 6/6`, `differences=0`이어야 합니다.

다음 기준선 작성 명령은 실행하지 않습니다.

```powershell
node .\tests\queryCandidateResolverCapture.js `
  --mode=write
```

## Live Shadow

이 패치는 API 잔액 문제를 변경하지 않습니다. 이전 실제 호출의
`credit_balance_exhausted`가 해결된 뒤 Live Shadow proposal 필수 모드를
별도로 재실행합니다.
