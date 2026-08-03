# 패치 12.1 — Live Shadow + Resolver Re-entry Bridge

## 목적

패치 12의 조건부 Planner가 실제로 `CALL_REQUIRED`가 되는 전용 fixture를 추가하고, 수용된 LLM proposal을 기존 결정론적 계층으로 재진입시킵니다.

```text
Conditional Planner
→ proposal validation
→ ACCEPTED_FOR_REVALIDATION
→ Resolver
→ Candidate Family
→ Feasibility Gate
→ Deterministic Ranker
```

전체 체인은 `shadow-only`입니다. 기존 사용자 후보군, production route, 원본 후보 상태를 변경하지 않습니다.

## 신규 계약

```text
query_candidate_planner_shadow_resolution_v1
query_candidate_planner_reentry_bundle_v1
conditional_llm_candidate_planner_shadow_policy_v1
```

## 전용 fixture

```text
tests/fixtures/query-candidate-planner-shadow/
  call_required_group_avg_time_count/
```

fixture 조건:

```text
분석 가능한 물리 테이블 1개
24행·4열
READY 후보 0개
미해결 후보 존재
GROUP_AGGREGATION·TIME_SERIES·OVERVIEW 기회 존재
```

예상 trigger:

```text
required   = true
reasonCode = NO_READY_CANDIDATE
```

mock proposal:

```text
부서별 만족도 평균  → group_avg
기준일별 응답 건수  → time_count
```

예상 shadow 결과:

```text
provider calls  1
accepted        2
RESOLVED        2
READY           2
RANKED          2
status          SHADOW_COMPLETED
```

## Resolver Re-entry 방식

Planner proposal의 실제 `tableId`·`columnId`를 사용해 shadow 전용 retrieval/capability 문서를 생성합니다.

- `bindingStatus`: `INFERRED`
- `executorSupport`: `GENERIC`
- 출력: `summarySheet`
- source: proposal이 참조한 단일 physical table
- recipeId: operation과 실제 열 이름으로 결정론적 구성
- 원본 후보와 병합하지 않음

`time_count`는 다음 계약으로 Resolver에 재진입합니다.

```text
operation: time_count
essential operand: period 1개
measure: 불필요
```

Planner re-entry 후보에서는 해시 기반 candidateId보다 실제 operand를 포함한 recipeId를 먼저 해석합니다. 일반 후보의 기존 해석 순서는 유지됩니다.

## Shadow 경계

```text
shadowOnly                  true
productionCandidateMerge    false
productionReadyAssignment   false
productionRouteChanged      false
sourceCandidateStatusMutation false
plaintextPersistenceAllowed false
```

LLM proposal이 shadow 체인에서 `READY`가 되더라도 production 후보로 승격되지 않습니다.

# 적용

```powershell
Expand-Archive `
  .\query_candidate_patch12_1_live_shadow_resolver_reentry_bridge.zip `
  -DestinationPath . `
  -Force
```

## 파일 확인

```powershell
Test-Path .\automation\queryCandidatePlannerShadowBridge.js
Test-Path .\automation\queryCandidatePlannerShadowBridge.schema.json
Test-Path .\tests\queryCandidatePlannerShadowCapture.js
Test-Path .\tests\queryCandidatePlannerLiveShadowSmokeTest.js
Test-Path .\tests\fixtures\query-candidate-planner-shadow\call_required_group_avg_time_count\semantic-profile.json
Test-Path .\PATCH_VALIDATION_PATCH12_1.json
Test-Path .\PATCH_MANIFEST_PATCH12_1.json
```

모두 `True`여야 합니다.

## 문법 검사

```powershell
node --check .\automation\queryCandidateResolver.js
node --check .\automation\queryCandidatePlannerShadowBridge.js
node --check .\tests\queryCandidatePlannerShadowCapture.js
node --check .\tests\queryCandidatePlannerLiveShadowSmokeTest.js
```

## 신규 기능 테스트

```powershell
node .\tests\queryCandidatePlannerShadowFixtureSmokeTest.js
node .\tests\queryCandidatePlannerResolverReentryBridgeSmokeTest.js
node .\tests\queryCandidatePlannerShadowNoProductionMutationSmokeTest.js
node .\tests\queryCandidatePlannerShadowFailureSafeSmokeTest.js
node .\tests\queryCandidatePlannerShadowSchemaSmokeTest.js

node .\tests\queryCandidatePatch12_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate planner shadow fixture smoke
PASS query candidate planner resolver re-entry bridge smoke
PASS query candidate planner shadow no production mutation smoke
PASS query candidate planner shadow failure safe smoke
PASS query candidate planner shadow schema smoke
PASS query candidate patch12.1 source integrity smoke
PASS query candidate patch12.1 manifest smoke
```

# 전용 fixture 기준선

```powershell
node .\tests\queryCandidatePlannerShadowCapture.js `
  --mode=write

node .\tests\queryCandidatePlannerShadowCapture.js `
  --mode=compare

node .\tests\queryCandidatePlannerShadowBaselineSmokeTest.js
node .\tests\queryCandidatePlannerShadowSampleAudit.js
```

정상 결과:

```text
[query-candidate-planner-shadow] cases: 1
[query-candidate-planner-shadow] PASS 1/1

required   true
invocation CALLED
calls      1
accepted   2
resolved   2
ready      2
ranked     2
status     SHADOW_COMPLETED
```

생성 파일:

```text
tests/fixtures/query-candidate-planner-shadow/
  candidate-planner-shadow-resolution-index.json
  candidate-planner-shadow-sample-audit.json
  candidate-planner-shadow-sample-audit.md

  call_required_group_avg_time_count/
    candidate-planner-shadow-resolution.json
```

# 실제 OpenAI Live Shadow 실행

기본 테스트에서는 실제 API를 호출하지 않습니다. 명시적인 환경변수를 설정한 경우에만 1회 실행합니다.

```powershell
$env:QUERY_CANDIDATE_PLANNER_LIVE_SHADOW = "1"
$env:OPENAI_API_KEY = "실제_API_키"

# 배포 환경에서 사용 가능한 Responses API 모델로 명시
$env:QUERY_CANDIDATE_PLANNER_MODEL = "현재_사용_모델"
$env:QUERY_CANDIDATE_PLANNER_REASONING_EFFORT = "low"

node .\tests\queryCandidatePlannerLiveShadowSmokeTest.js
```

proposal까지 반드시 확인할 때:

```powershell
$env:QUERY_CANDIDATE_PLANNER_LIVE_SHADOW_REQUIRE_PROPOSAL = "1"
node .\tests\queryCandidatePlannerLiveShadowSmokeTest.js
```

실행 후 환경변수 제거:

```powershell
Remove-Item Env:QUERY_CANDIDATE_PLANNER_LIVE_SHADOW -ErrorAction SilentlyContinue
Remove-Item Env:QUERY_CANDIDATE_PLANNER_LIVE_SHADOW_REQUIRE_PROPOSAL -ErrorAction SilentlyContinue
Remove-Item Env:QUERY_CANDIDATE_PLANNER_MODEL -ErrorAction SilentlyContinue
Remove-Item Env:QUERY_CANDIDATE_PLANNER_REASONING_EFFORT -ErrorAction SilentlyContinue
Remove-Item Env:OPENAI_API_KEY -ErrorAction SilentlyContinue
```

기본 출력:

```text
tests/fixtures/query-candidate-planner-shadow/
  call_required_group_avg_time_count/
    candidate-planner-live-shadow-resolution.json
```

실제 모델이 `NO_ADDITION`을 반환하면 provider 호출은 성공했지만 re-entry는 실행하지 않으며 `SKIPPED`로 기록됩니다. `QUERY_CANDIDATE_PLANNER_LIVE_SHADOW_REQUIRE_PROPOSAL=1`에서는 proposal이 없으면 테스트가 실패합니다.

# 패치 12 및 기존 계층 회귀

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

사용자 저장소의 실제 기준선에서는 모두 다음 결과여야 합니다.

```text
PASS 6/6
differences=0
```

## 현재 범위에서 하지 않는 것

```text
production route 자동 연결
기존 candidate-resolution에 proposal 병합
shadow READY의 production READY 승격
사용자 UI 후보군 변경
실제 생성 executor 실행
평문 캐시 저장
원본 행·샘플값·파일명 전송
```

다음 운영 단계는 실제 Live Shadow 산출물을 누적해 proposal 수용률, Resolver 통과율, Feasibility READY 비율을 확인한 뒤 production 승격 정책을 별도 패치로 설계하는 것입니다.
