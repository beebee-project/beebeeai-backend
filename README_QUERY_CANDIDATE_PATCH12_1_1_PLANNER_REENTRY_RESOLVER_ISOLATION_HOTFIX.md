# Patch 12.1.1 — Planner Re-entry Resolver Isolation Hotfix

## 목적

Patch 12.1은 Live Shadow에서 수용된 Planner proposal을 Resolver → Family →
Feasibility → Ranker로 재진입시키기 위해 `count_rows`와 `time_count` 계약을
Resolver에 추가했습니다.

초기 구현에서는 해당 계약 일부가 일반 후보에도 적용되어 기존
`real_world_event_applicant_workshop`의 candidate-resolution SHA가 변경됐습니다.
후보 수와 상태 통계는 같았지만 기존 Resolver 기준선은 다음과 같이 달라졌습니다.

```text
expected ae97854a2e8a5fd16bcd85189f23190f2447cf8ece5414f3175a6da18ea8e257
actual   f0649967b741290e6fa874c146d38ed1a554eb86c9551be63174828a84279383
```

이 Hotfix는 Planner 전용 계약을 일반 Resolver 경로에서 완전히 격리합니다.

## 변경 내용

### 일반 Resolver 경로

다음 동작을 Patch 12 이전 상태로 복원합니다.

```text
일반 time_count 후보
- 전역 structural generic 목록에 time_count를 추가하지 않음
- 전역 recipe operand spec에서 time_count를 해석하지 않음
- 기존 candidateId 우선 식별자 순서 유지
```

`count_rows`는 기존 전역 structural recipe 지위를 유지하지만, 새로 추가된
operand spec은 일반 후보에 적용하지 않습니다.

### Planner Re-entry 경로

다음 조건에서만 전용 계약을 활성화합니다.

```javascript
retrievalItem.provenance?.plannerReentry === true
```

활성화되는 전용 계약:

```text
count_rows
- operand 0개

time_count
- period operand 1개
- measure 불필요
- Planner recipeId 우선 해석
- structural generic 후보로 처리
```

### 불변 경계

```text
Shadow 전용
Production candidate merge 없음
Production READY assignment 없음
Production route 변경 없음
기존 source candidate 상태 변경 없음
평문 캐시 없음
```

## 적용

저장소 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch12_1_1_planner_reentry_resolver_isolation_hotfix.zip `
  -DestinationPath . `
  -Force
```

## 파일 확인

```powershell
Test-Path .\automation\queryCandidateResolver.js
Test-Path .\tests\queryCandidatePlannerResolverIsolationSmokeTest.js
Test-Path .\tests\queryCandidatePlannerResolverLegacyTimeCountCompatibilitySmokeTest.js
Test-Path .\tests\queryCandidatePatch12_1_1SourceIntegritySmokeTest.js
Test-Path .\tests\queryCandidatePatch12_1_1ManifestSmokeTest.js
Test-Path .\PATCH_VALIDATION_PATCH12_1_1.json
Test-Path .\PATCH_MANIFEST_PATCH12_1_1.json
```

모두 `True`여야 합니다.

## 문법 검사

```powershell
node --check .\automation\queryCandidateResolver.js
node --check .\tests\queryCandidatePlannerResolverIsolationSmokeTest.js
node --check .\tests\queryCandidatePlannerResolverLegacyTimeCountCompatibilitySmokeTest.js
```

## Hotfix 신규 테스트

```powershell
node .\tests\queryCandidatePlannerResolverIsolationSmokeTest.js
node .\tests\queryCandidatePlannerResolverLegacyTimeCountCompatibilitySmokeTest.js
node .\tests\queryCandidatePatch12_1_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_1ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate planner resolver isolation smoke
PASS query candidate planner resolver legacy time_count compatibility smoke
PASS query candidate patch12.1.1 source integrity smoke
PASS query candidate patch12.1.1 manifest smoke
```

Legacy compatibility 테스트는 일반 `time_count` 합성 후보의 의미 계약을
필드별로 검증하고, 해당 후보의 item SHA와 fixture 외부 메타데이터에 영향을
받지 않는 semantic compatibility fingerprint를 고정합니다.

```text
resolutionItemSha256
7b3616e2e2f6c377bcf76582fdc96f7a18200e3d86f7d7f7536292358bde81bc

semanticCompatibilityFingerprintSha256
fcd34372b8d1316284e51e816bd0c95ad7c1625f7cd2be7ae0605c3fb8c1503a
```

전체 `resolutionSha256`는 동일 seed fixture의 다른 후보와 source hash에도
의존하므로 exact 값으로 고정하지 않습니다. 대신 64자리 SHA-256 형식만
확인합니다. 이 변경은 Patch 12.1.1.1에서 적용됐으며 production 코드는
변경하지 않습니다.

## 기존 Shadow 재진입 검증

```powershell
node .\tests\queryCandidatePlannerShadowFixtureSmokeTest.js
node .\tests\queryCandidatePlannerResolverReentryBridgeSmokeTest.js
node .\tests\queryCandidatePlannerShadowNoProductionMutationSmokeTest.js
node .\tests\queryCandidatePlannerShadowFailureSafeSmokeTest.js
node .\tests\queryCandidatePlannerShadowSchemaSmokeTest.js

node .\tests\queryCandidatePatch12_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1ManifestSmokeTest.js
```

전용 fixture 기대값:

```text
required   true
invocation CALLED
calls      1
accepted   2
resolved   2
ready      2
ranked     2
status     SHADOW_COMPLETED
```

Fixture 기준선도 다시 비교합니다.

```powershell
node .\tests\queryCandidatePlannerShadowCapture.js `
  --mode=compare

node .\tests\queryCandidatePlannerShadowBaselineSmokeTest.js
node .\tests\queryCandidatePlannerShadowSampleAudit.js
```

기대 결과:

```text
[query-candidate-planner-shadow] PASS 1/1
PASS query candidate planner shadow baseline smoke 1
PASS candidate planner shadow sample audit
```

## 기존 6개 Resolver 기준선 복원 확인

기준선을 다시 쓰지 않고 비교만 실행합니다.

```powershell
node .\tests\queryCandidateResolverCapture.js `
  --mode=compare
```

정상 결과:

```text
[query-candidate-resolver] cases: 6
[query-candidate-resolver] PASS 6/6
differences=0
```

특히 행사 신청자 케이스가 다음 expected SHA와 다시 일치해야 합니다.

```text
real_world_event_applicant_workshop
candidate-resolution SHA-256
= ae97854a2e8a5fd16bcd85189f23190f2447cf8ece5414f3175a6da18ea8e257
```

다음 명령은 실행하지 않습니다.

```powershell
node .\tests\queryCandidateResolverCapture.js `
  --mode=write
```

## 전체 하위 계층 회귀

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

모두 사용자 저장소에서 `PASS 6/6`, `differences=0`이어야 합니다.

## Live Shadow 재실행

이 Hotfix는 API 크레딧 실패를 수정하지 않습니다. OpenAI API 잔액을 보충한 뒤
실제 proposal 필수 모드로 다시 실행합니다.

```powershell
$env:QUERY_CANDIDATE_PLANNER_LIVE_SHADOW = "1"
$env:QUERY_CANDIDATE_PLANNER_LIVE_SHADOW_REQUIRE_PROPOSAL = "1"
$env:OPENAI_API_KEY = "실제_OpenAI_API_키"
$env:QUERY_CANDIDATE_PLANNER_MODEL = "gpt-5.6-terra"
$env:QUERY_CANDIDATE_PLANNER_REASONING_EFFORT = "low"

node .\tests\queryCandidatePlannerLiveShadowSmokeTest.js
```

이전 `credit_balance_exhausted` 상태가 해소되지 않으면 Live 결과는 계속
`FAILED_SAFE`가 됩니다. 해당 경우에도 기존 production 후보는 변경되지 않습니다.
