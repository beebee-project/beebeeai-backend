# Patch 14.1 — Service Boundary Integration

## API Shadow Wiring + Shadow Comparator + 기존 응답 완전 유지

## 1. 적용 순서

```text
Patch 13.3 — Live Provider Cache-Hit Parity & Production Readiness Gate
→ Patch 14.0 — Feature Flags + Kill Switch
→ Patch 14.1 — Service Boundary Integration
```

Patch 14.0 신규 QA 7개가 모두 PASS인 상태에서 적용합니다.

## 2. 목적

실제 후보군 API 진입점인 다음 경로에 Shadow 관찰 경계를 연결합니다.

```text
POST /api/automation/analysis-candidates
→ automationController.getAnalysisCandidates
→ 기존 Primary JSON 응답 전송
→ API Shadow task 실행
→ Shadow Comparator
→ 내부 observation/log만 기록
```

Primary 응답을 먼저 전송하므로 Shadow Provider 지연, 오류, timeout,
Comparator 오류는 HTTP status·header·payload에 영향을 주지 않습니다.

## 3. 핵심 계약

```text
Boundary version   query_candidate_planner_api_shadow_boundary_v1
Service version    query_candidate_planner_api_shadow_service_v1
Runner version     query_candidate_planner_api_shadow_runner_v1
Comparator version query_candidate_planner_api_shadow_comparator_v1
Observation        query_candidate_planner_api_shadow_observation_v1
```

고정 Guardrail:

```text
primaryResponseAuthority      true
responsePayloadMutation       false
responseHeaderMutation        false
responseStatusMutation        false
productionCandidateMerge      false
productionReadyAssignment     false
productionRouteChanged        false
sourceCandidateStatusMutation false
```

## 4. 기본 동작

Patch 14.0 기본값 때문에 신규 Shadow는 기본 차단됩니다.

```text
QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED 기본 0
→ FEATURE_DISABLED
→ Shadow runner 호출 0회
→ 기존 API 응답만 반환
```

기본 OFF 상태에서는 OpenAI 비용이 발생하지 않습니다.

## 5. Shadow 실행 경계

Feature Control이 `SHADOW_EXECUTION`을 허용한 요청만
`queryCandidatePlannerShadowBridge.runCandidatePlannerLiveShadow()`로 전달합니다.

Provider·Cache는 각각 Patch 14.0 결정 결과를 별도로 전달합니다.

```text
Provider flag/kill switch 차단
→ blocked provider 주입
→ 실제 Provider 호출 차단

Cache read/write 차단
→ cacheReadAllowed/cacheWriteAllowed false 전달
```

Patch 14.2 전까지 업로드 생명주기 기반 tenant·upload cache 객체는 API에서
주입하지 않습니다. 이번 패치는 API 실행 경계와 비교 관찰까지만 연결합니다.

## 6. 개인정보 경계

API 응답 전체를 Shadow Provider에 전달하지 않습니다. 다음 값은 제거합니다.

```text
rows
rawRows
sampleValues
원본 파일명
queryTablesKey
tenantId
사용자 email
원본 Primary response
```

Shadow 입력에는 최대 20개 테이블·각 120개 열의 schema metadata와
후보 ID·유형·operation·table reference만 포함합니다.

## 7. Shadow Comparator

비교 대상:

```text
Primary: candidateUiPayload.recommendedCandidates 우선
         없으면 topCandidates 및 후보 배열
Shadow:  items / rankingResolution.items / rankedCandidates 등
```

비교 결과:

```text
MATCH
PARTIAL_MATCH
MISMATCH
NO_SHADOW_CANDIDATES
```

기록 지표:

```text
primary/shadow/shared count
Top-1 동일 여부
Top-3 overlap
Jaccard
rank agreement
순서 SHA-256
추가·누락 candidate identity SHA-256
```

원문 candidate ID는 observation 로그에 기록하지 않습니다.

## 8. 변경 파일

```text
automation/queryCandidatePlannerFeatureControlRuntime.js
automation/queryCandidatePlannerShadowComparator.js
automation/queryCandidatePlannerApiShadowRunner.js
automation/queryCandidatePlannerApiShadowService.js
automation/queryCandidatePlannerApiShadowBoundary.js
automation/queryCandidatePlannerApiShadowObservation.schema.json
routes/automationRoutes.js

tests/queryCandidatePatch14_1TestSupport.js
tests/queryCandidatePatch14_1DefaultBlockedSmokeTest.js
tests/queryCandidatePatch14_1ApiShadowWiringSmokeTest.js
tests/queryCandidatePatch14_1DefaultBridgeAdapterSmokeTest.js
tests/queryCandidatePatch14_1FailureIsolationSmokeTest.js
tests/queryCandidatePatch14_1TimeoutIsolationSmokeTest.js
tests/queryCandidatePatch14_1ShadowComparatorSmokeTest.js
tests/queryCandidatePatch14_1PrivacyBoundarySmokeTest.js
tests/queryCandidatePatch14_1KillSwitchBoundarySmokeTest.js
tests/queryCandidatePatch14_1RouteContractSmokeTest.js
tests/queryCandidatePatch14_1SourceIntegritySmokeTest.js
tests/queryCandidatePatch14_1ManifestSmokeTest.js

PATCH_VALIDATION_PATCH14_1.json
PATCH_MANIFEST_PATCH14_1.json
README_QUERY_CANDIDATE_PATCH14_1_SERVICE_BOUNDARY_API_SHADOW.md
```

Controller 본문, UI, Production merge, READY assignment, executor는 변경하지 않습니다.

## 9. 적용

백엔드 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch14_1_service_boundary_api_shadow.zip `
  -DestinationPath . `
  -Force
```

## 10. 문법 검사

```powershell
node --check .\automation\queryCandidatePlannerFeatureControlRuntime.js
node --check .\automation\queryCandidatePlannerShadowComparator.js
node --check .\automation\queryCandidatePlannerApiShadowRunner.js
node --check .\automation\queryCandidatePlannerApiShadowService.js
node --check .\automation\queryCandidatePlannerApiShadowBoundary.js
node --check .\routes\automationRoutes.js

node --check .\tests\queryCandidatePatch14_1TestSupport.js
node --check .\tests\queryCandidatePatch14_1DefaultBlockedSmokeTest.js
node --check .\tests\queryCandidatePatch14_1ApiShadowWiringSmokeTest.js
node --check .\tests\queryCandidatePatch14_1DefaultBridgeAdapterSmokeTest.js
node --check .\tests\queryCandidatePatch14_1FailureIsolationSmokeTest.js
node --check .\tests\queryCandidatePatch14_1TimeoutIsolationSmokeTest.js
node --check .\tests\queryCandidatePatch14_1ShadowComparatorSmokeTest.js
node --check .\tests\queryCandidatePatch14_1PrivacyBoundarySmokeTest.js
node --check .\tests\queryCandidatePatch14_1KillSwitchBoundarySmokeTest.js
node --check .\tests\queryCandidatePatch14_1RouteContractSmokeTest.js
node --check .\tests\queryCandidatePatch14_1SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
```

## 11. Patch 14.1 비용 없는 QA

```powershell
node .\tests\queryCandidatePatch14_1DefaultBlockedSmokeTest.js
node .\tests\queryCandidatePatch14_1ApiShadowWiringSmokeTest.js
node .\tests\queryCandidatePatch14_1DefaultBridgeAdapterSmokeTest.js
node .\tests\queryCandidatePatch14_1FailureIsolationSmokeTest.js
node .\tests\queryCandidatePatch14_1TimeoutIsolationSmokeTest.js
node .\tests\queryCandidatePatch14_1ShadowComparatorSmokeTest.js
node .\tests\queryCandidatePatch14_1PrivacyBoundarySmokeTest.js
node .\tests\queryCandidatePatch14_1KillSwitchBoundarySmokeTest.js
node .\tests\queryCandidatePatch14_1RouteContractSmokeTest.js
node .\tests\queryCandidatePatch14_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
```

정상 출력:

```text
PASS query candidate patch14.1 default blocked smoke
PASS query candidate patch14.1 API shadow wiring smoke
PASS query candidate patch14.1 default bridge adapter smoke
PASS query candidate patch14.1 failure isolation smoke
PASS query candidate patch14.1 timeout isolation smoke
PASS query candidate patch14.1 shadow comparator smoke
PASS query candidate patch14.1 privacy boundary smoke
PASS query candidate patch14.1 kill-switch boundary smoke
PASS query candidate patch14.1 route contract smoke
PASS query candidate patch14.1 source integrity smoke
PASS query candidate patch14.1 manifest smoke
```

위 검사는 mock Shadow runner만 사용하며 실제 OpenAI Provider를 호출하지 않습니다.

## 12. Patch 14.0 누적 QA

```powershell
node .\tests\queryCandidatePatch14_0DefaultFailClosedSmokeTest.js
node .\tests\queryCandidatePatch14_0FlagMatrixSmokeTest.js
node .\tests\queryCandidatePatch14_0KillSwitchPrecedenceSmokeTest.js
node .\tests\queryCandidatePatch14_0InvalidEnvironmentSmokeTest.js
node .\tests\queryCandidatePatch14_0ReadinessEvidenceSmokeTest.js
node .\tests\queryCandidatePatch14_0SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_0ManifestSmokeTest.js
```

## 13. 기존 Shadow·Planner 회귀

기준선은 다시 작성하지 않고 비교만 실행합니다.

```powershell
node .\tests\queryCandidatePlannerShadowCapture.js `
  --mode=compare

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

정상 조건:

```text
Shadow       PASS 1/1 differences=0
Planner      PASS 6/6 differences=0
Ranker       PASS 6/6 differences=0
Feasibility  PASS 6/6 differences=0
Family       PASS 6/6 differences=0
Resolver     PASS 6/6 differences=0
```

`--mode=write`는 실행하지 않습니다.

## 14. 내부 Shadow 활성화 순서

QA 완료 전 Railway 환경변수는 변경하지 않습니다.

비용 없는 경계 확인:

```powershell
$env:QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED = "1"
$env:QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED = "1"
$env:QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED = "0"
```

실제 Provider Shadow는 별도 내부 검증에서만 다음을 추가합니다.

```powershell
$env:QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED = "1"
```

기존 `OPENAI_API_KEY`, 접근 가능한 모델 ID, reasoning effort가 필요합니다.
일반 사용자 응답은 두 경우 모두 기존 후보군 그대로입니다.

## 15. 다음 단계

```text
Patch 14.2 — Upload Lifecycle + Cache Invalidation Wiring
```

14.2에서 API Shadow 실행에 tenant·upload fingerprint·queryJson SHA와
암호화 L3/L4 cache instance를 연결하고 업로드 삭제 시 전체 무효화를 검증합니다.
