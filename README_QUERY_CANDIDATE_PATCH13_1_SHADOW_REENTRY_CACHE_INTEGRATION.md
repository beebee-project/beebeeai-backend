# Patch 13.1 — Shadow Re-entry Cache Integration

## 1. 적용 순서

이 패치는 다음 순서로 적용합니다.

```text
Patch 12.1.2.2
→ Patch 13 — Encrypted Hierarchical Cache Foundation
→ Patch 13.1 — Shadow Re-entry Cache Integration
```

Patch 13의 신규 검사와 기존 Shadow·Planner·Resolver 기준선이 모두 PASS인 상태에서 적용합니다.

## 2. 목적

Patch 13에서 추가한 암호화 계층형 캐시를 실제 Shadow 경로에 **명시적 opt-in 방식**으로 연결합니다.

```text
첫 실행
Provider CALLED
→ L3 PLANNER_PROVIDER_RESULT 암호화 저장
→ 검증된 Planner resolution L3 암호화 저장
→ Resolver → Family → Feasibility → Ranker 실행
→ 검증된 Shadow re-entry artifact L4 암호화 저장

동일 입력 재실행
Provider CACHE_HIT
→ L3 Planner resolution 검증
→ L4 Shadow re-entry artifact 검증
→ Provider 추가 호출 없음
```

production route는 변경하지 않습니다. `runCandidatePlannerLiveShadow()`에 `hierarchicalCache`, `tenantId`, `cacheSecret`을 전달한 Shadow 실행에서만 활성화됩니다. 기존 인자만 사용하는 Shadow 경로는 이전 결과와 동일해야 합니다.

## 3. 핵심 계약

```text
integration version
query_candidate_planner_shadow_cache_integration_v1

re-entry artifact version
query_candidate_planner_shadow_reentry_cache_artifact_v1

policy version
encrypted_shadow_reentry_cache_policy_v1
```

캐시 계층:

```text
L3_SEMANTIC / PLANNER_PROVIDER_RESULT
기존 Planner provider 결과 캐시

L3_SEMANTIC / PLANNER_RESOLUTION
strict contract를 통과한 Planner resolution 감사 캐시

L4_REENTRY / SHADOW_REENTRY
Resolver·Family·Feasibility·Ranker 검증을 통과한 re-entry artifact 캐시
```

## 4. 결정론적 동일성

Planner invocation의 `CALLED`와 `CACHE_HIT` 상태는 서로 다르므로 전체 Planner resolution SHA는 달라질 수 있습니다. Re-entry 체인의 입력 계약은 invocation 상태가 아니라 다음 proposal set으로 고정합니다.

```text
planner input SHA
semantic profile SHA
accepted candidate ID
planner item SHA
```

이 proposal set SHA를 canonical re-entry source로 사용하므로 다음 값은 첫 실행, 동일 프로세스 재실행, 새 cache instance 재실행에서 같아야 합니다.

```text
planner item SHA 목록
bundle SHA
Resolver resolution SHA
Family resolution SHA
Feasibility resolution SHA
Ranking resolution SHA
re-entry item 목록
```

## 5. 보안 경계

영구 캐시에는 다음을 저장하지 않습니다.

```text
원본 파일명
sampleValues
rawRows
원본 파일 또는 원본 파일 바이트
평문 JSON 파일
평문 임시 파일
```

Semantic profile은 L4 저장 전에 persistent-cache 전용 사본으로 정리합니다. 파일명은 빈 값으로 치환하고 `sampleValues`, `rawRows`, 원본 파일 필드는 제거한 뒤 profile SHA를 다시 계산합니다.

영구 파일은 `.enc`만 허용하며 Patch 13의 AES-256-GCM codec과 AAD 경계를 그대로 사용합니다.

## 6. 캐시 허용 조건

Planner resolution은 다음 조건을 모두 만족할 때만 L3에 저장합니다.

```text
Planner resolution validation PASS
invocation = CALLED 또는 CACHE_HIT
failureCode 없음
accepted proposal 1개 이상
privacy boundary PASS
```

L4 re-entry artifact는 다음 조건을 모두 만족할 때만 저장합니다.

```text
Resolver validation PASS
Family validation PASS
Feasibility validation PASS
Ranking validation PASS
accepted = resolved = ready = ranked
persistent privacy boundary PASS
productionCandidateMerge = false
productionReadyAssignment = false
productionRouteChanged = false
```

`FAILED_SAFE`, failureCode가 있는 결과, 일부 proposal만 READY가 된 결과는 저장하지 않습니다.

## 7. 손상 캐시 처리

L4 암호문 인증 실패, payload SHA 불일치, artifact SHA 불일치 또는 결정론적 validation 실패 시 해당 항목을 삭제하고 다음처럼 안전 복귀합니다.

```text
Planner provider result L3 CACHE_HIT
→ 손상 L4 삭제
→ Resolver → Family → Feasibility → Ranker 재실행
→ 검증 성공 시 L4 암호화 재저장
```

Provider를 다시 호출하지 않으며 production 후보군으로 병합하지 않습니다.

## 8. 변경 파일

```text
automation/queryCandidatePlannerShadowBridge.js

tests/queryCandidatePatch13_1TestSupport.js
tests/queryCandidatePatch13_1ShadowReentryCacheIntegrationSmokeTest.js
tests/queryCandidatePatch13_1CorruptFallbackSmokeTest.js
tests/queryCandidatePatch13_1PrivacyBoundarySmokeTest.js
tests/queryCandidatePatch13_1SourceIntegritySmokeTest.js
tests/queryCandidatePatch13_1ManifestSmokeTest.js

PATCH_VALIDATION_PATCH13_1.json
README_QUERY_CANDIDATE_PATCH13_1_SHADOW_REENTRY_CACHE_INTEGRATION.md
```

다음 파일은 변경하지 않습니다.

```text
automation/queryCandidatePlanner.js
automation/queryCandidatePlannerHierarchicalEncryptedCache.js
automation/queryCandidateResolver.js
automation/queryCandidateFamilyResolver.js
automation/queryCandidateFeasibilityGate.js
automation/queryCandidateRanker.js
production route
기존 fixture와 baseline
```

## 9. 적용

저장소 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch13_1_shadow_reentry_cache_integration.zip `
  -DestinationPath . `
  -Force
```

## 10. 문법 검사

```powershell
node --check .\automation\queryCandidatePlannerShadowBridge.js
node --check .\tests\queryCandidatePatch13_1TestSupport.js
node --check .\tests\queryCandidatePatch13_1ShadowReentryCacheIntegrationSmokeTest.js
node --check .\tests\queryCandidatePatch13_1CorruptFallbackSmokeTest.js
node --check .\tests\queryCandidatePatch13_1PrivacyBoundarySmokeTest.js
node --check .\tests\queryCandidatePatch13_1SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch13_1ManifestSmokeTest.js
```

## 11. Patch 13.1 검사

실제 API 호출 없이 mock provider로 실행됩니다.

```powershell
node .\tests\queryCandidatePatch13_1ShadowReentryCacheIntegrationSmokeTest.js
node .\tests\queryCandidatePatch13_1CorruptFallbackSmokeTest.js
node .\tests\queryCandidatePatch13_1PrivacyBoundarySmokeTest.js
node .\tests\queryCandidatePatch13_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13_1ManifestSmokeTest.js
```

정상 결과:

```text
PASS query candidate patch13.1 shadow re-entry cache integration smoke
PASS query candidate patch13.1 corrupt fallback smoke
PASS query candidate patch13.1 privacy boundary smoke
PASS query candidate patch13.1 source integrity smoke
PASS query candidate patch13.1 manifest smoke
```

통합 검사에서 확인하는 핵심값:

```text
첫 실행 invocation                 CALLED
두 번째 invocation                 CACHE_HIT
새 cache instance invocation       CACHE_HIT
총 provider call                   1
동일 프로세스 L3/L4 source         L1_MEMORY
새 cache instance Planner source   L3_SEMANTIC
새 cache instance Re-entry source  L4_REENTRY
accepted/resolved/ready/ranked      2/2/2/2
production merge                    false
production route changed            false
```

## 12. Patch 13 및 누적 QA

```powershell
node .\tests\queryCandidatePatch13HierarchicalCacheKeySmokeTest.js
node .\tests\queryCandidatePatch13EncryptedHierarchySmokeTest.js
node .\tests\queryCandidatePatch13CachePolicySmokeTest.js
node .\tests\queryCandidatePatch13PlannerAdapterSmokeTest.js
node .\tests\queryCandidatePatch13SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13ManifestSmokeTest.js
```

Patch 12 누적 QA도 기존 명령으로 모두 PASS여야 합니다.

## 13. 기준선 비교

기준선을 다시 작성하지 않고 비교만 실행합니다.

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

## 14. 다음 단계

Patch 13.1 검사가 전부 통과하면 다음은 Patch 13.2에서 TTL·무효화·운영 제어를 Shadow 캐시 경로에 추가합니다. 아직 production promotion이나 production 후보 병합은 진행하지 않습니다.
