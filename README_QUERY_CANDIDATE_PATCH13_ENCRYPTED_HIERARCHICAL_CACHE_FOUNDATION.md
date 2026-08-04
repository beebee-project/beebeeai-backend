# Query Candidate Patch 13 — Encrypted Hierarchical Cache Foundation

## 1. 적용 순서

이 패치는 Patch 12.1.2.2 이후에 적용합니다.

```text
Patch 12.1.2.2
→ Patch 13 — Encrypted Hierarchical Cache Foundation
```

이 ZIP은 wrapper 폴더 없이 저장소 루트에 바로 덮어쓰는 구조입니다.

## 2. 패치 목적

조건부 Candidate Planner가 실제 Provider를 호출한 뒤 동일 입력을 반복 처리할 때 다음을 보장하는 암호화 계층형 캐시 기반을 추가합니다.

```text
L1_MEMORY   프로세스 메모리 임시 캐시
L2_UPLOAD   업로드 fingerprint + queryJson SHA 기반 암호화 캐시
L3_SEMANTIC semantic profile + Planner input 기반 암호화 캐시
L4_REENTRY  검증된 proposal set + 결정론적 정책 버전 기반 암호화 캐시
```

Patch 13은 저장 엔진·키 계약·보안 정책·Planner 호환 어댑터를 추가하는 기반 패치입니다.

- 기존 production route는 변경하지 않습니다.
- 기존 `runConditionalCandidatePlanner()`는 수정하지 않습니다.
- 기존 Shadow Bridge는 자동 연결하지 않습니다.
- 기존 기준선은 다시 쓰지 않습니다.
- 새 어댑터를 명시적으로 전달한 테스트·후속 wiring에서만 새 캐시를 사용합니다.

## 3. 새 파일

```text
automation/queryCandidatePlannerHierarchicalEncryptedCache.js
automation/queryCandidatePlannerHierarchicalEncryptedCache.schema.json

tests/queryCandidatePatch13HierarchicalCacheKeySmokeTest.js
tests/queryCandidatePatch13EncryptedHierarchySmokeTest.js
tests/queryCandidatePatch13CachePolicySmokeTest.js
tests/queryCandidatePatch13PlannerAdapterSmokeTest.js
tests/queryCandidatePatch13SourceIntegritySmokeTest.js
tests/queryCandidatePatch13ManifestSmokeTest.js

PATCH_VALIDATION_PATCH13.json
PATCH_MANIFEST_PATCH13.json
README_QUERY_CANDIDATE_PATCH13_ENCRYPTED_HIERARCHICAL_CACHE_FOUNDATION.md
```

## 4. 결정론적 캐시 키

캐시 키는 HMAC-SHA-256으로 생성합니다.

키 identity에는 계층에 따라 다음 값이 포함됩니다.

### L2_UPLOAD

```text
tenant HMAC digest
uploadFingerprintSha256
queryJsonSha256
artifactType
cache key contract version
```

### L3_SEMANTIC

```text
tenant HMAC digest
semanticProfileSha256
plannerInputSha256
model
reasoningEffort
promptVersion
schemaVersion
plannerPolicyVersion
```

기존 Planner의 HMAC cache key를 어댑터로 연결할 때는 `upstreamCacheKeySha256`를 사용할 수 있습니다.

### L4_REENTRY

```text
tenant HMAC digest
plannerProposalSetSha256
resolverPolicyVersion
familyPolicyVersion
feasibilityPolicyVersion
rankerPolicyVersion
```

다음 값이 달라지면 cache key도 달라집니다.

```text
tenant
model
reasoning effort
prompt version
schema version
Planner policy
Resolver/Family/Feasibility/Ranker policy
semantic input SHA
proposal set SHA
```

원문 tenant ID는 경로나 파일명에 저장하지 않습니다.

## 5. 암호화 저장 경계

영구 캐시는 `.enc` 파일만 허용합니다.

기본 codec:

```text
AES-256-GCM
12-byte random IV
GCM authentication tag
AAD = purpose + layer + artifactType + tenantDigest + keyDigest
```

암호화 키는 코드나 캐시 파일에 저장하지 않습니다. `createAes256GcmCacheCodec()`에 외부 key management에서 전달해야 합니다.

저장 순서:

```text
검증된 payload
→ entry envelope 구성
→ UTF-8 메모리 Buffer
→ AES-256-GCM 암호화
→ 암호화된 임시 파일
→ atomic rename
→ .enc 영구 파일
```

평문 JSON 파일과 평문 임시 파일은 생성하지 않습니다.

## 6. 캐시 가능 정책

다음 조건을 모두 만족해야 저장합니다.

```text
metadata.cacheable = true
metadata.validationValid = true
failureCode = 빈 값
privacy boundary 위반 없음
outcomeStatus가 허용 상태
```

허용 상태:

```text
CALLED
CACHE_HIT
VALIDATED
READY
SHADOW_COMPLETED
```

저장 금지 상태:

```text
FAILED_SAFE
REQUIRED_NOT_RUN
SKIPPED
ERROR
credit_balance_exhausted 등 failureCode가 존재하는 결과
schema validation 실패 결과
privacy boundary 위반 결과
```

## 7. TTL·손상·삭제 정책

기본 TTL:

```text
L1_MEMORY    5분
L2_UPLOAD    24시간
L3_SEMANTIC  7일
L4_REENTRY   7일
```

모든 TTL은 생성 시 재정의할 수 있습니다.

만료된 항목:

```text
MISS / EXPIRED 반환
해당 .enc 파일 삭제
```

복호화·인증·JSON·SHA·identity 검증에 실패한 항목:

```text
예외를 production 경로로 전파하지 않음
MISS / CORRUPT_ENTRY 반환
기본값으로 손상 .enc 삭제
Provider 또는 결정론적 체인으로 안전하게 재계산 가능
```

업로드 삭제 시 tenant 단위 또는 identity 단위 무효화를 사용할 수 있습니다.

## 8. 기존 Planner 호환 어댑터

`createPlannerProviderHierarchicalCacheAdapter()`는 기존 Planner가 기대하는 다음 interface를 제공합니다.

```text
cache.get(cacheKey)
cache.set(cacheKey, value)
cache.delete(cacheKey)
```

따라서 기존 `runConditionalCandidatePlanner()`를 수정하지 않고 L3 암호화 캐시를 테스트할 수 있습니다.

정상 흐름:

```text
첫 실행                 CALLED / providerCallCount 1
같은 프로세스 재실행     CACHE_HIT / L1_MEMORY
새 cache instance 재실행 CACHE_HIT / L3_SEMANTIC encrypted disk
총 Provider 호출         1
```

## 9. 적용

```powershell
Expand-Archive `
  .\query_candidate_patch13_encrypted_hierarchical_cache_foundation.zip `
  -DestinationPath . `
  -Force
```

## 10. 문법 검사

```powershell
node --check .\automation\queryCandidatePlannerHierarchicalEncryptedCache.js

node --check .\tests\queryCandidatePatch13HierarchicalCacheKeySmokeTest.js
node --check .\tests\queryCandidatePatch13EncryptedHierarchySmokeTest.js
node --check .\tests\queryCandidatePatch13CachePolicySmokeTest.js
node --check .\tests\queryCandidatePatch13PlannerAdapterSmokeTest.js
node --check .\tests\queryCandidatePatch13SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch13ManifestSmokeTest.js
```

## 11. Patch 13 QA

```powershell
node .\tests\queryCandidatePatch13HierarchicalCacheKeySmokeTest.js
node .\tests\queryCandidatePatch13EncryptedHierarchySmokeTest.js
node .\tests\queryCandidatePatch13CachePolicySmokeTest.js
node .\tests\queryCandidatePatch13PlannerAdapterSmokeTest.js
node .\tests\queryCandidatePatch13SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13ManifestSmokeTest.js
```

정상 예상 결과:

```text
PASS query candidate patch13 hierarchical cache key smoke
PASS query candidate patch13 encrypted hierarchy smoke
PASS query candidate patch13 cache policy smoke
PASS query candidate patch13 planner adapter smoke
PASS query candidate patch13 source integrity smoke
PASS query candidate patch13 manifest smoke
```

## 12. 기존 전체 회귀

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

## 13. Patch 13 완료 후 다음 단계

이 패치가 통과하면 다음은 Patch 13.1입니다.

```text
Patch 13.1 — Shadow Re-entry Cache Integration
```

Patch 13.1에서 수행할 작업:

```text
검증된 Planner resolution을 L3에 명시적으로 연결
검증된 Shadow re-entry artifact를 L4에 연결
첫 실행 CALLED 확인
두 번째 실행 CACHE_HIT 확인
proposal item SHA 동일 확인
Resolver/Family/Feasibility/Ranker 결과 parity 확인
productionCandidateMerge false 유지
productionRouteChanged false 유지
FAILED_SAFE 및 failureCode 결과 미저장 확인
```

Patch 13 기반만으로 production promotion을 진행하지 않습니다. L4 Shadow parity와 invalidation 회귀가 완료된 뒤 production 연결 여부를 별도로 판단합니다.
