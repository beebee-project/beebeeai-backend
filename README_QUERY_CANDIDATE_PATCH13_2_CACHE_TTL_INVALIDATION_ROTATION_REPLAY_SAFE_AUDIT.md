# Query Candidate Patch 13.2

## Cache TTL, Invalidation, Rotation & Replay-Safe Audit

### 적용 순서

`Patch 13.1 — Shadow Re-entry Cache Integration` 이후 적용합니다.

### 목적

Patch 13.1에서 연결한 L3 Planner 및 L4 Shadow Re-entry 암호화 캐시에 운영 수명주기 제어를 추가합니다.

- 만료된 항목의 읽기 시 삭제 및 명시적 sweep
- 업로드 삭제 시 관련 Provider/Planner/Re-entry 캐시 전체 무효화
- tenant·layer·artifact 단위 무효화
- AES-256-GCM 키 교체 시 legacy 복호화 후 primary key 재암호화
- 모델·프롬프트·스키마·정책·운영 epoch 변경 시 결정론적 cache miss
- `CALLED → CACHE_HIT` 결과의 replay-safe 동일성 감사
- payload, 원본 파일명, tenant ID를 운영 audit에 기록하지 않음

Production 후보 병합과 Production route 연결은 이번 패치에서도 하지 않습니다.

## 주요 계약

```text
Operational control  query_candidate_planner_cache_operational_control_v1
Audit event           query_candidate_planner_cache_audit_event_v1
Replay audit          query_candidate_planner_cache_replay_audit_v1
Upload invalidation   query_candidate_planner_upload_invalidation_v1
Encryption            AES-256-GCM
Persistent extension  .enc
```

## TTL

기존 L2/L3/L4 TTL을 유지하면서 다음 동작을 고정합니다.

```text
get()에서 expiresAt 경과
→ 메모리 항목 제거
→ .enc 파일 삭제
→ MISS / EXPIRED 반환
→ TTL_EXPIRED audit 기록
```

명시적 정리도 지원합니다.

```javascript
await hierarchicalCache.sweepExpired({
  tenantId,
  cacheSecret,
});
```

Audit에는 payload나 평문 tenant ID가 포함되지 않습니다.

## 업로드 삭제 무효화

Shadow 실행 시 선택적으로 다음 SHA 태그를 전달합니다.

```javascript
await runCandidatePlannerLiveShadow({
  ...input,
  hierarchicalCache,
  tenantId,
  cacheSecret,
  uploadFingerprintSha256,
  queryJsonSha256,
});
```

해당 태그는 암호화 envelope metadata 내부에만 저장됩니다. 원본 파일명과 원본 파일은 저장하지 않습니다.

업로드 삭제 시:

```javascript
await invalidateCandidatePlannerUploadCache({
  hierarchicalCache,
  tenantId,
  cacheSecret,
  uploadFingerprintSha256,
  queryJsonSha256,
});
```

동일 태그를 가진 다음 항목이 제거됩니다.

```text
L3 PLANNER_PROVIDER_RESULT
L3 PLANNER_RESOLUTION
L4 SHADOW_REENTRY
```

무효화 이후 동일 입력을 다시 실행하면 Provider가 다시 호출되며 새 암호화 캐시가 생성됩니다.

## 계층 및 tenant 무효화

```javascript
await hierarchicalCache.invalidateLayer({
  tenantId,
  cacheSecret,
  layer: "L3_SEMANTIC",
});

hierarchicalCache.invalidateTenant({
  tenantId,
  cacheSecret,
});
```

지원 reason:

```text
TTL_EXPIRED
TENANT_DELETED
UPLOAD_DELETED
LAYER_INVALIDATED
MANUAL_DELETE
CORRUPT_ENTRY
KEY_ROTATED
```

## 암호화 키 교체

새 키는 primary, 이전 키는 legacy로 구성합니다.

```javascript
const codec = createRotatingAes256GcmCacheCodec({
  primary: {
    key: newKey,
    keyId: "cache-key-2026-08",
  },
  legacy: [
    {
      key: previousKey,
      keyId: "cache-key-2026-07",
    },
  ],
});

const hierarchicalCache =
  createEncryptedHierarchicalCandidatePlannerCache({
    rootDir,
    ...codec,
  });
```

Legacy key로 암호화된 항목을 읽으면:

```text
legacy key 복호화
→ payload SHA 및 cache policy 검증
→ primary key로 같은 .enc 경로에 원자적 재암호화
→ KEY_ROTATED audit
→ 정상 cache hit 반환
```

legacy key 목록에 없는 keyId는 `CORRUPT_ENTRY` fail-safe 경로로 처리됩니다.

## 계약 변경에 따른 자동 cache miss

결정론적 key material에 이미 포함된 다음 값이 변경되면 keyDigest가 달라집니다.

```text
model
reasoningEffort
promptVersion
schemaVersion
plannerPolicyVersion
resolverPolicyVersion
familyPolicyVersion
feasibilityPolicyVersion
rankerPolicyVersion
extraIdentity.cacheEpoch
```

따라서 과거 결과를 새 계약에서 재사용하지 않습니다.

## Replay-safe 감사

```javascript
const replayAudit = compareReplaySafeShadowResolutions({
  origin: calledResolution,
  replay: cacheHitResolution,
});
```

성공 조건:

```text
origin status                 SHADOW_COMPLETED
replay status                 SHADOW_COMPLETED
replay invocation             CACHE_HIT
replay providerCallCount      0
accepted planner item SHA     동일
Resolver SHA                  동일
Family SHA                    동일
Feasibility SHA               동일
Ranker SHA                    동일
Production merge              false
Production READY assignment   false
Production route changed      false
```

Response ID, 토큰 사용량, 비용, 원본 파일명은 replay fingerprint에서 제외합니다.

## 변경 파일

```text
automation/queryCandidatePlannerHierarchicalEncryptedCache.js
automation/queryCandidatePlannerCacheOperationalControls.js
automation/queryCandidatePlannerCacheOperationalControls.schema.json
automation/queryCandidatePlannerShadowBridge.js

tests/queryCandidatePatch13ManifestSmokeTest.js
tests/queryCandidatePatch13_1ManifestSmokeTest.js
tests/queryCandidatePatch13_2TtlSweepAuditSmokeTest.js
tests/queryCandidatePatch13_2UploadInvalidationSmokeTest.js
tests/queryCandidatePatch13_2KeyRotationSmokeTest.js
tests/queryCandidatePatch13_2ReplaySafeAuditSmokeTest.js
tests/queryCandidatePatch13_2ShadowUploadDeletionBridgeSmokeTest.js
tests/queryCandidatePatch13_2SourceIntegritySmokeTest.js
tests/queryCandidatePatch13_2ManifestSmokeTest.js

PATCH_VALIDATION_PATCH13_2.json
PATCH_MANIFEST_PATCH13_2.json
README_QUERY_CANDIDATE_PATCH13_2_CACHE_TTL_INVALIDATION_ROTATION_REPLAY_SAFE_AUDIT.md
```

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch13_2_cache_ttl_invalidation_rotation_replay_safe_audit.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check .\automation\queryCandidatePlannerHierarchicalEncryptedCache.js
node --check .\automation\queryCandidatePlannerCacheOperationalControls.js
node --check .\automation\queryCandidatePlannerShadowBridge.js

node --check .\tests\queryCandidatePatch13ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch13_1ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch13_2TtlSweepAuditSmokeTest.js
node --check .\tests\queryCandidatePatch13_2UploadInvalidationSmokeTest.js
node --check .\tests\queryCandidatePatch13_2KeyRotationSmokeTest.js
node --check .\tests\queryCandidatePatch13_2ReplaySafeAuditSmokeTest.js
node --check .\tests\queryCandidatePatch13_2ShadowUploadDeletionBridgeSmokeTest.js
node --check .\tests\queryCandidatePatch13_2SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch13_2ManifestSmokeTest.js
```

## Patch 13.2 검사

실제 OpenAI API를 호출하지 않으므로 추가 API 비용은 발생하지 않습니다.

```powershell
node .\tests\queryCandidatePatch13_2TtlSweepAuditSmokeTest.js
node .\tests\queryCandidatePatch13_2UploadInvalidationSmokeTest.js
node .\tests\queryCandidatePatch13_2KeyRotationSmokeTest.js
node .\tests\queryCandidatePatch13_2ReplaySafeAuditSmokeTest.js
node .\tests\queryCandidatePatch13_2ShadowUploadDeletionBridgeSmokeTest.js
node .\tests\queryCandidatePatch13_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13_2ManifestSmokeTest.js
```

정상 출력:

```text
PASS query candidate patch13.2 TTL sweep audit smoke
PASS query candidate patch13.2 upload invalidation smoke
PASS query candidate patch13.2 key rotation smoke
PASS query candidate patch13.2 replay-safe audit smoke
PASS query candidate patch13.2 shadow upload deletion bridge smoke
PASS query candidate patch13.2 source integrity smoke
PASS query candidate patch13.2 manifest smoke
```

## Patch 13·13.1 회귀 검사

```powershell
node .\tests\queryCandidatePatch13HierarchicalCacheKeySmokeTest.js
node .\tests\queryCandidatePatch13EncryptedHierarchySmokeTest.js
node .\tests\queryCandidatePatch13CachePolicySmokeTest.js
node .\tests\queryCandidatePatch13PlannerAdapterSmokeTest.js
node .\tests\queryCandidatePatch13SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13ManifestSmokeTest.js

node .\tests\queryCandidatePatch13_1ShadowReentryCacheIntegrationSmokeTest.js
node .\tests\queryCandidatePatch13_1CorruptFallbackSmokeTest.js
node .\tests\queryCandidatePatch13_1PrivacyBoundarySmokeTest.js
node .\tests\queryCandidatePatch13_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13_1ManifestSmokeTest.js
```

정상 historical manifest 출력:

```text
PASS query candidate patch13 manifest smoke superseded=2
PASS query candidate patch13.1 manifest smoke superseded=2
```

## 전체 기준선 비교

기준선을 다시 작성하지 않습니다.

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
