# Query Candidate Patch 14.2

## Upload Lifecycle + Cache Invalidation Wiring

### 적용 순서

```text
Patch 13.2  Cache TTL / upload invalidation foundation
Patch 14.0  Feature Flags + Kill Switch
Patch 14.1  API Shadow Wiring + Shadow Comparator
Patch 14.2  Upload Lifecycle + Cache Invalidation Wiring
```

이 ZIP은 wrapper 폴더 없이 `beebeeai-backend` 저장소 루트에 바로 덮어쓰는 구조입니다.

## 목적

Patch 14.1에서 연결한 API Shadow 경로에 Patch 13.2의 암호화 계층형 캐시를 실제 파일 수명주기와 연결합니다.

```text
동일 업로드의 후보군 재실행
→ 동일 upload/queryJson SHA 태그
→ L3 Planner / L4 Re-entry 캐시 재사용 가능

원본 또는 생성 결과 다운로드
→ 캐시 유지
→ 다시 보기·재생성에서 재사용 가능

업로드 파일 삭제
→ 해당 upload/queryJson SHA 태그의 L3/L4 캐시 무효화
→ 파일·queryJson 삭제는 기존 Controller가 수행

같은 이름으로 파일 교체 업로드
→ 기존 업로드 캐시 먼저 무효화
→ 새 queryJson/storage 객체로 새 SHA 태그 생성
→ 과거 캐시 재사용 금지

삭제 후 같은 파일 재업로드
→ 새 업로드 객체이므로 새 SHA 태그
→ 이전 캐시는 재사용하지 않음
```

## 핵심 계약

```text
Cache runtime              query_candidate_planner_cache_runtime_v1
Upload lifecycle           query_candidate_planner_upload_lifecycle_v1
Upload identity            query_candidate_planner_upload_identity_v1
File lifecycle boundary    query_candidate_planner_file_lifecycle_boundary_v1
Invalidation wiring        query_candidate_planner_upload_invalidation_wiring_v1
```

## API Shadow 캐시 연결

`/api/automation/analysis-candidates`의 기존 Primary 응답은 그대로 유지됩니다.

Shadow 실행 시에만 내부적으로 다음 값을 Patch 13.2 Shadow Bridge에 전달합니다.

```javascript
{
  hierarchicalCache,
  tenantId,
  cacheSecret,
  uploadFingerprintSha256,
  queryJsonSha256,
}
```

관찰 로그와 API 응답에는 다음 원문을 남기지 않습니다.

```text
tenantId
cacheSecret
원본 파일명
queryTablesKey
storage object key
원본 행·표본 값
```

외부 관찰에는 SHA-256 태그와 상태만 남습니다.

## 업로드 식별자 규칙

우선순위는 다음과 같습니다.

```text
1. queryTablesKey / fileInfo.queryJsonKey
2. storage object key
3. fileHash + sheetStateSig fallback
```

내부 원문 객체 식별자를 tenant 범위와 함께 SHA-256으로 변환합니다.

```text
같은 업로드 + 같은 queryJson 객체  → 같은 태그
같은 이름의 새 업로드 객체         → 다른 태그
삭제 후 재업로드                   → 다른 태그
```

파일 내용이 같아도 삭제 후 재업로드는 새 객체로 취급하므로 과거 Provider 결과를 자동 재사용하지 않습니다.

## 파일 라우트 수명주기

기존 라우트 path와 Controller는 유지하고 경계 wrapper만 추가합니다.

```text
POST   /api/files/upload
GET    /api/files/download/:originalName
DELETE /api/files/:originalName
GET    /api/automation/download
```

### 최초 업로드

이전 업로드가 없으므로 무효화하지 않습니다.

```text
cacheDisposition = NO_PREVIOUS_UPLOAD
```

### 같은 이름 교체 업로드

기존 파일 Controller가 같은 이름 파일을 교체하기 전에 이전 업로드 식별자를 포착하고 캐시를 무효화합니다.

```text
cacheDisposition = INVALIDATED
reason           = UPLOAD_DELETED
```

새 업로드가 저장되면 새 queryJsonKey가 생성되어 새 캐시 identity를 사용합니다.

### 다운로드

원본 파일과 생성 결과 다운로드 모두 캐시를 삭제하지 않습니다.

```text
cacheDisposition = RETAINED
reason           = DOWNLOAD_DOES_NOT_INVALIDATE_CACHE
```

생성된 복호화 결과 파일의 기존 다운로드 후 정리 정책은 그대로 유지합니다. Planner의 암호화 캐시와 queryJson은 별개이므로 다운로드 정리의 영향을 받지 않습니다.

### 업로드 삭제

기존 파일 삭제 전에 해당 업로드 태그의 Planner 캐시를 무효화합니다.

```text
L3 PLANNER_PROVIDER_RESULT
L3 PLANNER_RESOLUTION
L4 SHADOW_REENTRY
```

무효화 후 기존 Controller가 원본 저장 객체, 암호화 queryJson, 사용자 파일 목록을 삭제합니다.

## 기본 비활성 상태

패치 적용 직후 Railway 환경변수를 변경하지 않습니다.

```text
Patch 14.0 Feature Flag 기본 OFF
Cache runtime key 미설정
실제 Provider 호출 없음
실제 Cache read/write 없음
기존 API/UI 동작 유지
```

Cache runtime이 구성되지 않은 상태에서 파일을 삭제하거나 교체하면 파일 API는 정상 실행되고 다음 상태만 기록합니다.

```text
cacheDisposition = NO_ACTIVE_CACHE
reason           = CACHE_RUNTIME_ENV_NOT_CONFIGURED
```

이는 오류가 아니라 아직 캐시가 활성화되지 않았다는 뜻입니다.

## 향후 내부 Shadow 활성화 시 환경변수

Patch 15의 내부 평가·Allowlist 단계 전까지 설정하지 않습니다.

```text
QUERY_CANDIDATE_PLANNER_CACHE_KEY
QUERY_CANDIDATE_PLANNER_CACHE_SECRET
QUERY_CANDIDATE_PLANNER_CACHE_KEY_ID
QUERY_CANDIDATE_PLANNER_CACHE_ROOT
```

권장 사항:

```text
CACHE_KEY       32-byte 랜덤 키의 64자리 hex 또는 base64
CACHE_SECRET    tenant HMAC용 별도 랜덤 secret
CACHE_KEY_ID    운영 키 버전 식별자
CACHE_ROOT      암호화 캐시 저장 경로
```

## 오류 격리

캐시 무효화가 실패해도 업로드·삭제 응답 payload와 status를 변경하지 않습니다.

```text
cacheDisposition = INVALIDATION_FAILED_SAFE
```

Shadow 캐시 runtime 초기화 실패도 Primary 후보군 응답에 전파하지 않습니다.

## Production 격리

이번 패치에서도 다음 작업은 하지 않습니다.

```text
Production 후보 병합
Production READY 승격
Production route 변경
UI 후보 교체
Rollout percentage 적용
```

## 변경 파일

```text
automation/queryCandidatePlannerCacheRuntime.js
automation/queryCandidatePlannerUploadLifecycle.js
automation/queryCandidatePlannerUploadLifecycle.schema.json
automation/queryCandidatePlannerFileLifecycleBoundary.js
automation/queryCandidatePlannerApiShadowRunner.js
automation/queryCandidatePlannerApiShadowService.js
automation/queryCandidatePlannerApiShadowObservation.schema.json

routes/fileRoutes.js
routes/automationRoutes.js

tests/queryCandidatePatch14_1ManifestSmokeTest.js
tests/queryCandidatePatch14_1RouteContractSmokeTest.js
tests/queryCandidatePatch14_1SourceIntegritySmokeTest.js

tests/queryCandidatePatch14_2TestSupport.js
tests/queryCandidatePatch14_2UploadIdentitySmokeTest.js
tests/queryCandidatePatch14_2ApiCacheWiringSmokeTest.js
tests/queryCandidatePatch14_2OperationalInvalidationAdapterSmokeTest.js
tests/queryCandidatePatch14_2DownloadRetentionSmokeTest.js
tests/queryCandidatePatch14_2DeleteInvalidationSmokeTest.js
tests/queryCandidatePatch14_2ReuploadInvalidationSmokeTest.js
tests/queryCandidatePatch14_2InvalidationFailureIsolationSmokeTest.js
tests/queryCandidatePatch14_2CacheBlockedBoundarySmokeTest.js
tests/queryCandidatePatch14_2ApiLifecycleObservationSmokeTest.js
tests/queryCandidatePatch14_2CacheRuntimeDefaultSmokeTest.js
tests/queryCandidatePatch14_2DefaultInactiveLifecycleSmokeTest.js
tests/queryCandidatePatch14_2RouteContractSmokeTest.js
tests/queryCandidatePatch14_2SchemaSmokeTest.js
tests/queryCandidatePatch14_2SourceIntegritySmokeTest.js

PATCH_VALIDATION_PATCH14_2.json
PATCH_MANIFEST_PATCH14_2.json
README_QUERY_CANDIDATE_PATCH14_2_UPLOAD_LIFECYCLE_CACHE_WIRING.md
```

## 적용

```powershell
Get-FileHash `
  .\query_candidate_patch14_2_upload_lifecycle_cache_wiring.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch14_2_upload_lifecycle_cache_wiring.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check .\automation\queryCandidatePlannerCacheRuntime.js
node --check .\automation\queryCandidatePlannerUploadLifecycle.js
node --check .\automation\queryCandidatePlannerFileLifecycleBoundary.js
node --check .\automation\queryCandidatePlannerApiShadowRunner.js
node --check .\automation\queryCandidatePlannerApiShadowService.js
node --check .\routes\fileRoutes.js
node --check .\routes\automationRoutes.js

Get-ChildItem .\tests\queryCandidatePatch14_2*.js |
  ForEach-Object { node --check $_.FullName }
```

## Patch 14.2 QA

아래 테스트는 실제 OpenAI Provider를 호출하지 않습니다.

```powershell
node .\tests\queryCandidatePatch14_2UploadIdentitySmokeTest.js
node .\tests\queryCandidatePatch14_2ApiCacheWiringSmokeTest.js
node .\tests\queryCandidatePatch14_2OperationalInvalidationAdapterSmokeTest.js
node .\tests\queryCandidatePatch14_2DownloadRetentionSmokeTest.js
node .\tests\queryCandidatePatch14_2DeleteInvalidationSmokeTest.js
node .\tests\queryCandidatePatch14_2ReuploadInvalidationSmokeTest.js
node .\tests\queryCandidatePatch14_2InvalidationFailureIsolationSmokeTest.js
node .\tests\queryCandidatePatch14_2CacheBlockedBoundarySmokeTest.js
node .\tests\queryCandidatePatch14_2ApiLifecycleObservationSmokeTest.js
node .\tests\queryCandidatePatch14_2CacheRuntimeDefaultSmokeTest.js
node .\tests\queryCandidatePatch14_2DefaultInactiveLifecycleSmokeTest.js
node .\tests\queryCandidatePatch14_2RouteContractSmokeTest.js
node .\tests\queryCandidatePatch14_2SchemaSmokeTest.js
node .\tests\queryCandidatePatch14_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
```

정상 출력:

```text
PASS query candidate patch14.2 upload identity smoke
PASS query candidate patch14.2 API cache wiring smoke
PASS query candidate patch14.2 operational invalidation adapter smoke
PASS query candidate patch14.2 download retention smoke
PASS query candidate patch14.2 delete invalidation smoke
PASS query candidate patch14.2 reupload invalidation smoke
PASS query candidate patch14.2 invalidation failure isolation smoke
PASS query candidate patch14.2 cache blocked boundary smoke
PASS query candidate patch14.2 API lifecycle observation smoke
PASS query candidate patch14.2 cache runtime default smoke
PASS query candidate patch14.2 default inactive lifecycle smoke
PASS query candidate patch14.2 route contract smoke
PASS query candidate patch14.2 schema smoke
PASS query candidate patch14.2 source integrity smoke
PASS query candidate patch14.2 manifest smoke
```

## Patch 14.1 회귀

14.2가 Runner, Service, route를 정식 후속 변경하므로 14.1의 source/route/manifest 검사는 successor-compatible 방식으로 갱신됩니다.

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

핵심 정상 조건:

```text
기존 Primary payload 동일
Shadow 오류·timeout 격리
Production merge false
Production READY assignment false
Production route changed false
Patch 14.1 manifest superseded 항목은 모두 Patch 14.2 manifest에 명시
```

## Patch 13.2 회귀 권장

```powershell
node .\tests\queryCandidatePatch13_2UploadInvalidationSmokeTest.js
node .\tests\queryCandidatePatch13_2ReplaySafeAuditSmokeTest.js
node .\tests\queryCandidatePatch13_2ShadowUploadDeletionBridgeSmokeTest.js
node .\tests\queryCandidatePatch13_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13_2ManifestSmokeTest.js
```

## 실제 운영 수명주기 검증 시점

환경변수를 열지 않은 현재 단계에서는 비용 없는 mock 검증만 수행합니다.

실제 암호화 캐시와 Provider를 사용한 다음 검증은 Patch 15.1 비용·Cache-Hit·Latency Evaluation에서 수행합니다.

```text
1회차 동일 업로드 Shadow        Provider CALLED
2회차 동일 업로드 Shadow        CACHE_HIT / Provider 0회
다운로드 후 동일 업로드 Shadow  CACHE_HIT / Provider 0회
삭제                              L3/L4 invalidated
삭제 후 재업로드 Shadow          새 identity / Provider CALLED 가능
```

`--mode=write`로 기존 기준선을 다시 작성하지 않습니다.
