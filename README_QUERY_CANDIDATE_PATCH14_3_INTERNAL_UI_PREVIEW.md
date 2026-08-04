# Patch 14.3 — Internal UI Preview

## 1. 목적

Patch 14.1의 API Shadow 관찰 결과를 내부 사용자가 읽기 전용 화면에서 확인할 수 있도록 연결합니다.

```text
POST /api/automation/analysis-candidates
→ 기존 Primary 응답 먼저 전송
→ Shadow 관찰 비동기 실행
→ 정제된 observation을 메모리 Ring Buffer에 기록
→ 내부 Preview 페이지에서 조회
```

일반 사용자 후보군 UI, 후보 선택, 다운로드, 생성 실행, Production 병합 경로는 수정하지 않습니다.

## 2. 고정 안전 계약

```text
Preview 기본값                         OFF
기존 BeeBee AI 인증                    필수
내부 Preview Token                    필수
Token query string 전달               금지
관찰 데이터 저장                       MEMORY_ONLY
서버 재시작 시 관찰 데이터             삭제
후보 실행 버튼                         없음
후보 선택 버튼                         없음
Production 후보 병합                  없음
Production READY 승격                 없음
기존 /analysis-candidates 응답 변경    없음
```

`productionRouteChanged: false`는 기존 후보군 Production 응답·실행 경로를 변경하지 않는다는 뜻입니다. Patch 14.3은 별도의 내부 읽기 전용 GET 경로만 추가합니다.

## 3. 내부 Preview 경로

백엔드가 `http://localhost:3000`에서 실행 중인 경우:

```text
http://localhost:3000/api/automation/internal/query-candidate-shadow-preview
```

JSON 경로:

```text
GET /api/automation/internal/query-candidate-shadow-preview/status
GET /api/automation/internal/query-candidate-shadow-preview/observations
```

JSON 경로는 다음 헤더가 필요합니다.

```text
x-beebee-internal-preview-token: <내부 토큰>
```

페이지는 토큰을 URL이나 `localStorage`에 저장하지 않습니다. 현재 탭의 `sessionStorage`에만 유지하고 탭을 닫으면 삭제됩니다.

현재 `automationRoutes.js`의 기존 `router.use(protect)` 뒤에 배치되므로 BeeBee AI 인증도 함께 필요합니다. 직접 URL을 열 때 401이 반환되면 먼저 같은 브라우저에서 BeeBee AI에 로그인한 상태인지 확인합니다.

## 4. 표시 항목

```text
관찰 시각
Shadow 상태 및 차단·실패 사유
Comparator verdict
Primary / Shadow / Shared 후보 수
Top-1 일치 여부
Top-3 overlap
Jaccard
Rank agreement
Provider 호출 수
Cache read/write 허용 상태
Shadow latency
정제된 observation JSON
```

후보 ID 원문, 원본 행, 샘플 값, 파일명, queryTablesKey, tenantId, 이메일은 표시하거나 저장하지 않습니다.

## 5. 메모리 보관 정책

기본값:

```text
최대 관찰 건수      100
TTL                 24시간
최대 설정 가능 건수 500
Persistence         MEMORY_ONLY
```

환경변수:

```text
QUERY_CANDIDATE_INTERNAL_PREVIEW_MAX_ENTRIES
QUERY_CANDIDATE_INTERNAL_PREVIEW_TTL_MS
```

이 저장소는 평가 데이터셋이나 장기 로그가 아닙니다. Patch 15.2 Shadow Accuracy Evaluation 전까지 내부 화면 확인용으로만 사용합니다.

## 6. 적용

백엔드 루트에서 실행합니다.

```powershell
Get-FileHash `
  .\query_candidate_patch14_3_internal_ui_preview.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch14_3_internal_ui_preview.zip `
  -DestinationPath . `
  -Force
```

## 7. 문법 검사

```powershell
node --check .\automation\queryCandidatePlannerInternalPreviewConfig.js
node --check .\automation\queryCandidatePlannerInternalPreviewStore.js
node --check .\automation\queryCandidatePlannerInternalPreviewAccess.js
node --check .\automation\queryCandidatePlannerInternalPreviewPage.js
node --check .\automation\queryCandidatePlannerInternalPreviewController.js
node --check .\routes\automationRoutes.js

Get-ChildItem .\tests\queryCandidatePatch14_3*.js |
  ForEach-Object { node --check $_.FullName }
```

## 8. Patch 14.3 QA

실제 Provider를 호출하지 않습니다.

```powershell
node .\tests\queryCandidatePatch14_3DefaultDisabledSmokeTest.js
node .\tests\queryCandidatePatch14_3AccessControlSmokeTest.js
node .\tests\queryCandidatePatch14_3StorePrivacySmokeTest.js
node .\tests\queryCandidatePatch14_3RetentionSmokeTest.js
node .\tests\queryCandidatePatch14_3FilterSmokeTest.js
node .\tests\queryCandidatePatch14_3ObservationWiringSmokeTest.js
node .\tests\queryCandidatePatch14_3ControllerSmokeTest.js
node .\tests\queryCandidatePatch14_3PageContractSmokeTest.js
node .\tests\queryCandidatePatch14_3SecurityHeadersSmokeTest.js
node .\tests\queryCandidatePatch14_3ReadOnlyIsolationSmokeTest.js
node .\tests\queryCandidatePatch14_3RouteContractSmokeTest.js
node .\tests\queryCandidatePatch14_3SchemaSmokeTest.js
node .\tests\queryCandidatePatch14_3SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_3DisabledRecorderNoopSmokeTest.js
node .\tests\queryCandidatePatch14_3ManifestSmokeTest.js
```

정상 출력:

```text
PASS query candidate patch14.3 default disabled smoke
PASS query candidate patch14.3 access control smoke
PASS query candidate patch14.3 store privacy smoke
PASS query candidate patch14.3 retention smoke
PASS query candidate patch14.3 filter smoke
PASS query candidate patch14.3 observation wiring smoke
PASS query candidate patch14.3 controller smoke
PASS query candidate patch14.3 page contract smoke
PASS query candidate patch14.3 security headers smoke
PASS query candidate patch14.3 read-only isolation smoke
PASS query candidate patch14.3 route contract smoke
PASS query candidate patch14.3 schema smoke
PASS query candidate patch14.3 source integrity smoke
PASS query candidate patch14.3 disabled recorder no-op smoke
PASS query candidate patch14.3 manifest smoke
```

## 9. Patch 14.1·14.2 호환성 검사

```powershell
node .\tests\queryCandidatePatch14_1RouteContractSmokeTest.js
node .\tests\queryCandidatePatch14_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js

node .\tests\queryCandidatePatch14_2RouteContractSmokeTest.js
node .\tests\queryCandidatePatch14_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_2SchemaSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch14_2_1BoundaryRestoreSmokeTest.js
node .\tests\queryCandidatePatch14_2_2ProtectedIntegrityConvergenceSmokeTest.js
```

Patch 14.2 Manifest 정상 마지막 줄:

```text
PASS query candidate patch14.2 manifest smoke superseded=2
```

Patch 14.3이 정식으로 대체하는 Patch 14.2 파일은 다음 두 개뿐입니다.

```text
routes/automationRoutes.js
tests/queryCandidatePatch14_2ManifestSmokeTest.js
```

## 10. 내부 Preview 활성화

QA 전에는 환경변수를 설정하지 않습니다.

PowerShell에서 32바이트 난수 토큰을 생성합니다.

```powershell
$bytes = New-Object byte[] 32
[Security.Cryptography.RandomNumberGenerator]::Fill($bytes)
$previewToken = [Convert]::ToHexString($bytes)
$previewToken
```

로컬 내부 Preview만 켭니다.

```powershell
$env:QUERY_CANDIDATE_INTERNAL_PREVIEW_ENABLED = "1"
$env:QUERY_CANDIDATE_INTERNAL_PREVIEW_TOKEN = $previewToken
```

Shadow 관찰을 비용 없이 확인하려면 다음 상태를 사용합니다.

```powershell
$env:QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED = "1"
$env:QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED = "1"
$env:QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED = "0"
```

이 상태에서는 실제 Provider 호출이 차단됩니다. 기존 Production Feature Flag와 Production Kill Switch는 변경하지 않습니다.

실제 Provider Shadow는 별도의 비용 승인과 내부 검증에서만 활성화합니다.

## 11. 비활성화

```powershell
Remove-Item Env:\QUERY_CANDIDATE_INTERNAL_PREVIEW_ENABLED -ErrorAction SilentlyContinue
Remove-Item Env:\QUERY_CANDIDATE_INTERNAL_PREVIEW_TOKEN -ErrorAction SilentlyContinue
```

비활성화 상태:

```text
Preview 페이지 404
Preview JSON API 404
Observation recorder no-op
기존 Primary API 정상 유지
```

## 12. 변경 파일

```text
automation/queryCandidatePlannerInternalPreviewConfig.js
automation/queryCandidatePlannerInternalPreviewStore.js
automation/queryCandidatePlannerInternalPreviewAccess.js
automation/queryCandidatePlannerInternalPreviewPage.js
automation/queryCandidatePlannerInternalPreviewController.js
automation/queryCandidatePlannerInternalPreview.schema.json
routes/automationRoutes.js
tests/queryCandidatePatch14_2ManifestSmokeTest.js
tests/queryCandidatePatch14_3*.js
PATCH_VALIDATION_PATCH14_3.json
PATCH_MANIFEST_PATCH14_3.json
README_QUERY_CANDIDATE_PATCH14_3_INTERNAL_UI_PREVIEW.md
```

## 13. 다음 단계

```text
Patch 14.4 — Controlled Production Merge Adapter
기본값 OFF
```

Patch 14.3 화면에는 Merge Adapter를 호출하거나 Gate를 여는 기능이 없습니다.
