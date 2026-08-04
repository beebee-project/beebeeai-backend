# Query Candidate Patch 14.0 — Feature Flags + Kill Switch

## 적용 선행조건

Patch 13.3의 Live Provider cache-hit parity와 Production Readiness Gate가 모두 PASS인 뒤 적용합니다.

```text
Patch 13.3
→ Patch 14.0 — Feature Flags + Kill Switch
```

Patch 13.3은 `ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW`까지만 판정하며 Production route를 자동으로 연결하지 않습니다. Patch 14.0도 기존 Production route를 변경하지 않습니다.

## 목적

후속 API·UI 연결 및 Controlled Production Promotion 전에 중앙 운영 제어 계약을 먼저 고정합니다.

```text
기본값                         신규 Planner feature 비활성
Production kill switch         기본 활성
환경변수 누락                  안전한 기본값 사용
환경변수 값 오류               전체 기능 fail-closed
런타임 Kill Switch 활성화      즉시 다음 evaluate()부터 차단
환경변수 Kill Switch           런타임 코드로 해제 불가
Patch 13.3 Readiness 증거 없음 Production 작업 차단
Production route 변경          없음
Production 후보 병합           없음
Production READY 부여          없음
```

## 변경 파일

```text
automation/queryCandidatePlannerFeatureControl.js
automation/queryCandidatePlannerFeatureControl.schema.json

tests/queryCandidatePatch14_0TestSupport.js
tests/queryCandidatePatch14_0DefaultFailClosedSmokeTest.js
tests/queryCandidatePatch14_0FlagMatrixSmokeTest.js
tests/queryCandidatePatch14_0KillSwitchPrecedenceSmokeTest.js
tests/queryCandidatePatch14_0InvalidEnvironmentSmokeTest.js
tests/queryCandidatePatch14_0ReadinessEvidenceSmokeTest.js
tests/queryCandidatePatch14_0SourceIntegritySmokeTest.js
tests/queryCandidatePatch14_0ManifestSmokeTest.js

PATCH_VALIDATION_PATCH14_0.json
PATCH_MANIFEST_PATCH14_0.json
README_QUERY_CANDIDATE_PATCH14_0_FEATURE_FLAGS_KILL_SWITCH.md
```

기존 `queryCandidatePlannerShadowBridge.js`, Production route, API controller, UI는 수정하지 않습니다.

## 환경변수

### 기능 플래그

```powershell
$env:QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED = "1"
$env:QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_CACHE_READ_ENABLED = "1"
$env:QUERY_CANDIDATE_PLANNER_CACHE_WRITE_ENABLED = "1"

$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED = "0"
```

### Kill Switch

```powershell
$env:QUERY_CANDIDATE_PLANNER_KILL_SWITCH = "0"
$env:QUERY_CANDIDATE_PLANNER_PROVIDER_KILL_SWITCH = "0"
$env:QUERY_CANDIDATE_PLANNER_CACHE_KILL_SWITCH = "0"
$env:QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH = "1"
```

Boolean 값은 `0`, `1`, `false`, `true`만 허용합니다. 다른 값은 전체 제어를 fail-closed로 전환합니다.

## Operation 계약

```text
SHADOW_EXECUTION
PROVIDER_CALL
CACHE_READ
CACHE_WRITE
PRODUCTION_CANDIDATE_MERGE
PRODUCTION_READY_ASSIGNMENT
PRODUCTION_ROUTE
```

우선순위:

```text
잘못된 환경변수
→ Global Kill Switch
→ Feature Enabled
→ Scope Kill Switch
→ Operation별 Feature Flag
→ Patch 13.3 Readiness Evidence
→ ALLOW
```

## 런타임 즉시 차단

```javascript
const {
  SCOPES,
  createQueryCandidatePlannerFeatureControl,
} = require("./automation/queryCandidatePlannerFeatureControl");

const control = createQueryCandidatePlannerFeatureControl();

control.activateKillSwitch({
  scope: SCOPES.GLOBAL,
  reason: "INCIDENT_RESPONSE",
  actor: "on-call",
});
```

같은 프로세스에서 이후 `evaluate()` 호출부터 즉시 `GLOBAL_KILL_SWITCH_ACTIVE`로 차단됩니다.

환경변수로 활성화한 Kill Switch는 코드의 `releaseRuntimeKillSwitch()`로 해제할 수 없습니다. Railway 환경변수 또는 배포 설정에서 직접 해제해야 합니다.

## 적용

백엔드 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch14_0_feature_flags_kill_switch.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check .\automation\queryCandidatePlannerFeatureControl.js
node --check .\tests\queryCandidatePatch14_0TestSupport.js
node --check .\tests\queryCandidatePatch14_0DefaultFailClosedSmokeTest.js
node --check .\tests\queryCandidatePatch14_0FlagMatrixSmokeTest.js
node --check .\tests\queryCandidatePatch14_0KillSwitchPrecedenceSmokeTest.js
node --check .\tests\queryCandidatePatch14_0InvalidEnvironmentSmokeTest.js
node --check .\tests\queryCandidatePatch14_0ReadinessEvidenceSmokeTest.js
node --check .\tests\queryCandidatePatch14_0SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch14_0ManifestSmokeTest.js
```

## Patch 14.0 QA

실제 OpenAI API를 호출하지 않으므로 추가 API 비용은 발생하지 않습니다.

```powershell
node .\tests\queryCandidatePatch14_0DefaultFailClosedSmokeTest.js
node .\tests\queryCandidatePatch14_0FlagMatrixSmokeTest.js
node .\tests\queryCandidatePatch14_0KillSwitchPrecedenceSmokeTest.js
node .\tests\queryCandidatePatch14_0InvalidEnvironmentSmokeTest.js
node .\tests\queryCandidatePatch14_0ReadinessEvidenceSmokeTest.js
node .\tests\queryCandidatePatch14_0SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_0ManifestSmokeTest.js
```

정상 출력:

```text
PASS query candidate patch14.0 default fail-closed smoke
PASS query candidate patch14.0 flag matrix smoke
PASS query candidate patch14.0 kill-switch precedence smoke
PASS query candidate patch14.0 invalid environment fail-closed smoke
PASS query candidate patch14.0 readiness evidence smoke
PASS query candidate patch14.0 source integrity smoke
PASS query candidate patch14.0 manifest smoke
```

## 기존 회귀

Patch 13.3·13.2·13.1·13 검사와 기존 Shadow/Planner/Ranker/Feasibility/Family/Resolver `--mode=compare`를 다시 실행합니다. 기준선은 다시 작성하지 않습니다.

## 완료 판정

```text
중앙 Feature Control 계약               추가 완료
기본 Production 접근                    차단
환경변수 오류                           fail-closed
Global/Provider/Cache/Production Switch 분리
Runtime 즉시 차단                       검증
Environment Switch 우선권               검증
Patch 13.3 Readiness 증거                필수
Production route                        미변경
API/UI 연결                             미수행
```

Patch 14.0 통과 후 다음 단계에서 실제 backend 후보군 service 경계에 `evaluate()`를 연결합니다. API/UI에 직접 연결하기 전에 Shadow 경로와 기존 결정론적 fallback이 그대로 유지되는지 별도 통합 스모크로 확인해야 합니다.
