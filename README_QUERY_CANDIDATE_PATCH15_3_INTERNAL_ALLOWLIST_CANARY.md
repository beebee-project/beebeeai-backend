# Query Candidate Patch 15.3 — Internal Allowlist Canary

## 목적

Patch 15.3은 신규 Query Candidate Planner 결과를 기존 `/analysis-candidates` 응답에 처음으로 제한 적용할 수 있는 내부 Canary 경로를 추가한다.

코드 설치와 실제 활성화는 분리한다.

- 설치 직후: 모든 요청 `BLOCKED`, 기존 Primary 후보군 반환
- 내부 활성화 후: Allowlist에 포함된 내부 계정만 Controlled Merge 가능
- 일반 사용자: 항상 기존 Primary 후보군 반환
- Rollout: 0% 고정
- Production READY 할당: 금지
- Production route 변경: 없음
- LLM 정책: `SEMANTIC_PROFILER_ONLY`
- 조건부 Planner LLM escalation: 금지

## 요청 경로

```text
POST /api/automation/analysis-candidates
→ 기존 Primary 후보군 생성
→ Internal Canary Preflight
   ├─ BLOCK: Primary 즉시 반환 + 기존 Shadow 관찰 비동기 유지
   └─ ALLOW:
      → Shadow Planner 1회 실행
      → Comparator
      → Promotion Gate
      → Controlled Production Merge Adapter
      → 성공 시 merged copy 반환
      → 오류·Timeout·Guardrail 위반 시 Primary fallback
```

기존 Controller는 변경하지 않는다. `routes/automationRoutes.js`의 기존 `/analysis-candidates` 서비스 경계만 Canary Boundary로 교체한다.

## 기본 안전 상태

```text
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED=0
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH=1
QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE=BLOCKED
QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH=1
```

따라서 패치 적용만으로 실제 Production Merge가 발생하지 않는다.

## Allowlist 식별자

원본 이메일이나 이름을 Allowlist에 저장하지 않는다. 로그인 사용자의 불변 account ID와 선택적 tenant ID를 정규화한 뒤 SHA-256으로 변환한다.

```powershell
node .\scripts\queryCandidatePlannerCanarySubjectHash.js `
  <immutableAccountId> `
  <tenantId>
```

출력된 64자리 SHA-256만 다음 환경변수에 등록한다.

```text
QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256=<sha256>
```

## 실제 Shadow 증거 필수

합성 데이터셋 PASS만으로 Canary를 열 수 없다. `QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_EVIDENCE_JSON`에는 다음 조건을 충족하는 실제 Shadow 증거 Bundle이 필요하다.

```json
{
  "version": "query_candidate_planner_internal_canary_evidence_bundle_v1",
  "source": "REAL_SHADOW_TRAFFIC",
  "synthetic": false,
  "actualTraffic": true,
  "evaluatedAt": "2026-08-05T05:30:00.000Z",
  "expiresAt": "2026-08-06T05:30:00.000Z",
  "readiness": {
    "eligible": true,
    "decision": "ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW",
    "guardrails": {
      "manualPromotionReviewRequired": true,
      "failClosed": true,
      "productionRouteAutoWired": false,
      "productionCandidateMergeAllowed": false,
      "productionReadyAssignmentAllowed": false
    }
  },
  "accuracy": {
    "version": "query_candidate_planner_accuracy_evaluation_report_v1",
    "decision": "EVALUATION_PASS",
    "failClosed": true,
    "evaluationOnly": true,
    "promotionAuthorized": false,
    "sampleSize": 30,
    "reportSha256": "<sha256>"
  },
  "operational": {
    "version": "query_candidate_planner_cost_cache_latency_evaluation_report_v1",
    "decision": "EVALUATION_PASS",
    "failClosed": true,
    "evaluationOnly": true,
    "promotionAuthorized": false,
    "sampleSize": 30,
    "reportSha256": "<sha256>",
    "pricingSource": "APPROVED_ACTUAL"
  },
  "shadow": {
    "version": "query_candidate_planner_shadow_accuracy_evaluation_report_v1",
    "decision": "EVALUATION_PASS",
    "failClosed": true,
    "evaluationOnly": true,
    "promotionAuthorized": false,
    "sampleSize": 30,
    "reportSha256": "<sha256>",
    "primaryResponseUnchangedRate": 1,
    "guardrailViolationCount": 0,
    "privacyViolationCount": 0
  },
  "llmPolicy": {
    "mode": "SEMANTIC_PROFILER_ONLY",
    "plannerEscalationAllowed": false
  }
}
```

증거는 기본 7일 이내이며 `expiresAt`이 지나면 자동 차단된다. 원본 행, 파일명, 이메일, account ID, tenant ID, storage key, prompt를 포함하면 차단된다.

## 최초 활성화 조건

다음 조건을 모두 충족해야 한다.

```text
Internal Canary Enabled = 1
Internal Canary Kill Switch = 0
Feature Enabled = 1
Shadow Enabled = 1
Production Enabled = 1
Production Candidate Merge Enabled = 1
Production Kill Switch = 0
Promotion Gate Enabled = 1
Promotion Audience Mode = ALLOWLIST
Promotion Rollout Percent = 0
Allowlist SHA-256 일치
Patch 13.3 Readiness 증거 유효
Patch 15.0 실제 Accuracy 증거 PASS
Patch 15.1 실제 Cost/Cache/Latency 증거 PASS
Patch 15.2 실제 Shadow 증거 PASS
Semantic Profiler-only 정책
```

일반 사용자는 Allowlist 불일치로 계속 차단된다.

## 활성화 환경변수 예시

실제 증거 확보와 QA 완료 전에는 설정하지 않는다.

```text
QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED=1
QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED=1
QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED=1
QUERY_CANDIDATE_PLANNER_CACHE_READ_ENABLED=1
QUERY_CANDIDATE_PLANNER_CACHE_WRITE_ENABLED=1

QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED=1
QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED=1
QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH=0

QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED=1
QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE=ALLOWLIST
QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256=<internal-subject-sha256>
QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT=0

QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED=1
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH=0
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_TIMEOUT_MS=15000
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_LLM_MODE=SEMANTIC_PROFILER_ONLY
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_EVIDENCE_JSON=<validated-json>
```

## LLM 호출 제한

Canary 경로는 다음 정책을 runner에 전달한다.

```text
mode = SEMANTIC_PROFILER_ONLY
maxProviderCalls = 1
plannerEscalationAllowed = false
```

Shadow 결과가 Provider 호출 2회 이상 또는 Planner escalation 사용을 보고하면 Merge하지 않고 Primary로 복귀한다.

## Rollback

즉시 복귀 우선순위:

```text
1. Runtime Production Kill Switch ON
2. Internal Canary Kill Switch = 1
3. Promotion Audience Mode = BLOCKED
4. Promotion Gate Enabled = 0
5. Production Kill Switch = 1
6. Global Kill Switch = 1
```

Runtime Production Kill Switch를 켜면 다음 요청부터 Allowlist 계정도 Primary 후보군을 받는다.

## 개인정보 경계

Canary observation에 저장·로그 가능한 값:

```text
subjectTagSha256
evidenceSha256
primaryResponseSha256
Gate reason
Allowlist match 여부
Provider call count
Comparator summary
Merge 적용 여부
Latency
Fallback reason
```

포함 금지:

```text
원본 account ID
원본 tenant ID
이메일·이름
원본 엑셀 행·셀 값
원본 파일명
queryTablesKey
storage key
Prompt·Provider raw response
```

## 적용

```powershell
Get-FileHash `
  .\query_candidate_patch15_3_internal_allowlist_canary.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_3_internal_allowlist_canary.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
Get-ChildItem .\automation\queryCandidatePlannerInternal*.js |
  ForEach-Object { node --check $_.FullName }

Get-ChildItem .\tests\queryCandidatePatch15_3*.js |
  ForEach-Object { node --check $_.FullName }

node --check .\routes\automationRoutes.js
node --check .\scripts\queryCandidatePlannerCanarySubjectHash.js
```

## Patch 15.3 QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch15_3*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object { node $_.FullName }
```

## 선행 Manifest 회귀

```powershell
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_4ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_4_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5FullQualityGateSmokeTest.js
node .\tests\queryCandidatePatch15_0ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_2ManifestSmokeTest.js
```

## Patch 15.3에서 하지 않는 작업

```text
일반 사용자 Rollout
1% Traffic 활성화
자동 Rollout 확대
Production READY 할당
Production route 변경
조건부 Planner LLM 호출
사용자 UI 변경
Allowlist UI 편집
자동 Gate 승인
```

일반 사용자 확대는 Patch 15.4에서 별도 단계로 수행한다.
