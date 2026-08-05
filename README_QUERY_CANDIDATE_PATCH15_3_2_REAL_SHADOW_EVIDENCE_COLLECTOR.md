# Patch 15.3.2 — Real Shadow Evidence Collector & Bundle Builder

## 목적

Patch 15.3.2는 실제 인증된 내부 Allowlist 요청에서 발생한 Shadow 실행을 비식별·암호화 Observation으로 보존하고, Patch 15.0 정확도·Patch 15.1 비용/캐시/지연·Patch 15.2 Shadow 정확도 평가를 하나의 만료형 Patch 15.3 Evidence Bundle로 조립한다.

이 패치는 Internal Canary를 자동으로 활성화하지 않는다.

- Internal Canary 기본 OFF 유지
- 일반 사용자 Production Merge 차단 유지
- Production READY 할당 없음
- Production route 변경 없음
- Semantic Profiler-only 유지
- Planner LLM escalation 금지
- Evidence Builder 결과도 `promotionAuthorized: false`

## 실행 흐름

```text
인증된 내부 Allowlist 요청
→ 기존 Primary 응답을 사용자에게 반환
→ 실제 Shadow Planner 실행
→ Shadow 결과를 메모리에서 일시적으로 안전 캡처
→ API Shadow Observation과 결합
→ 서버 관리 Case Registry로 caseId 매핑
→ MongoDB에 AES-256-GCM 암호화 저장

파일 다운로드·삭제·재업로드
→ 기존 File Lifecycle Boundary Observation
→ 동일 Evidence Store에 암호화 저장

최소 실제 요청 확보
→ 암호화 Observation export
→ Patch 15.0/15.1/15.2 evaluator 실행
→ Evidence Bundle 생성
→ Patch 15.3 validator 재검증
```

## 저장하지 않는 정보

- 엑셀 원본 행과 셀 값
- 샘플 값
- 원본 파일명
- 이메일과 이름
- MongoDB User `_id` 원문
- tenant ID 원문
- Storage Key·queryTablesKey
- Provider prompt·원문 응답
- 암호화 키

저장하는 식별 정보는 `subjectTagSha256`, 요청/업로드 fingerprint SHA-256, 승인된 평가 `caseId`와 `scenarioId`뿐이다. 후보 원문 설명은 저장하지 않고 Manifest 기반 candidate ID만 평가에 사용한다.

## 신규 환경변수

설치 직후에는 아무 변수도 추가하지 않아도 되며 Collector는 OFF이다.

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0
```

실제 수집을 시작할 때만 다음을 설정한다.

```text
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=1
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET=<32자 이상 별도 무작위 비밀값>
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_TTL_DAYS=7
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_MAX_RECORDS=5000
QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON=<한 줄 JSON>
```

`QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256`에 등록된 내부 계정만 수집 대상이다. 별도의 사용자 ID를 요청 Body 또는 임의 Header에서 신뢰하지 않는다.

Evidence secret은 `JWT_SECRET`, `FILE_ENCRYPTION_SECRET`, `QUERY_JSON_SECRET`을 재사용하지 말고 별도 생성한다. 이 값은 저장 데이터 복호화에 필요하므로 수집 기간 중 변경하면 안 된다.

## Case Registry

템플릿:

```text
evaluation/queryCandidatePlannerRealShadowCaseRegistry.template.json
```

Registry는 클라이언트가 보내는 caseId를 신뢰하지 않고 서버가 fingerprint로 평가 case를 결정하기 위한 계약이다.

```json
{
  "version": "query_candidate_planner_real_shadow_case_registry_v1",
  "registryId": "internal_real_shadow_2026_08",
  "cases": [
    {
      "caseId": "seed_sales_ready",
      "scenarioId": "seed_sales_ready_internal_01",
      "requestFingerprintSha256": "64자리 fingerprint",
      "uploadFingerprintSha256": "",
      "expectedColdCostMicrousd": 0,
      "modelId": "semantic_profiler_default"
    }
  ]
}
```

Patch 15.0 Accuracy Dataset의 10개 case를 모두 등록해야 한다. 하나의 case당 실제 Observation을 최소 3건 요구하며, 전체 실행 Observation은 최소 30건이어야 한다.

## Fingerprint 발견 단계

Collector는 OFF로 유지한 상태에서 Shadow만 내부 요청으로 실행한다. Internal Preview에서 다음 값을 확인한다.

- `requestFingerprintSha256`
- `uploadFingerprintSha256`
- 테스트에 사용한 Accuracy Dataset `caseId`

확인한 값을 Registry에 수동 매핑한다. 원본 파일명이나 사용자 ID는 Registry에 넣지 않는다.

## 권장 실제 QA 실행 세트

기본 Threshold를 안정적으로 충족하려면 Patch 15.0의 10개 case 각각에 다음 순서를 수행한다.

```text
1. 최초 업로드 후 후보 조회       COLD
2. 같은 업로드 후보 재조회        WARM 1
3. 같은 업로드 후보 재조회        WARM 2
4. 다운로드
5. 같은 업로드 후보 재조회        DOWNLOAD_REUSE
6. 파일 삭제
7. 동일 파일 재업로드 후 조회     REUPLOAD
```

10개 case 기준 예상 기록:

```text
실행 Observation      50건
Cold                  10건
Warm                  20건
Download reuse        10건
Reupload              10건
직접 Lifecycle         20건 이상
파생 Reupload event    10건
```

파일을 삭제하지 않은 후보 변경·동일 후보 재생성은 기존 캐시를 사용해야 하며 Provider 추가 호출이 없어야 한다. 삭제 후 동일 파일 재업로드는 새 upload identity로 처리되고 Semantic Profiler Provider 호출이 다시 1회 발생해야 한다.

## 실제 Shadow 수집 시 Production 안전값

수집 중에도 Canary Merge는 계속 차단한다.

```text
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_ENABLED=0
QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_KILL_SWITCH=1
QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE=BLOCKED
QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED=0
QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH=1
```

실제 Shadow 자체를 실행하려면 기존 Patch 14 Feature Control의 Shadow·Provider·Cache 플래그가 허용되어야 한다. Production 관련 플래그와는 별개다.

## 암호화 저장

MongoDB collection:

```text
query_candidate_planner_real_shadow_evidence
```

Payload는 AES-256-GCM으로 암호화되며 다음 메타데이터만 평문 인덱스로 저장한다.

- recordId
- kind
- observedAt / expiresAt
- subjectTagSha256
- requestFingerprintSha256
- uploadFingerprintSha256
- caseId / scenarioId

`expiresAt` TTL index로 자동 삭제한다. 기본 보존 기간은 7일이며 최대 30일이다.

## 실제 Observation Export

Railway 또는 동일 환경변수가 설정된 관리 터미널에서 실행한다.

```powershell
node `
  .\scripts\queryCandidatePlannerExportRealShadowEvidence.js `
  --from "2026-08-05T00:00:00.000Z" `
  --to "2026-08-06T00:00:00.000Z" `
  --output ".\real-shadow-evidence-records.json"
```

정상 출력:

```text
PASS exported real shadow evidence records=<건수> output=<경로>
```

Export 파일도 원본 데이터는 포함하지 않지만 내부 평가 자료이므로 공개 저장소에 커밋하지 않는다.

## Approved Actual Pricing

템플릿:

```text
evaluation/queryCandidatePlannerApprovedActualPricingPolicy.template.json
```

템플릿은 의도적으로 `DRAFT_NOT_APPROVED`와 0원으로 제공된다. 실제 Provider의 현재 승인 단가를 확인해 다음을 충족하도록 별도 파일을 만든다.

```text
mode = APPROVED_ACTUAL
policyId = 변경 이력 식별 가능한 값
모델별 input/output 단가 = MICROUSD_PER_MILLION_TOKENS
approvedByOperator = true
```

가격 템플릿 그대로는 Builder가 반드시 BLOCKED된다.

## Evidence Bundle 생성

Patch 13.3 실제 Readiness 결과 JSON과 승인된 실제 가격 정책을 준비한 후 실행한다.

```powershell
node `
  .\scripts\queryCandidatePlannerBuildRealShadowEvidenceBundle.js `
  --records ".\real-shadow-evidence-records.json" `
  --readiness ".\candidate-planner-live-cache-parity-readiness.json" `
  --pricing ".\queryCandidatePlannerApprovedActualPricingPolicy.json" `
  --expires-hours 24 `
  --output-dir ".\real-shadow-evidence-output"
```

PASS일 때 생성되는 주요 파일:

- `queryCandidatePlannerInternalCanaryEvidenceBundle.json`
- `queryCandidatePlannerRealShadowAccuracyReport.json`
- `queryCandidatePlannerRealShadowOperationalReport.json`
- `queryCandidatePlannerRealShadowEvaluationReport.json`
- `queryCandidatePlannerRealShadowOperationalDataset.json`
- `queryCandidatePlannerRealShadowObservationDataset.json`
- `railway-evidence-variable.txt`

Evidence Bundle은 기본 24시간 후 만료되며 최대 7일까지만 허용된다.

## Bundle 통과 조건

- 실제 Execution Observation 30건 이상
- Accuracy Dataset 모든 case 포함
- case당 실제 Observation 3건 이상
- Patch 15.0 Accuracy `EVALUATION_PASS`
- Patch 15.1 Operational `EVALUATION_PASS`
- Patch 15.2 Shadow `EVALUATION_PASS`
- Patch 13.3 Readiness 유효
- Approved Actual Pricing 사용
- Primary response unchanged rate 1.0
- Guardrail violation 0
- Privacy violation 0
- Semantic Profiler-only
- Planner escalation false

동일 case의 반복 Prediction은 다수결로 대표 Prediction을 결정하며 동률은 Prediction SHA-256 사전순으로 결정한다. 입력 순서가 바뀌어도 동일 결과가 생성된다.

## 설치 후 기본 상태

Patch 적용만으로는 다음 변화가 없다.

```text
Collector                 OFF
MongoDB Evidence 저장     없음
Internal Canary           OFF
일반 사용자               Primary
내부 사용자               Primary
Production Merge          없음
Provider 추가 호출        없음
Railway 변수 자동 변경    없음
```

## 적용 및 QA

```powershell
Get-FileHash `
  .\query_candidate_patch15_3_2_real_shadow_evidence_collector.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_3_2_real_shadow_evidence_collector.zip `
  -DestinationPath . `
  -Force
```

문법 검사:

```powershell
Get-ChildItem .\automation\queryCandidatePlannerRealShadow*.js |
  ForEach-Object { node --check $_.FullName }

node --check .\routes\automationRoutes.js
node --check .\routes\fileRoutes.js
node --check .\scripts\queryCandidatePlannerExportRealShadowEvidence.js
node --check .\scripts\queryCandidatePlannerBuildRealShadowEvidenceBundle.js
```

신규 QA:

```powershell
Get-ChildItem .\tests\queryCandidatePatch15_3_2*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    node $_.FullName
    if ($LASTEXITCODE -ne 0) { throw "Failed: $($_.Name)" }
  }
```

선행 회귀:

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
node .\tests\queryCandidatePatch15_3ManifestSmokeTest.js
```

## 활성화 금지 조건

다음 중 하나라도 해당하면 Evidence Collector만 중단하고 Canary는 계속 OFF로 유지한다.

- Registry 누락 또는 fingerprint 불일치
- 암호화 secret 누락·변경
- 실제 가격 승인 전
- 30건 미만
- case coverage 미달
- Accuracy·Operational·Shadow 중 하나라도 BLOCKED
- Timeout·Error 증가
- Warm 요청에서 Provider 재호출
- Stale cache 재사용
- 개인정보·Guardrail 위반

