# Patch 15.2 — Shadow Accuracy Evaluation

## 목적

Patch 15.2는 Patch 14.1의 Shadow 실행 결과를 Patch 15.0 정확도 지표로 평가할 수 있도록 독립적인 평가 연결 계층을 추가한다.

핵심 흐름:

```text
Sanitized Shadow Result
→ Shadow Accuracy Observation
→ Accuracy Prediction 변환
→ Patch 15.0 Accuracy Evaluator
→ Shadow Health Threshold 평가
→ EVALUATION_PASS 또는 EVALUATION_BLOCKED
```

이번 패치는 평가 계층만 추가한다. API Route, Controller, 일반 사용자 UI, Internal Preview, Promotion Gate, Merge Adapter에는 연결하지 않는다.

## 중요한 데이터 구분

패키지에 포함된 Shadow Observation Dataset은 평가 계산 경로를 검증하기 위한 합성 기준선이다.

```text
실제 Production Shadow Traffic        아님
실제 Provider 호출 결과               아님
Canary 승인 근거                       아님
평가 파이프라인 계약 검증              맞음
실제 Shadow 결과 주입 전 기준선         맞음
```

Patch 14.3 Internal Preview는 개인정보 보호를 위해 후보 ID 원문을 저장하지 않는다. 정확도 계산에는 후보 ID와 순위가 필요하므로 Patch 15.2는 별도의 평가 전용 `Shadow Accuracy Observation` 계약을 정의한다.

이 계약은 다음 값만 보존한다.

- 평가 caseId
- 요청 fingerprint SHA-256
- Shadow 상태
- 후보 ID·순위·수락 상태
- 업무 도메인과 데이터셋 의도
- fallback·unsupported·review 판정
- Comparator 요약
- Primary 응답 불변 여부
- 안전 Guardrail과 Privacy 선언

원본 행, 실제 파일명, 이메일, 사용자 ID, tenant ID, storage key, queryTablesKey, 원본 API 응답, 원본 Provider 응답은 저장하지 않는다.

## 신규 파일

```text
automation/queryCandidatePlannerShadowAccuracyEvaluator.js
automation/queryCandidatePlannerShadowAccuracyEvaluator.schema.json
automation/queryCandidatePlannerShadowAccuracyObservationDataset.schema.json
automation/queryCandidatePlannerShadowAccuracyThresholdPolicy.schema.json

evaluation/queryCandidatePlannerShadowAccuracyObservationDataset.v1.json
evaluation/queryCandidatePlannerShadowAccuracyThresholdPolicy.v1.json
```

## 평가 데이터셋

```text
Dataset version  query_candidate_planner_shadow_accuracy_observation_dataset_v1
Dataset ID       beebeeai_query_candidate_shadow_accuracy_core_v1
Observation      10건
Accuracy case    10건
Provider 실호출   0회
```

Patch 15.0의 다음 10개 정답 사례와 일대일로 연결된다.

```text
예산 집행
행사 신청·출석
조건형 출석
매출 거래
강의 평가
비정형 unsupported
모호한 데이터 fallback
재고 이동
프로젝트 업무 추적
지출·경비 검토
```

## Shadow Health 지표

- Observation Count
- Completed Rate
- Blocked Rate
- Failed-Safe Rate
- Timeout-Safe Rate
- Primary Response Unchanged Rate
- Comparator Coverage
- Prediction Capture Coverage
- Guardrail Violation Count
- Privacy Violation Count
- Accuracy Case Coverage
- Patch 15.0 Accuracy Evaluation Decision
- Comparator Verdict Distribution
- 평균 Shadow Latency

## 정확도 지표

Shadow Prediction을 Patch 15.0 Evaluator에 전달하므로 동일한 정확도 계약을 사용한다.

- Candidate Precision
- Candidate Recall
- Top-1 Accuracy
- Top-k Recall
- Ranking Agreement
- Domain Accuracy
- Intent Accuracy
- Fallback Accuracy
- Unsupported Rejection Accuracy
- Review Decision Accuracy
- False Promotion Rate
- Overall Weighted Score

## 기본 Threshold Policy

```text
Observation Count                    >= 10
Completed Rate                       >= 1.00
Blocked Rate                         <= 0.00
Failed-Safe Rate                     <= 0.00
Timeout-Safe Rate                    <= 0.00
Primary Response Unchanged Rate      >= 1.00
Comparator Coverage                  >= 1.00
Prediction Capture Coverage          >= 1.00
Guardrail Violation Count            <= 0
Privacy Violation Count              <= 0
Accuracy Case Coverage               = 전체
Patch 15.0 Accuracy Decision         = EVALUATION_PASS
```

Threshold 통과는 Production 승인이 아니다.

```text
evaluationOnly                   true
promotionAuthorized              false
productionGateWired              false
productionCandidateMergeApplied  false
productionReadyAssignment        false
providerCallsExecutedByEvaluator 0
```

## Capture Adapter

`buildShadowAccuracyObservation()`은 평가용 caseId, API Shadow Observation, Shadow Resolution을 입력받아 개인정보가 제거된 평가 Observation을 생성한다.

지원 후보 경로:

```text
plannerResolution.items
plannerResolution.candidates
candidateResolution.items
rankingResolution.items
items
candidates
topCandidates
```

입력에 원본 행, 파일명, 사용자 식별자, tenant ID, queryTablesKey 등의 금지 필드가 포함되면 즉시 예외를 발생시켜 fail-closed 처리한다.

현재 서비스 경로에는 이 Adapter가 연결되지 않는다. 실제 Shadow Accuracy 데이터 수집 연결은 평가 승인 후 별도 패치에서 수행해야 한다.

## 적용

백엔드 저장소 루트에서 실행한다.

```powershell
Get-FileHash `
  .\query_candidate_patch15_2_shadow_accuracy_evaluation.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_2_shadow_accuracy_evaluation.zip `
  -DestinationPath . `
  -Force
```

이 ZIP은 저장소 루트 구조로 구성되어 추가 최상위 폴더를 만들지 않는다.

설치 확인:

```powershell
Test-Path `
  .\automation\queryCandidatePlannerShadowAccuracyEvaluator.js

Test-Path `
  .\evaluation\queryCandidatePlannerShadowAccuracyObservationDataset.v1.json
```

정상 결과:

```text
True
True
```

## 문법 검사

```powershell
node --check `
  .\automation\queryCandidatePlannerShadowAccuracyEvaluator.js

Get-ChildItem .\tests\queryCandidatePatch15_2*.js |
  ForEach-Object {
    node --check $_.FullName
  }
```

## Patch 15.2 QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch15_2*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    node $_.FullName
  }
```

정상적으로 36개 PASS가 출력되어야 한다.

주요 검증:

```text
Capture Adapter
Capture Privacy fail-closed
Observation Dataset 계약·Coverage·Privacy
Threshold Policy
완전 정답 Shadow 평가
Prediction 변환
Patch 15.0 Evaluator 연동
Candidate·Domain·Intent 정확도
Fallback·Unsupported·Review 정확도
Comparator Coverage·Verdict 분포
Primary 응답 불변
Blocked·Failed·Timeout fail-closed
Guardrail·Privacy 위반 fail-closed
누락·중복·Malformed Observation fail-closed
Accuracy Regression fail-closed
입력 불변
결정론적 Report
Report Privacy
Route 격리
Schema
Source Integrity
Manifest
```

## 선행 패치 회귀 검사

Patch 15.2는 adds-only이며 기존 파일을 변경하지 않는다.

```powershell
node .\tests\queryCandidatePatch14_4_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5FullQualityGateSmokeTest.js
node .\tests\queryCandidatePatch15_0ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_1ManifestSmokeTest.js
```

선행 파일에 이미 Manifest drift가 존재하는 경우 Patch 15.2 적용 전에 정식 해시로 복원해야 한다. Patch 15.2 ZIP은 선행 파일을 포함하거나 덮어쓰지 않는다.

## 현재 상태

```text
Shadow Accuracy Evaluator          구현 완료
합성 Shadow 기준선                 포함
실제 Shadow Traffic 연결           없음
Internal Preview 연결              없음
API·Controller 연결                없음
Promotion Gate 연결                없음
Production Merge                   없음
Railway 환경변수 변경              없음
Provider 실호출                    0회
```

실제 Shadow Accuracy 평가는 향후 실제 관찰 데이터를 같은 계약으로 주입해 다시 실행해야 한다. Patch 15.2의 합성 기준선 PASS만으로 Patch 15.3 Internal Allowlist Canary를 활성화해서는 안 된다.
