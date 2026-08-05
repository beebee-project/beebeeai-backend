# Patch 15.0 — Accuracy Metric & Evaluation Dataset

## 목적

Patch 15.0은 Query Candidate Planner의 정확도를 동일한 기준으로 반복 측정하기 위한 평가 전용 계층을 추가한다.

이번 패치는 다음을 고정한다.

- 정답 후보(required), 허용 후보(acceptable), 금지 후보(forbidden) 라벨 계약
- Top-1, Top-k, 후보 Precision·Recall, 순위 일치도
- 업무 도메인·데이터셋 의도 정확도
- fallback 판단 정확도
- unsupported 데이터 거부 정확도
- review 필요 판단 정확도
- 잘못된 Production 후보 승격률(False Promotion Rate)
- 평가 통과 기준과 fail-closed 보고서
- 원본 행·파일명·사용자 식별자를 포함하지 않는 labels-only 평가 데이터셋

기존 API Route, Controller, 일반 사용자 UI, Promotion Gate, Merge Adapter에는 연결하지 않는다.

## 고정 안전 상태

- 평가 전용(`evaluationOnly=true`)
- Production 승인 생성 금지(`promotionAuthorized=false`)
- Promotion Gate 연결 없음
- Production 후보 병합 없음
- READY 승격 없음
- Production Route 변경 없음
- 실제 Provider 호출 0회
- Railway 환경변수 변경 없음
- 원본 행·파일명·사용자 식별자 저장 없음
- 잘못된 데이터셋·예측·정책은 `EVALUATION_BLOCKED`

## 평가 데이터셋

파일:

```text
evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json
```

Core v1은 10개 사례를 포함한다.

```text
재정·예산 실행
행사 신청·출석
조건형 출석
매출 거래
강의 평가
비정형 unsupported 거부
업무 의미가 모호한 generic fallback
재고 이동
프로젝트 업무 추적
지출·경비 검토
```

데이터셋에는 fixture 식별자와 평가 라벨만 포함한다. 원본 셀 값과 저장 객체 키는 포함하지 않는다.

## 지표 정의

```text
Candidate Precision
  출력 후보 중 required 또는 acceptable 후보의 비율

Candidate Recall
  required 후보 중 실제로 출력된 후보의 비율

Top-1 Accuracy
  첫 번째 후보가 preferredTop1CandidateIds에 포함되는 비율

Top-k Recall
  상위 k개 안에 포함된 required 후보의 비율

Ranking Agreement
  required 후보 쌍의 이상적 순서와 실제 순서가 일치하는 비율

Domain Accuracy / Intent Accuracy
  expected 또는 명시된 acceptable 값과 일치하는 비율

Fallback Accuracy
  fallback 적용 여부와 허용 사유가 라벨과 일치하는 비율

Unsupported Rejection Accuracy
  unsupported 사례에서 후보를 승격하지 않고 거부했는지 평가

Review Decision Accuracy
  내부 검토 필요 여부가 라벨과 일치하는 비율

False Promotion Rate
  출력 후보 중 forbidden 후보의 비율
```

## 기준 정책

파일:

```text
evaluation/queryCandidatePlannerAccuracyThresholdPolicy.v1.json
```

Core v1 기준:

```text
최소 사례 수                         10
모든 사례 예측                       필수
Overall Score                       >= 0.85
Candidate Precision                 >= 0.85
Candidate Recall                    >= 0.85
Top-1 Accuracy                      >= 0.80
Top-k Recall                        >= 0.85
Ranking Agreement                   >= 0.75
Domain Accuracy                     >= 0.90
Intent Accuracy                     >= 0.90
Fallback Accuracy                   >= 0.90
Unsupported Rejection Accuracy      >= 1.00
Review Decision Accuracy            >= 0.90
False Promotion Rate                <= 0.02
```

이 기준 통과는 평가 결과일 뿐 Production Promotion 허가가 아니다.

## 적용

백엔드 루트에서 실행한다.

```powershell
Get-FileHash `
  .\query_candidate_patch15_0_accuracy_metric_evaluation_dataset.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_0_accuracy_metric_evaluation_dataset.zip `
  -DestinationPath . `
  -Force
```

ZIP은 저장소 루트 구조로 구성되어 있어 추가 최상위 폴더를 만들지 않는다.

## 문법 검사

```powershell
node --check `
  .\automation\queryCandidatePlannerAccuracyEvaluator.js

Get-ChildItem .\tests\queryCandidatePatch15_0*.js |
  ForEach-Object {
    node --check $_.FullName
  }
```

## Patch 15.0 QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch15_0*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    node $_.FullName
  }
```

정상적으로 24개 PASS가 출력되어야 한다.

핵심 출력:

```text
PASS query candidate patch15.0 accuracy dataset contract smoke
PASS query candidate patch15.0 perfect evaluation smoke
PASS query candidate patch15.0 false promotion smoke
PASS query candidate patch15.0 unsupported rejection smoke
PASS query candidate patch15.0 missing prediction fail-closed smoke
PASS query candidate patch15.0 manifest smoke
```

## 선행 패치 호환성

```powershell
node .\tests\queryCandidatePatch14_4_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5FullQualityGateSmokeTest.js
```

Patch 15.0은 기존 파일을 교체하지 않는 adds-only 패치이므로 선행 Manifest를 변경하지 않는다.

## 다음 단계

Patch 15.1에서 실제 실행의 비용·Cache-Hit·Latency를 측정한다. Patch 15.2에서는 Shadow 관찰 결과를 이번 데이터셋과 지표 계약으로 평가한다. Promotion Gate는 Patch 15.3 Internal Allowlist Canary까지 계속 비활성 상태로 유지한다.
