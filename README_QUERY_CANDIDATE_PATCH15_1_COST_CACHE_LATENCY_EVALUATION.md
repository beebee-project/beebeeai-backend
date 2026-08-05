# Patch 15.1 — Cost, Cache-Hit, Latency Evaluation

## 목적

Patch 15.1은 Query Candidate Planner의 운영 효율을 하나의 결정론적 평가 리포트로 고정한다.

평가 범위:

- Cold / Warm 실행 분리
- 전체 Cache Hit Rate
- Warm Cache Hit Rate
- 다운로드 후 재실행 Cache Hit
- L1 / L2 / L3 / L4 적중 분포
- Provider Call Rate
- Warm Provider Call Rate
- 삭제 후 재업로드 Provider 재호출
- 입력·출력 Token 사용량
- 실행 비용과 Cache 비용 회피량
- 월간 실행량 기준 비용 추정
- p50 / p95 / p99 Latency
- Timeout / Error Rate
- 다운로드 Cache 유지
- 삭제 Cache 무효화
- 재업로드 Identity 분리
- Stale Cache 재사용 위반

이번 패치는 평가 계층만 추가한다. API Route, Controller, 일반 사용자 UI, Promotion Gate, Merge Adapter에는 연결하지 않는다.

## 비용 정책 주의사항

패키지의 기본 가격 정책은 테스트를 위한 `SYNTHETIC_BENCHMARK_ONLY` 정책이다.

```text
Provider 청구서                    아님
현재 상용 모델 가격의 주장          아님
Production 비용 승인 근거           아님
테스트에서 계산 경로 검증           맞음
Live 평가 시 최신 가격 주입 필요     맞음
```

Evaluator의 비용 우선순위:

```text
1. provider.observedCostMicrousd
2. Token 수 × 외부 주입 Pricing Policy
3. 가격 정보가 없으면 EVALUATION_BLOCKED
```

따라서 실제 운영 평가에서는 Provider 사용량·청구 데이터 또는 당시의 승인된 가격 정책을 별도로 주입해야 한다.

## 데이터셋

```text
Dataset version   query_candidate_planner_operational_evaluation_dataset_v1
Dataset ID        beebeeai_query_candidate_cost_cache_latency_core_v1
Execution         25건
Lifecycle event   15건
Provider 실호출    0회
```

실행 구성:

```text
Cold                  5건
Warm                 10건
Download Reuse        5건
Delete 후 Reupload    5건
```

각 시나리오는 다음 수명주기를 포함한다.

```text
Cold MISS → Provider Call → Cache Write
Warm L1/L2/L3/L4 HIT → Provider Call 없음
Download → Cache RETAINED
Download 후 재실행 → L4 HIT
Delete → Cache INVALIDATED
Reupload → 새 Identity + MISS + Provider Call
```

원본 행, 실제 파일명, 이메일, 사용자 ID, tenant ID, storage key, queryTablesKey, cache secret은 데이터셋과 리포트에 포함하지 않는다.

## 지표

### Cost

- Total Cost (microusd)
- Average Cost per Execution
- Average Cost per Provider Call
- Warm Average Cost
- p95 Cost per Execution
- Cache Avoided Cost
- Cache Cost Avoidance Rate
- Monthly Projected Cost
- Observed Cost / Token Policy / No Call 분포

### Cache

- Eligible Read Count
- Hit / Miss Count
- Overall Hit Rate
- Warm Hit Rate
- Download Reuse Hit Rate
- L1 / L2 / L3 / L4 Count와 Rate

### Provider

- Call Count / Rate
- Warm Call Rate
- Reupload Call Rate
- Input / Output Token Total

### Latency

각 구간에서 Average, p50, p95, p99, Max를 계산한다.

```text
Overall
Cold
Warm
Cache Hit
Reupload
```

Percentile은 결정론적 nearest-rank 방식이다.

### Reliability와 Lifecycle

- Success / Timeout / Error Rate
- Download Retention Accuracy
- Delete Invalidation Coverage
- Reupload Identity Separation Accuracy
- Stale Cache Reuse Violation Count

## 기본 Threshold Policy

```text
Execution 표본                         >= 25
Cold 표본                              >= 5
Warm 표본                              >= 15
Lifecycle 표본                         >= 15

Overall Cache Hit Rate                 >= 0.60
Warm Cache Hit Rate                    >= 1.00
Download Reuse Cache Hit Rate          >= 1.00
Provider Call Rate                     <= 0.40
Warm Provider Call Rate                <= 0.00
Reupload Provider Call Rate            >= 1.00

Overall p95 Latency                    <= 1400 ms
Overall p99 Latency                    <= 1500 ms
Warm p95 Latency                       <= 120 ms
Cache Hit p95 Latency                  <= 120 ms

Timeout Rate                           <= 0.00
Error Rate                             <= 0.00
Average Cost                           <= 130 microusd
Provider Call Average Cost             <= 325 microusd
Warm Average Cost                      <= 0 microusd
10,000회 월간 추정 비용                <= 1,300,000 microusd
Cache Cost Avoidance Rate              >= 0.59

Delete Invalidation Coverage           >= 1.00
Download Retention Accuracy            >= 1.00
Reupload Identity Separation Accuracy  >= 1.00
Stale Cache Reuse Violation             = 0
```

Threshold 통과는 Production 승인이 아니다.

```text
evaluationOnly                   true
promotionAuthorized              false
productionCandidateMergeApplied  false
productionReadyAssignment        false
providerCallsExecutedByEvaluator 0
```

## 적용

백엔드 저장소 루트에서 실행한다.

```powershell
Get-FileHash `
  .\query_candidate_patch15_1_cost_cache_latency_evaluation.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_1_cost_cache_latency_evaluation.zip `
  -DestinationPath . `
  -Force
```

이 ZIP은 저장소 루트 구조로 구성되어 추가 최상위 폴더를 만들지 않는다.

설치 확인:

```powershell
Test-Path `
  .\automation\queryCandidatePlannerCostCacheLatencyEvaluator.js

Test-Path `
  .\evaluation\queryCandidatePlannerOperationalEvaluationDataset.v1.json
```

## 문법 검사

```powershell
node --check `
  .\automation\queryCandidatePlannerCostCacheLatencyEvaluator.js

Get-ChildItem .\tests\queryCandidatePatch15_1*.js |
  ForEach-Object {
    node --check $_.FullName
  }
```

## Patch 15.1 QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch15_1*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    node $_.FullName
  }
```

정상적으로 31개 PASS가 출력되어야 한다.

주요 검증:

```text
Dataset / Pricing / Threshold 계약
Dataset / Report 개인정보 경계
Cold / Warm 구간 분리
Cache Hit와 L1~L4 분포
Provider Call Rate
Token 기반 비용 계산
Observed Cost 우선 적용
Cache Cost Avoidance
월간 비용 Projection
p50 / p95 / p99 Latency
Timeout / Error Rate
Download Retention
Delete Invalidation
Reupload Identity 분리
Stale Cache fail-closed
가격 누락 fail-closed
표본 부족 fail-closed
입력 불변
결정론적 Report
Route 격리
Schema
Source Integrity
Manifest
```

## 선행 패치 회귀 검사

Patch 15.1은 adds-only이므로 기존 파일을 변경하지 않는다.

```powershell
node .\tests\queryCandidatePatch14_4_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5FullQualityGateSmokeTest.js
node .\tests\queryCandidatePatch15_0ManifestSmokeTest.js
```

## 실제 운영 데이터 연결 시점

이번 단계에서는 합성 결정론적 데이터로 Metric과 Threshold 계산 계약을 검증한다.

실제 Shadow observation과 Provider 사용량을 평가 데이터로 변환하는 연결은 Patch 15.2에서 수행한다. 실제 가격·환율·청구 정책은 평가 실행 시점의 승인된 값을 주입해야 한다.
