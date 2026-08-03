# Query Candidate Patch 10.1 — Metric Aggregation Affinity + Degenerate Group Penalty

> 적용 순서: Patch 9 → Patch 11 → Patch 11.1 → Patch 10 → **Patch 10.1**

이 패치는 Deterministic Candidate Ranker의 실행 가능성·Tier·추천 개수 계약은 유지하면서, 실제 업무 효용에 맞게 점수 구성을 보정합니다.

## 정책 버전

```text
deterministic_candidate_ranking_policy_v1_1
```

## 1. Metric Aggregation Affinity

척도형 지표:

```text
만족도
평가점수
추천점수
평점
score / rating / satisfaction / NPS
```

우선순위 보정:

```text
group_avg / time_avg       +1.5
top_bottom                 +0.5
group_sum / time_sum       -3
cumulative_sum             -4
```

가산형 지표:

```text
금액
매출
예산
비용
지출
수량
건수
currency / amount / revenue / sales / budget
```

우선순위 보정:

```text
group_sum / time_sum       +1
top_bottom                 +1
cumulative_sum             +0.5
group_avg / time_avg       -1.5
```

이 보정은 `READY` 여부를 바꾸지 않고 Ranker 점수에만 반영됩니다.

## 2. Degenerate Group Penalty

`category_count` 후보의 group column이 행마다 거의 고유하면 각 그룹의 결과가 대부분 `1건`으로 같아질 수 있습니다.

적용 조건:

```text
operation = category_count
rowCount >= 8
semantic-profile column stats.uniqueRatio 사용
```

감점:

```text
uniqueRatio >= 0.98   -12
uniqueRatio >= 0.90   -10
uniqueRatio >= 0.80    -7
uniqueRatio >= 0.65    -4
```

통계가 없으면 감점하지 않습니다. 통계는 후보 탈락 Gate가 아니며 순위 신호로만 사용합니다.

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch10_1_metric_aggregation_affinity_degenerate_group_penalty.zip `
  -DestinationPath . `
  -Force
```

## 신규 테스트

```powershell
node --check .\automation\queryCandidateRanker.js

node .\tests\queryCandidateRankerMetricAggregationAffinitySmokeTest.js
node .\tests\queryCandidateRankerAdditiveAggregationAffinitySmokeTest.js
node .\tests\queryCandidateRankerDegenerateGroupPenaltySmokeTest.js

node .\tests\queryCandidatePatch10_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch10_1ManifestSmokeTest.js
```

## 기존 Ranker 회귀

```powershell
node .\tests\queryCandidateRankerSmokeTest.js
node .\tests\queryCandidateRankerReadyOnlySmokeTest.js
node .\tests\queryCandidateRankerEvidencePrioritySmokeTest.js
node .\tests\queryCandidateRankerDiversitySmokeTest.js
node .\tests\queryCandidateRankerTierSmokeTest.js
node .\tests\queryCandidateRankerDeterminismSmokeTest.js
node .\tests\queryCandidateRankerPrivacyBoundarySmokeTest.js
node .\tests\queryCandidateRankerSchemaSmokeTest.js
node .\tests\queryCandidatePatch10SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch10ManifestSmokeTest.js
```

## 기준선 재작성

정책 버전과 순위가 바뀌므로 Ranking 기준선을 다시 작성합니다.

```powershell
node .\tests\queryCandidateRankerCapture.js `
  --mode=write

node .\tests\queryCandidateRankerCapture.js `
  --mode=compare

node .\tests\queryCandidateRankerBaselineSmokeTest.js
```

정책 버전 확인:

```powershell
(Get-Content `
  .\tests\fixtures\query-candidate-baseline\candidate-ranking-resolution-index.json `
  -Raw | ConvertFrom-Json).policyVersion
```

기대값:

```text
deterministic_candidate_ranking_policy_v1_1
```

## 표본 감사

```powershell
node .\tests\queryCandidateRankerSampleAudit.js `
  --limit=20
```

감사표의 추가 열:

```text
affinity
  metric aggregation affinity 점수

degenerate
  고카디널리티 category_count 감점
```

확인 방향:

```text
생활폐기물
- 데이터t1_categorycount_구분1이 1위에서 내려감
- 예산 합계·순위 후보가 먼저 배치됨

행사 신청자
- 신청일자별 건수 개요는 상위 유지
- 만족도 평균이 동일 축의 만족도 합계보다 앞섬

강좌평가
- course_evaluation_report 1위 유지
- 평가점수·추천점수 평균이 합계보다 우선
```

## 변경하지 않는 항목

- Feasibility 상태
- Family disposition
- Resolver 결과
- 모든 READY 후보의 ranking coverage
- PRIMARY 3개·추천 8개 계약
- Production route
- OpenAI 호출과 LLM 비용
- 원본 행·sample value 저장
