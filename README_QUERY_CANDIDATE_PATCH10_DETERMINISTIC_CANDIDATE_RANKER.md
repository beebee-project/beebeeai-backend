# Query Candidate Patch 10 — Deterministic Candidate Ranker

> 적용 순서: Patch 9 → Patch 11 → Patch 11.1 → **Patch 10 Ranker**

이 패치는 Patch 11.1의 `candidate-feasibility-resolution.json`에서 `feasibilityStatus = READY`인 후보만 결정론적으로 순위화합니다. 번호는 Patch 10이지만 Feasibility Gate를 먼저 안정화한 뒤 적용하도록 순서를 의도적으로 뒤로 이동했습니다.

## 출력

```text
candidate-resolution.json
candidate-family-resolution.json
candidate-feasibility-resolution.json
↓
candidate-ranking-resolution.json
```

계약 버전:

```text
query_candidate_ranking_resolution_v1
query_candidate_ranking_item_v1
deterministic_candidate_ranking_policy_v1_1
```

## 입력 경계

순위 대상:

```text
familyDisposition = SELECTED
feasibilityStatus = READY
```

순위 제외:

```text
REVIEW
UNSUPPORTED
NOT_APPLICABLE
SUPPRESSED
STILL_DEFERRED
EXCLUDED
```

제외 후보도 출력에서 삭제되지 않고 `rankingDisposition = NOT_APPLICABLE`로 보존됩니다.

## 점수 구성

점수는 0~100 범위이며 고정된 세부 항목으로 계산합니다.

- READY 기본점수
- 이전 Retriever 결과
- BOUND/PARTIAL/INFERRED binding 강도
- DECLARED/GENERIC executor 확실성
- template ID가 있는 명명형 후보의 정체성
- operation의 업무 활용도
- 척도형·가산형 metric의 집계 방식 적합성
- 고카디널리티 category_count의 단조로운 결과 가능성
- operand 또는 required role 실행계약
- domain alignment
- metric family alignment
- source table 단순성
- Resolver 점수의 제한된 보조 신호
- 기존 후보 점수의 제한된 보조 신호

`originalScore`와 `resolutionScore`는 후보를 통과·탈락시키는 Gate가 아니며 각각 최대 3점·5점의 보조 신호로만 사용합니다.

Patch 10.1부터 `semantic-profile.json`의 column `uniqueRatio`와 `rowCount`를 선택적인 순위 신호로 사용합니다. 프로파일이 없더라도 Ranker는 동작하며, 이 통계는 후보를 탈락시키는 Gate가 아닙니다.

## 다양성 재정렬

기본점수가 비슷한 후보가 한 종류로 몰리지 않도록 다음 반복에 감점을 적용합니다.

```text
같은 operation 반복
같은 group/period axis 반복
같은 measure 반복
```

감점은 순위만 조정하며 후보를 삭제하지 않습니다. 모든 READY 후보는 반드시 하나의 연속된 rank를 갖습니다.

## Tier

```text
1~3위   PRIMARY
4~8위   SECONDARY
9위~    ADDITIONAL
```

`recommendedCandidateIds`에는 상위 8개가 순서대로 기록됩니다. READY 후보가 8개 미만이면 전체 READY 후보가 기록됩니다.

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch10_deterministic_candidate_ranker.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check .\automation\queryCandidateRanker.js
node --check .\tests\queryCandidateRankerCapture.js
node --check .\tests\queryCandidateRankerSampleAudit.js
```

## 신규 테스트

```powershell
node .\tests\queryCandidateRankerSmokeTest.js
node .\tests\queryCandidateRankerReadyOnlySmokeTest.js
node .\tests\queryCandidateRankerEvidencePrioritySmokeTest.js
node .\tests\queryCandidateRankerDiversitySmokeTest.js
node .\tests\queryCandidateRankerTierSmokeTest.js
node .\tests\queryCandidateRankerDeterminismSmokeTest.js
node .\tests\queryCandidateRankerPrivacyBoundarySmokeTest.js
node .\tests\queryCandidateRankerSchemaSmokeTest.js
node .\tests\queryCandidateRankerMetricAggregationAffinitySmokeTest.js
node .\tests\queryCandidateRankerAdditiveAggregationAffinitySmokeTest.js
node .\tests\queryCandidateRankerDegenerateGroupPenaltySmokeTest.js
node .\tests\queryCandidatePatch10SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch10ManifestSmokeTest.js
```

## 입력 기준선 확인

```powershell
node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=compare
```

## Ranking 기준선 작성

```powershell
node .\tests\queryCandidateRankerCapture.js `
  --mode=write

node .\tests\queryCandidateRankerCapture.js `
  --mode=compare

node .\tests\queryCandidateRankerBaselineSmokeTest.js
```

정상 조건:

```text
[query-candidate-ranker] cases: 6
[query-candidate-ranker] PASS 6/6
errors=0
warnings=0
differences=0
```

각 케이스의 불변식:

```text
ranked = readyInput
ranked + notApplicable = total
rank는 1부터 연속
recommendedCandidateIds = rankedCandidateIds의 앞 8개
```

## 표본 감사

```powershell
node .\tests\queryCandidateRankerSampleAudit.js `
  --limit=20
```

생성 파일:

```text
tests\fixtures\query-candidate-baseline\candidate-ranking-sample-audit.json
tests\fixtures\query-candidate-baseline\candidate-ranking-sample-audit.md
```

감사 항목:

```text
rank
tier
candidateId
rankingScore
baseScore
diversityPenalty
metricAggregationAffinity
degenerateGroupPenalty
operation
operationCategory
sourceTableIds
```

## 변경하지 않는 항목

- Production route
- Resolver 결과
- Family disposition
- Feasibility 상태
- 원본 candidate score/rank
- OpenAI 호출
- LLM 비용
- 원본 행 또는 sample value 저장
