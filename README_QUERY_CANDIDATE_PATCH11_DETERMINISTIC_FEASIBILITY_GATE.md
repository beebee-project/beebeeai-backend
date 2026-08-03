# Patch 11 — Deterministic Feasibility Gate

## 목적

Patch 9의 `candidate-family-resolution.json`에서 `familyDisposition=SELECTED`인 대표 후보만 대상으로 실제 자동화 시트 실행 계약을 결정론적으로 검사합니다.

이 단계는 Candidate Ranker보다 먼저 실행되며 `originalRank`, `originalScore`, `resolutionScore`를 Feasibility 판정에 사용하지 않습니다.

## 입력과 출력

입력:

- `candidate-resolution.json`
- `candidate-family-resolution.json`

출력:

- `candidate-feasibility-resolution.json`
- `candidate-feasibility-resolution-index.json`
- `candidate-feasibility-sample-audit.json`
- `candidate-feasibility-sample-audit.md`

계약 버전:

- `query_candidate_feasibility_resolution_v1`
- `query_candidate_feasibility_item_v1`
- `query_candidate_execution_plan_v1`
- `deterministic_candidate_feasibility_policy_v1_1`

## 상태

### READY

다음 조건을 모두 충족합니다.

- Patch 9 SELECTED 대표 후보
- source scope PASS 및 source root 존재
- `summarySheet` 출력 지원
- executor가 DECLARED 또는 GENERIC
- operation 계약 지원
- generic aggregate/time/ranking/cross/count 후보의 필수 operand columnId 결속
- required role PASS
- required capability PASS
- 실행 constraint PASS
- metric family 명시 실패 없음

### REVIEW

source와 summarySheet 출력은 유효하지만 현재 generic executor 계약만으로 실행 계획을 완전히 확정할 수 없습니다.

- `single_source_dashboard`
- `multi_source_schema_union`

이 후보들은 자동 탈락시키지 않고 수동 확인 사유를 execution plan에 기록합니다.

### UNSUPPORTED

명시적인 실행 차단 사유가 있습니다.

- source scope 실패 또는 source 누락
- executor UNKNOWN/FAIL
- summarySheet 출력 미지원
- generic operation 미지원
- 필수 operand columnId 누락
- required role/capability/constraint 실패
- metric family 실패

### NOT_APPLICABLE

다음 후보에는 Gate를 실행하지 않습니다.

- Patch 9 SUPPRESSED 후보
- Resolver의 STILL_DEFERRED 또는 EXCLUDED 후보

## Generic operation 계약

READY 가능:

- `count_rows`
- `category_count`
- `group_sum`
- `group_avg`
- `group_summary`
- `composition_ratio`
- `top_bottom`
- `time_sum`
- `time_avg`
- `time_count` (`period` 1개, measure 불필요)
- `cumulative_sum`
- `cross_sum`
- `cross_count`

REVIEW:

- `single_source_dashboard`
- `multi_source_schema_union`

## 실행 계획

평가된 대표 후보에는 결정론적 execution plan을 기록합니다.

- candidateId/familyId
- recipeId/templateId
- executorMode
- operation
- sourceTableIds
- outputType (`summarysheet`)
- operandBindings의 columnId
- requiredRoleBindings의 columnId
- manual confirmation 여부와 reason code

원본 행과 sample value는 저장하지 않습니다.

## 보호 조건

- production route 변경 없음
- OpenAI 호출 없음
- 추가 LLM 비용 없음
- source candidate mutation 없음
- Patch 9 family disposition mutation 없음
- Ranker 의존 없음
- 원문 행 또는 sample value 영속화 없음

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch11_deterministic_feasibility_gate.zip `
  -DestinationPath . `
  -Force
```

## 신규 테스트

```powershell
node --check .\automation\queryCandidateFeasibilityGate.js
node --check .\tests\queryCandidateFeasibilityGateCapture.js
node --check .\tests\queryCandidateFeasibilityGateSampleAudit.js

node .\tests\queryCandidateFeasibilityGateSmokeTest.js
node .\tests\queryCandidateFeasibilityGateDeclaredExecutorSmokeTest.js
node .\tests\queryCandidateFeasibilityGateGenericOperandSmokeTest.js
node .\tests\queryCandidateFeasibilityGateStructuralReviewSmokeTest.js
node .\tests\queryCandidateFeasibilityGateUnsupportedSmokeTest.js
node .\tests\queryCandidateFeasibilityGateMissingOperandSmokeTest.js
node .\tests\queryCandidateFeasibilityGateSelectedOnlySmokeTest.js
node .\tests\queryCandidateFeasibilityGateRankIndependenceSmokeTest.js
node .\tests\queryCandidateFeasibilityGatePrivacyBoundarySmokeTest.js
node .\tests\queryCandidateFeasibilityGateSchemaSmokeTest.js
node .\tests\queryCandidateFeasibilityGateBaselineSmokeTest.js
node .\tests\queryCandidateFeasibilityGateTimeCountSmokeTest.js
node .\tests\queryCandidateFeasibilityGateTimeCountAliasSmokeTest.js
node .\tests\queryCandidateFeasibilityGateTimeCountMissingPeriodSmokeTest.js
node .\tests\queryCandidateFeasibilityGateTimeCountAmbiguousPeriodSmokeTest.js
node .\tests\queryCandidatePatch11SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch11ManifestSmokeTest.js
```

## 기준선 작성과 비교

Patch 9 family 기준선이 먼저 존재해야 합니다.

```powershell
node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=write

node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=compare
```

정상 조건:

- `PASS 6/6`
- `errors=0`
- `differences=0`
- `ready + review + unsupported = selectedInput`
- 전체 상태 합계 = total

## 표본 감사

```powershell
node .\tests\queryCandidateFeasibilityGateSampleAudit.js `
  --limit=20
```

확인 항목:

- aggregate/time/ranking 후보의 operand columnId
- 선언형 executor 후보의 required role columnId
- dashboard와 multi-source schema union의 REVIEW 사유
- UNSUPPORTED 후보의 blocking reason
- summarySheet 외 출력만 있는 후보의 차단

## 예상 실제 케이스

현재 Patch 9 감사 기준으로 예상되는 방향입니다.

- 생활폐기물 하드케이스: aggregate 후보 READY, single-source dashboard와 multi-source schema union REVIEW
- 행사 신청자: 기간·그룹·순위 후보 및 `time_count` READY
- 강좌평가: 기간·그룹 및 올바른 교육 template 후보 READY
- seed 선언형 executor: READY

실제 수치는 사용자 저장소의 6개 fixture 실행 결과로 확정합니다.
