# BeeBee AI Query Candidate Patch 8.2

## Recipe Operand Binding Guard

패치 8.2는 패치 8.1의 domain·identity guardrail 다음 단계입니다.

기존 Resolver는 `group_sum`, `time_avg` 같은 operation과 범용 역할이 맞으면
후보를 `RESOLVED`할 수 있었지만, 후보 식별자에 적힌 실제 그룹 열·기간 열·측정값
열이 정확히 연결됐는지는 별도로 확인하지 않았습니다.

예:

```text
후보: 신청자별 만족도 평균
범용 group 매칭: 소속
```

이 상태에서는 operation은 가능해도 후보가 약속한 축과 실제 생성 축이 다를 수
있습니다.

패치 8.2는 다음 결속을 추가합니다.

```text
candidateId 또는 동적 recipeId
→ operation과 operand 추출
→ 선택된 source table의 실제 열 검색
→ operandBinding PASS / UNKNOWN / FAIL
```

새 policy:

```text
deterministic_candidate_resolution_policy_v1_2
```

Production route, OpenAI 호출, 후보 `READY` 상태는 변경하지 않습니다.

---

## 1. 검사 대상 operation

```text
time_sum          period + measure
time_avg          period + measure
cumulative_sum    period + measure
group_sum         group + measure
group_avg         group + measure
group_summary     group + measure
composition_ratio group + measure
top_bottom        group + measure
category_count    group
cross_sum         dimension + dimension + measure
cross_count       dimension + dimension
```

`single_source_dashboard`, `multi_source_schema_union`처럼 특정 열 operand를
식별자에 선언하지 않는 구조형 후보는 `NOT_APPLICABLE`로 유지합니다.

---

## 2. 결속 규칙

### 정확한 열 우선

```text
신청자 → 신청자
소속 → 소속
강사명 → 강사명
평가점수 → 평가점수
만족도 → 만족도
```

후보명과 다른 범용 dimension 열은 대신 사용할 수 없습니다.

### 안전한 측정값 표기 호환

다음과 같이 같은 의미의 접미어 차이는 허용합니다.

```text
매출액 ↔ 매출금액
```

다른 업무 의미의 열로 확장하지 않습니다.

### 기간 operand

`연월`, `기간`, `일자`처럼 일반 기간 표현인데 정확히 같은 헤더가 없으면,
선택된 source scope에 period/date 열이 정확히 하나일 때만 결속합니다.

```text
연월 → 평가일
matchMode: UNIQUE_SEMANTIC_PERIOD
```

기간 열이 여러 개면 임의로 선택하지 않고 `UNKNOWN`으로 남깁니다.

---

## 3. 판정

```text
operandBinding PASS 또는 NOT_APPLICABLE
+ 기존 source/domain/executor/role 검사 PASS
→ RESOLVED 가능
```

```text
INFERRED + operand 없음/불일치/모호
→ STILL_DEFERRED
```

```text
BOUND/PARTIAL + 필수 operand 명백히 없음
→ EXCLUDED 가능
```

주요 reason code:

```text
RECIPE_OPERANDS_BOUND
RECIPE_OPERAND_BINDING_NOT_CONFIRMED
RECIPE_OPERAND_BINDING_AMBIGUOUS
```

Resolver 결과의 `matchedColumnIds`와 `evidence`에는 operand로 직접 확인된 열도
추가됩니다. 따라서 감사 보고서에서 범용 역할 열이 아니라 실제 그룹·측정값 열을
확인할 수 있습니다.

---

## 4. 적용

백엔드 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch8_2_recipe_operand_binding_guard.zip `
  -DestinationPath . `
  -Force
```

패치 8, 8.1의 manifest는 Resolver와 schema의 새 해시를 반영한 버전으로 함께
갱신됩니다. 패치 8·8.1·8.2 메타파일을 삭제하지 마세요.

---

## 5. 파일 확인

```powershell
Test-Path .\automation\queryCandidateResolver.js
Test-Path .\tests\queryCandidateResolverOperandBindingSmokeTest.js
Test-Path .\tests\queryCandidateResolverSampleAudit.js
Test-Path .\PATCH_MANIFEST_PATCH8_2.json
Test-Path .\PATCH_VALIDATION_PATCH8_2.json
Test-Path .\README_QUERY_CANDIDATE_PATCH8_2_RECIPE_OPERAND_BINDING_GUARD.md
```

모두 `True`여야 합니다.

---

## 6. 신규 스모크

```powershell
node --check .\automation\queryCandidateResolver.js
node --check .\tests\queryCandidateResolverSampleAudit.js

node .\tests\queryCandidateResolverOperandBindingSmokeTest.js
node .\tests\queryCandidateResolverOperandMismatchSmokeTest.js
node .\tests\queryCandidateResolverPeriodOperandSmokeTest.js
node .\tests\queryCandidateResolverDynamicRecipeOperandSmokeTest.js
node .\tests\queryCandidatePatch8_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8_2ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate resolver operand binding smoke
PASS query candidate resolver operand mismatch smoke
PASS query candidate resolver period operand smoke
PASS query candidate resolver dynamic recipe operand smoke
PASS query candidate patch8.2 source integrity smoke
PASS query candidate patch8.2 manifest smoke
```

---

## 7. 기존 패치 회귀

```powershell
node .\tests\queryCandidateResolverSmokeTest.js
node .\tests\queryCandidateResolverInferredSemanticSmokeTest.js
node .\tests\queryCandidateResolverSourceScopeSmokeTest.js
node .\tests\queryCandidateResolverConservativeExclusionSmokeTest.js
node .\tests\queryCandidateResolverIntegritySmokeTest.js
node .\tests\queryCandidateResolverPrivacyBoundarySmokeTest.js
node .\tests\queryCandidateResolverSchemaSmokeTest.js
node .\tests\queryCandidateResolverDomainEvidenceSmokeTest.js
node .\tests\queryCandidateResolverIdentityGuardrailSmokeTest.js

node .\tests\queryCandidatePatch8SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8ManifestSmokeTest.js
node .\tests\queryCandidatePatch8_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8_1ManifestSmokeTest.js
```

모두 PASS해야 합니다.

---

## 8. 기준선 재작성

policy와 candidate check 구조가 변경되므로 기준선 재작성은 필수입니다.

```powershell
node .\tests\queryCandidateResolverCapture.js `
  --mode=write

node .\tests\queryCandidateResolverCapture.js `
  --mode=compare
```

정상 기준:

```text
[query-candidate-resolver] cases: 6
[query-candidate-resolver] PASS 6/6
errors=0
differences=0
```

`RESOLVED` 수가 줄어들 수 있습니다. 이는 operand가 확인되지 않은 후보를
보수적으로 `STILL_DEFERRED`로 내린 의도된 변화입니다.

---

## 9. 재감사

```powershell
node .\tests\queryCandidateResolverSampleAudit.js `
  --resolved-limit=10 `
  --excluded-limit=8
```

감사 보고서에 새 섹션이 추가됩니다.

```text
Recipe operand 결속
kind / expected / status / match mode / matched
```

대표 확인 항목:

```text
신청자별 만족도 → 신청자 + 만족도 열 결속
소속별 만족도   → 소속 + 만족도 열 결속
강좌·수업 평가  → 강사명 + 평가점수 열 결속
연월별 평가점수 → period 열 + 평가점수 열 결속
operand 없는 누적합계 후보 → STILL_DEFERRED
```

재감사 결과가 정상인 것을 확인한 뒤 패치 9 Candidate Family·중복 제거로
진행합니다.
