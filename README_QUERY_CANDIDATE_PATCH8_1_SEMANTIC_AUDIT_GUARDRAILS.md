# Patch 8.1 — Candidate Resolver Semantic Audit Guardrails

## 목적

패치 8 표본 감사에서 확인된 두 종류의 오판을 수정합니다.

1. 동적 recipe ID의 열 이름이 후보의 업무영역으로 섞여 잘못 `RESOLVED`되는 문제
2. 하나의 파일에 행사 신청과 만족도처럼 복수 의미가 있는데 primary domain 하나만 보고 유효 후보를 `EXCLUDED`하는 문제

이 패치는 패치 9를 진행하기 전에 적용해야 합니다.

## 변경 정책

### 1. Domain 근거의 강도 분리

후보의 강한 업무영역 신호는 다음에서만 읽습니다.

```text
candidateId
templateId
manifest의 matchedTemplateIds
```

동적으로 생성된 `recipeId`, `bindingKey`, source column 이름은 INFERRED 후보의 강한 domain 근거로 사용하지 않습니다.

예:

```text
candidateId: monthly_expense_report
recipeId: sheet1t1_timesum_신청일자_만족도
```

결과:

```text
기대 domain = FINANCE_BUDGET
EVENT_ATTENDANCE·SURVEY_FEEDBACK으로 확장하지 않음
```

### 2. 토큰 경계 보존

업무영역 토큰을 찾을 때 `_`, 공백 등으로 구분된 segment 단위로 검사합니다.

따라서 다음 문자열에서 `명단`을 잘못 찾지 않습니다.

```text
지표명_단위
```

### 3. Secondary semantic domain evidence

`classification.secondaryDomains`에 없더라도 실제 열 의미가 강하면 보조 domain 근거로 인정합니다.

예:

```text
primaryDomain = EVENT_ATTENDANCE
열 = 만족도
```

이 경우 `SURVEY_FEEDBACK` 후보는 자동 충돌로 제외하지 않습니다.

### 4. INFERRED 명명형 template 보호

다음 조건의 후보는 generic role만 맞는다는 이유로 `RESOLVED`하지 않습니다.

```text
bindingStatus = INFERRED
templateId 존재
domain PASS 근거 없음
```

결과는 `STILL_DEFERRED`이며 다음 사유를 기록합니다.

```text
INFERRED_TEMPLATE_IDENTITY_NOT_CONFIRMED
```

반면 다음 구조형 recipe는 templateId가 없고 기계 조건을 충족하면 domain 비적용 상태에서도 의미 resolution을 허용합니다.

```text
single_source_dashboard
multi_source_schema_union
time_sum / time_avg
group_sum / group_avg
cumulative_sum / top_bottom
cross_sum / cross_count
```

## 감사 결과에 대한 기대 변화

### 생활폐기물 하드케이스

다음 명명형 후보는 `RESOLVED`에서 `STILL_DEFERRED`로 이동해야 합니다.

```text
regional_performance_report
asset_lifecycle_report
asset_equipment_management
```

다음 구조 후보는 잘못된 EVENT domain conflict가 사라져야 합니다.

```text
cross_sum
cross_count
```

필수 role/capability가 부족하면 `STILL_DEFERRED`가 정상입니다.

### 행사 신청자

```text
monthly_expense_report
→ FINANCE_BUDGET 강한 신호 유지
→ EVENT_ATTENDANCE 파일에서는 EXCLUDED
```

```text
소속별·참가상태별 만족도 합계/평균
→ 만족도 열을 SURVEY_FEEDBACK evidence로 인정
→ domain conflict만으로 EXCLUDED하지 않음
```

기계 capability까지 충족하면 `RESOLVED`, 부족하면 `STILL_DEFERRED`가 정상입니다.

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch8_1_semantic_audit_guardrails.zip `
  -DestinationPath . `
  -Force
```

## 신규 검증

```powershell
node --check .\automation\queryCandidateResolver.js
node --check .\tests\queryCandidateResolverDomainEvidenceSmokeTest.js
node --check .\tests\queryCandidateResolverIdentityGuardrailSmokeTest.js

node .\tests\queryCandidateResolverDomainEvidenceSmokeTest.js
node .\tests\queryCandidateResolverIdentityGuardrailSmokeTest.js
node .\tests\queryCandidatePatch8_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch8ManifestSmokeTest.js
```

## 기존 패치 8 검증

```powershell
node .\tests\queryCandidateResolverSmokeTest.js
node .\tests\queryCandidateResolverInferredSemanticSmokeTest.js
node .\tests\queryCandidateResolverSourceScopeSmokeTest.js
node .\tests\queryCandidateResolverConservativeExclusionSmokeTest.js
node .\tests\queryCandidateResolverIntegritySmokeTest.js
node .\tests\queryCandidateResolverPrivacyBoundarySmokeTest.js
node .\tests\queryCandidateResolverSchemaSmokeTest.js
node .\tests\queryCandidateResolverBaselineSmokeTest.js
node .\tests\queryCandidatePatch8SourceIntegritySmokeTest.js
```

## 기준선 재작성

정책 버전과 판정 결과가 의도적으로 변경되므로 기존 `candidate-resolution.json` 기준선을 다시 작성해야 합니다.

```powershell
node .\tests\queryCandidateResolverCapture.js `
  --mode=write

node .\tests\queryCandidateResolverCapture.js `
  --mode=compare
```

정상 기준:

```text
PASS 6/6
errors=0
differences=0
```

그 후 표본 감사를 다시 실행합니다.

```powershell
node .\tests\queryCandidateResolverSampleAudit.js `
  --resolved-limit=8 `
  --excluded-limit=8
```

## 범위

```text
OpenAI 호출 추가 없음
추가 토큰 비용 없음
Production route 변경 없음
READY 상태 부여 없음
후보 family 처리 없음
```

패치 9는 재감사에서 명백한 semanticResolved 오탐과 domain-only false exclusion이 해소된 뒤 진행합니다.
