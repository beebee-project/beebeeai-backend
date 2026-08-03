# Patch 9 — Candidate Family & Deterministic Deduplication

## 목적

패치 8.2.1의 `candidate-resolution.json`에서 `RESOLVED`된 후보를 실제 생성 의도 기준으로 family화하고, 같은 family에서 대표 후보 하나만 선택합니다.

이 단계는 후보를 삭제하거나 `READY`로 바꾸지 않습니다. 모든 원본 후보를 보존하면서 다음 정보만 별도 결과로 기록합니다.

- `familyId`
- `familyCategory`
- `SELECTED | SUPPRESSED | NOT_APPLICABLE`
- `selectedCandidateId`
- `suppressionReason`

## 입력과 출력

```text
candidate-resolution.json
↓
candidate-family-resolution.json
```

출력 계약:

```text
query_candidate_family_resolution_v1
deterministic_candidate_family_policy_v1
```

## Family signature

다음 항목이 모두 같은 `RESOLVED` 후보만 동일 family로 묶습니다.

1. 실제 source root table
2. operation 또는 구조형 recipe
3. group·period·measure operand
4. 명명형 template anchor
5. output type

따라서 제목이 같더라도 operand가 다르면 합치지 않습니다.

```text
만족도 상위/하위 — 소속 + 만족도
만족도 상위/하위 — 신청자 + 만족도
만족도 상위/하위 — 참가상태 + 만족도
→ 서로 다른 family
```

반면 candidateId의 suffix만 다르고 source·recipe·operand가 모두 같으면 중복 family가 됩니다.

## Family category

- `DASHBOARD`
- `MULTI_SOURCE`
- `TIME_SERIES`
- `GROUP_AGGREGATION`
- `RANKING`
- `CROSS_TAB`
- `COUNTING`
- `OTHER`

category는 탐색과 후속 Feasibility Gate 분류용입니다. 중복 판정은 더 구체적인 exact signature로 수행합니다.

## 대표 후보 선택 순서

1. 패치 5에서 이미 `RETRIEVED`된 후보
2. `BOUND > PARTIAL > INFERRED > UNBOUND`
3. `DECLARED > GENERIC > UNKNOWN` executor
4. 강한 recipe operand 근거
5. 높은 Resolver 점수
6. 높은 원래 후보 점수
7. 낮은 원래 순위
8. candidateId 사전순 tie-break

억제 후보에는 다음과 같은 사유가 기록됩니다.

- `PRIOR_RETRIEVED_PREFERRED`
- `STRONGER_BINDING_PREFERRED`
- `DECLARED_EXECUTOR_PREFERRED`
- `STRONGER_OPERAND_EVIDENCE_PREFERRED`
- `HIGHER_RESOLUTION_SCORE`
- `HIGHER_ORIGINAL_SCORE`
- `LOWER_ORIGINAL_RANK`
- `LEXICOGRAPHIC_TIE_BREAK`

## 안전 경계

- `RESOLVED` 후보만 family화합니다.
- `STILL_DEFERRED`와 `EXCLUDED`는 `NOT_APPLICABLE`로 그대로 보존합니다.
- 원본 candidate-resolution 후보를 수정하거나 삭제하지 않습니다.
- production route를 변경하지 않습니다.
- `READY`를 부여하지 않습니다.
- OpenAI 호출을 추가하지 않습니다.
- 원본 row, sample value, XLSX를 출력에 포함하지 않습니다.

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch9_candidate_family_deduplication.zip `
  -DestinationPath . `
  -Force
```

## 파일 확인

```powershell
Test-Path .\automation\queryCandidateFamilyResolver.js
Test-Path .\automation\queryCandidateFamilyResolver.schema.json
Test-Path .\tests\queryCandidateFamilyResolverCapture.js
Test-Path .\tests\queryCandidateFamilyResolverSampleAudit.js
Test-Path .\PATCH_MANIFEST_PATCH9.json
```

## 문법 검사

```powershell
node --check .\automation\queryCandidateFamilyResolver.js
node --check .\tests\queryCandidateFamilyResolverCapture.js
node --check .\tests\queryCandidateFamilyResolverSampleAudit.js
```

## 신규 스모크

```powershell
node .\tests\queryCandidateFamilyResolverSmokeTest.js
node .\tests\queryCandidateFamilyResolverDuplicateSuppressionSmokeTest.js
node .\tests\queryCandidateFamilyResolverOperandSeparationSmokeTest.js
node .\tests\queryCandidateFamilyResolverNamedTemplateIsolationSmokeTest.js
node .\tests\queryCandidateFamilyResolverRepresentativeSelectionSmokeTest.js
node .\tests\queryCandidateFamilyResolverPrivacyBoundarySmokeTest.js
node .\tests\queryCandidateFamilyResolverSchemaSmokeTest.js
node .\tests\queryCandidateFamilyResolverBaselineSmokeTest.js
node .\tests\queryCandidatePatch9SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch9ManifestSmokeTest.js
```

## 기준선 작성

```powershell
node .\tests\queryCandidateFamilyResolverCapture.js `
  --mode=write

node .\tests\queryCandidateFamilyResolverCapture.js `
  --mode=compare
```

정상 조건:

```text
cases: 6
PASS 6/6
errors=0
differences=0
```

각 케이스에서 반드시 다음 관계가 성립해야 합니다.

```text
selected = familyCount
selected + suppressed = resolvedInput
```

생성 파일:

```text
tests\fixtures\query-candidate-baseline\<case-id>\candidate-family-resolution.json
tests\fixtures\query-candidate-baseline\candidate-family-resolution-index.json
```

## 실제 중복 감사

```powershell
node .\tests\queryCandidateFamilyResolverSampleAudit.js `
  --limit=20
```

생성 파일:

```text
tests\fixtures\query-candidate-baseline\candidate-family-sample-audit.md
tests\fixtures\query-candidate-baseline\candidate-family-sample-audit.json
```

감사에서는 다음을 확인합니다.

- candidateId suffix만 다른 동일 dashboard 후보가 하나로 묶이는지
- 서로 다른 source table 후보가 합쳐지지 않는지
- `sum`, `average`, `rank`가 서로 합쳐지지 않는지
- 소속·신청자·참가상태처럼 operand 축이 다른 후보가 유지되는지
- 서로 다른 templateId의 명명형 후보가 합쳐지지 않는지
- 대표 후보 선택 reason이 타당한지

## 패치 8.2.1 회귀 확인

```powershell
node .\tests\queryCandidateResolverNamedTemplatePrimaryDomainSmokeTest.js
node .\tests\queryCandidateResolverPolicyMetadataSmokeTest.js
node .\tests\queryCandidateResolverCapture.js --mode=compare
```
