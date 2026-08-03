# 패치 12 — Conditional LLM Candidate Planner

## 1. 목적

Deterministic Resolver → Family → Feasibility → Ranker 결과가 충분하면 LLM을 호출하지 않습니다.
결정론적 후보가 실제로 부족한 복합 데이터에서만 LLM이 최대 3개의 보완 후보를 제안합니다.

```text
candidate-resolution.json
candidate-feasibility-resolution.json
candidate-ranking-resolution.json
semantic-profile.json
↓
조건부 호출 판단
↓
candidate-planner-resolution.json
```

새 계약:

```text
query_candidate_planner_input_v1
query_candidate_planner_model_output_v1
query_candidate_planner_resolution_v1
query_candidate_planner_item_v1
conditional_llm_candidate_planner_policy_v1
```

## 2. 호출 조건

### LLM 호출 생략

```text
READY 후보 3개 이상
분석 가능한 물리 테이블 없음
지원 가능한 의미 기회 없음
소규모 데이터에 기존 READY 후보 존재
기존 READY 후보가 현재 의미 기회를 충분히 대표
```

### LLM 호출 필요

```text
분석 가능한 데이터인데 READY 후보가 0개

또는

행 8개 이상 복합 데이터에서 READY가 1~2개뿐이고
OVERVIEW / GROUP_AGGREGATION / TIME_SERIES / RANKING / CROSS_TAB
의미 범주가 2개 이상 비어 있음

또는

미해결 후보와 의미 범주 공백이 동시에 존재
```

`REVIEW` 후보가 있다는 이유만으로 호출하지 않습니다. 이미 READY 추천 후보가 충분하면 결정론적 결과를 우선합니다.

## 3. 호출 상태

```text
SKIPPED
REQUIRED_NOT_RUN
CALLED
CACHE_HIT
FAILED_SAFE
```

- `SKIPPED`: 호출 조건 불충족, 비용 0
- `REQUIRED_NOT_RUN`: 호출 조건은 충족했지만 provider 미설정
- `CALLED`: provider 1회 호출
- `CACHE_HIT`: 암호화 캐시 재사용, 신규 호출 0
- `FAILED_SAFE`: provider 또는 strict output 실패, 기존 결정론적 결과 유지

## 4. 제안 가능한 operation

```text
count_rows
category_count
group_sum
group_avg
top_bottom
time_sum
time_avg
time_count
cumulative_sum
cross_sum
cross_count
```

각 제안은 다음 조건을 모두 만족해야 합니다.

```text
기존 semantic-profile에 존재하는 물리 tableId 1개
기존 semantic-profile에 존재하는 columnId만 사용
operation별 필수 operand kind 정확히 일치
summarysheet 출력만 사용
confidence 0.72 이상
기존 후보와 같은 실행 signature가 아님
최대 3개
```

LLM이 임의 tableId·columnId·operation을 만들면 결정론적 validator가 `REJECTED` 처리합니다.

## 5. 제안 상태

```text
ACCEPTED_FOR_REVALIDATION
REJECTED
```

`ACCEPTED_FOR_REVALIDATION`은 READY가 아닙니다.

```text
LLM Planner 제안
→ Candidate Contract 변환
→ Resolver 재진입
→ Family
→ Feasibility
→ Ranker
```

이 재검증을 통과하기 전에는 사용자 후보군이나 production 실행에 사용할 수 없습니다.

## 6. 개인정보 경계

LLM 입력에는 다음만 포함됩니다.

```text
tableId
columnId
열 제목
semantic role/type
metric family
rowCount / columnCount
uniqueRatio / nonEmptyRatio
결정론적 후보의 operation·상태·reason code
```

다음은 포함하지 않습니다.

```text
원본 행
sampleValues
원본 파일
파일명
복호화 파일 내용
개인 식별값
```

OpenAI Responses 요청은 `store: false`와 strict JSON schema를 사용합니다.

## 7. 캐시

Planner 캐시는 암호화 codec을 주입받는 별도 모듈입니다.

```text
queryCandidatePlannerEncryptedCache.js
```

- 파일 확장자 `.enc`
- plaintext 파일 저장 없음
- tenantId + inputSha256 + model + prompt/policy version으로 HMAC cache key 생성
- 동일 입력은 `CACHE_HIT` 가능
- 업로드 삭제 정책과 연결하는 것은 production route 단계에서 수행

## 8. 적용

```powershell
Expand-Archive `
  .\query_candidate_patch12_conditional_llm_candidate_planner.zip `
  -DestinationPath . `
  -Force
```

파일 확인:

```powershell
Test-Path .\automation\queryCandidatePlanner.js
Test-Path .\automation\queryCandidatePlanner.schema.json
Test-Path .\automation\queryCandidatePlannerPrompt.js
Test-Path .\automation\queryCandidatePlannerOpenAIAdapter.js
Test-Path .\automation\queryCandidatePlannerEncryptedCache.js

Test-Path .\tests\queryCandidatePlannerCapture.js
Test-Path .\tests\queryCandidatePlannerSampleAudit.js
Test-Path .\PATCH_VALIDATION_PATCH12.json
Test-Path .\PATCH_MANIFEST_PATCH12.json
```

모두 `True`여야 합니다.

## 9. 문법 검사

```powershell
node --check .\automation\queryCandidatePlanner.js
node --check .\automation\queryCandidatePlannerPrompt.js
node --check .\automation\queryCandidatePlannerOpenAIAdapter.js
node --check .\automation\queryCandidatePlannerEncryptedCache.js
node --check .\tests\queryCandidatePlannerCapture.js
node --check .\tests\queryCandidatePlannerSampleAudit.js
```

## 10. 신규 기능 테스트

```powershell
node .\tests\queryCandidatePlannerTriggerSmokeTest.js
node .\tests\queryCandidatePlannerConditionalInvocationSmokeTest.js
node .\tests\queryCandidatePlannerProposalValidationSmokeTest.js
node .\tests\queryCandidatePlannerInvalidReferenceSmokeTest.js
node .\tests\queryCandidatePlannerDuplicateProposalSmokeTest.js
node .\tests\queryCandidatePlannerFailureSafeSmokeTest.js
node .\tests\queryCandidatePlannerEncryptedCacheSmokeTest.js
node .\tests\queryCandidatePlannerPrivacyBoundarySmokeTest.js
node .\tests\queryCandidatePlannerOpenAIAdapterSmokeTest.js
node .\tests\queryCandidatePlannerSchemaSmokeTest.js

node .\tests\queryCandidatePatch12SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate planner trigger smoke
PASS query candidate planner conditional invocation smoke
PASS query candidate planner proposal validation smoke
PASS query candidate planner invalid reference smoke
PASS query candidate planner duplicate proposal smoke
PASS query candidate planner failure safe smoke
PASS query candidate planner encrypted cache smoke
PASS query candidate planner privacy boundary smoke
PASS query candidate planner OpenAI adapter smoke
PASS query candidate planner schema smoke
PASS query candidate patch12 source integrity smoke
PASS query candidate patch12 manifest smoke
```

테스트의 OpenAI adapter는 mock client를 사용하므로 API 비용이 발생하지 않습니다.

## 11. 6개 기준선 작성

기준선 capture는 `decision-only` 모드입니다. 실제 OpenAI provider를 전달하지 않으며 API 호출이 없습니다.

```powershell
node .\tests\queryCandidatePlannerCapture.js `
  --mode=write

node .\tests\queryCandidatePlannerCapture.js `
  --mode=compare

node .\tests\queryCandidatePlannerBaselineSmokeTest.js
```

정상 기준:

```text
[query-candidate-planner] cases: 6
[query-candidate-planner] PASS 6/6
errors=0
warnings=0
differences=0

PASS query candidate planner baseline smoke 6
```

생성 파일:

```text
tests\fixtures\query-candidate-baseline\<case-id>\candidate-planner-resolution.json

tests\fixtures\query-candidate-baseline\candidate-planner-resolution-index.json
```

현재 실제 6개 기준선의 예상 결과:

```text
hardcase_two_tables_one_sheet_waste
→ READY 14개이므로 SKIPPED

real_world_event_applicant_workshop
→ READY 14개이므로 SKIPPED

seed_attendance_conditional
→ 소규모 데이터 + READY 1개이므로 SKIPPED

seed_sales_ready
→ 소규모 데이터 + READY 2개이므로 SKIPPED

seed_unstructured_unsupported
→ 분석 가능한 물리 테이블이 없어 SKIPPED

template_course_evaluation_report
→ READY 21개이므로 SKIPPED
```

따라서 6개 기준선의 정상 예상 신규 LLM 호출 수는 `0`입니다.

## 12. 표본 감사

```powershell
node .\tests\queryCandidatePlannerSampleAudit.js
```

생성 파일:

```text
tests\fixtures\query-candidate-baseline\candidate-planner-sample-audit.json
tests\fixtures\query-candidate-baseline\candidate-planner-sample-audit.md
```

감사 항목:

```text
trigger required
trigger reason code
invocation status
provider call count
READY count
semantic opportunity count
missing category count
accepted/rejected proposal count
```

## 13. 기존 계층 회귀 확인

```powershell
node .\tests\queryCandidateRankerCapture.js `
  --mode=compare

node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=compare

node .\tests\queryCandidateFamilyResolverCapture.js `
  --mode=compare

node .\tests\queryCandidateResolverCapture.js `
  --mode=compare
```

모두 `PASS 6/6`, `differences=0`이어야 합니다.

## 14. 현재 비연결 경계

```text
Production route 변경                  없음
실제 OpenAI provider 자동 연결         없음
Planner 제안의 Candidate Contract 병합 없음
Planner 제안의 READY 부여              없음
사용자 UI 노출                         없음
원본 후보 상태 변경                    없음
plaintext persistence                  없음
```

패치 12는 조건 판단·strict LLM adapter·암호화 캐시·제안 검증 계층까지만 추가합니다.
