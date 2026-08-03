# BeeBee AI Query Candidate Patch 8.2.1

## Named Template Primary Domain Anchor + Policy Metadata Fix

패치 8.2.1은 패치 8.2 재감사에서 확인된 두 문제를 수정합니다.

```text
1. 명명형 template 후보가 공통 보조 domain만 일치해도 RESOLVED되는 문제
2. candidate-resolution-index.json이 구 정책 버전 v1_1을 기록하는 문제
```

Production route, OpenAI 호출, Candidate Contract 상태는 변경하지 않습니다.

---

## 1. Named Template Primary Domain Anchor

`templateId`가 있는 명명형 후보는 후보 이름에서 확인된 핵심 업무영역이
데이터의 `classification.primaryDomain`과 일치해야 합니다.

예:

```text
후보: event_satisfaction_report
후보 domain: EVENT_ATTENDANCE + SURVEY_FEEDBACK
데이터 primary: EDUCATION_EVALUATION
데이터 secondary: SURVEY_FEEDBACK
```

기존에는 `SURVEY_FEEDBACK` 일치만으로 PASS할 수 있었지만, 패치 후에는:

```text
primary anchor: EVENT_ATTENDANCE
actual primary: EDUCATION_EVALUATION
→ NAMED_TEMPLATE_PRIMARY_DOMAIN_CONFLICT
→ EXCLUDED
```

`SURVEY_FEEDBACK`은 만족도·설문·피드백 후보에서 여러 업무영역에 공통으로
나타날 수 있으므로 명명형 후보의 primary anchor를 대신할 수 없습니다.

반면 다음 구조형 후보는 `templateId`가 없으므로 기존 동작을 유지합니다.

```text
소속별 만족도 평균
참가상태별 만족도 합계
신청일자별 만족도 추이
```

실제 만족도 열과 secondary semantic domain 근거가 확인되면 계속
`RESOLVED`될 수 있습니다.

---

## 2. Policy Metadata Fix

`tests/queryCandidateResolverCapture.js`가 정책 버전을 문자열로 중복 선언하지
않고 Resolver에서 export한 아래 상수를 사용합니다.

```text
QUERY_CANDIDATE_RESOLUTION_POLICY_VERSION
= deterministic_candidate_resolution_policy_v1_2
```

따라서 새 기준선 인덱스에는 다음이 기록돼야 합니다.

```json
{
  "policyVersion": "deterministic_candidate_resolution_policy_v1_2"
}
```

---

## 3. 적용

백엔드 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch8_2_1_named_template_primary_domain_anchor.zip `
  -DestinationPath . `
  -Force
```

---

## 4. 신규 검증

```powershell
node --check .\automation\queryCandidateResolver.js
node --check .\tests\queryCandidateResolverCapture.js

node .\tests\queryCandidateResolverNamedTemplatePrimaryDomainSmokeTest.js
node .\tests\queryCandidateResolverPolicyMetadataSmokeTest.js
node .\tests\queryCandidatePatch8_2_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8_2_1ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate resolver named template primary domain smoke
PASS query candidate resolver policy metadata smoke
PASS query candidate patch8.2.1 source integrity smoke
PASS query candidate patch8.2.1 manifest smoke
```

---

## 5. 이전 패치 무결성

```powershell
node .\tests\queryCandidatePatch8SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8ManifestSmokeTest.js
node .\tests\queryCandidatePatch8_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch8_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch8_2ManifestSmokeTest.js
```

모두 PASS여야 합니다.

---

## 6. 기준선 재작성

판정 결과와 index metadata가 변경되므로 기준선을 다시 작성합니다.

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

정책 버전 확인:

```powershell
(Get-Content `
  .\tests\fixtures\query-candidate-baseline\candidate-resolution-index.json `
  -Raw | ConvertFrom-Json).policyVersion
```

기대 결과:

```text
deterministic_candidate_resolution_policy_v1_2
```

---

## 7. 재감사

```powershell
node .\tests\queryCandidateResolverSampleAudit.js `
  --resolved-limit=12 `
  --excluded-limit=10
```

중점 확인:

```text
template_course_evaluation_report:
  event_satisfaction_report
  → RESOLVED 목록에서 제거
  → newlyExcluded 또는 EXCLUDED 목록에 표시
  → NAMED_TEMPLATE_PRIMARY_DOMAIN_CONFLICT

real_world_event_applicant_workshop:
  소속별·신청자별·참가상태별 만족도 후보
  → 구조형 후보이므로 기존 RESOLVED 유지
```

---

## 8. 다음 단계

재감사에서 명명형 primary domain 오탐이 제거되고 다른 구조형 후보 회귀가
없으면 패치 9 Candidate Family 및 중복 제거로 진행합니다.
