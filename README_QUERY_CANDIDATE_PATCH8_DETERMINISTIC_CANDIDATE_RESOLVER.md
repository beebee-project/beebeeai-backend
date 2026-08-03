# BeeBee AI Query Candidate Patch 8

## Deterministic Candidate Resolver

패치 8은 패치 5의 `query_candidate_retrieval_v1`과 패치 7의
`resolved_semantic_profile_v1`을 결합해 후보를 다시 판정합니다.

```text
retrieval.json
+ capability-manifest.json
+ resolved-semantic-profile.json
↓
candidate-resolution.json
```

새 출력 계약:

```text
query_candidate_resolution_v1
```

후보 결과:

```text
RESOLVED
STILL_DEFERRED
EXCLUDED
```

`RESOLVED`는 의미·구조 요구조건을 결정론적으로 확인했다는 뜻입니다.
최종 `READY`가 아니며 실제 실행 가능성은 후속 Feasibility Gate가 판정합니다.

---

## 1. 판정 정책

### 패치 5 결과 유지

```text
RETRIEVED → RESOLVED
EXCLUDED  → EXCLUDED
DEFERRED  → 의미 재평가
```

기존 `EXCLUDED`는 명백한 필수조건 누락으로 제외된 결과이므로 보수적으로
유지합니다. LLM 의미가 추가됐다는 이유만으로 자동 복구하지 않습니다.

### DEFERRED 재평가

다음 조건을 내부 코드로 검사합니다.

```text
source table 귀속
필수 column role
필수 data/semantic type
required capability와 operation
metric family
테이블·행 제약
업무 domain 일치
recipe 연결
executor declaration
```

### INFERRED

`INFERRED` 후보는 다음 조건이 모두 확인되면 `RESOLVED`될 수 있습니다.

```text
source가 단일하게 확정
필수 역할·operation·metric 충족
업무영역 충돌 없음
recipe 존재
DECLARED 또는 GENERIC executor 연결 존재
```

`GENERIC` executor는 의미 단계에서는 허용하지만 실제 실행 지원은 후속
Feasibility Gate가 반드시 다시 검증합니다.

### UNBOUND

```text
UNBOUND → STILL_DEFERRED
```

manifest 요구조건이 없으므로 안전하게 확정하거나 제외할 수 없습니다.

### INFERRED 요구조건 실패

식별자 기반 추론 자체가 불완전할 수 있으므로 필수 역할이 보이지 않는다는
이유만으로 즉시 제외하지 않습니다.

```text
INFERRED + role/capability 불충족
→ STILL_DEFERRED
```

다만 데이터의 높은 confidence 업무영역과 후보의 강한 업무영역 신호가
충돌하면 `EXCLUDED`할 수 있습니다.

---

## 2. 다중 테이블 source 선택

명시적인 `sourceTableIds`가 있으면 해당 참조를 우선합니다.

source가 없고 물리 테이블이 여러 개이면 각 물리 root와 virtual 파생표를
묶어 후보 요구조건을 평가합니다.

```text
충족 root 1개  → RESOLVED 가능
충족 root 2개+ → STILL_DEFERRED
충족 root 0개  → binding에 따라 EXCLUDED 또는 STILL_DEFERRED
```

점수가 조금 높다는 이유만으로 임의의 primary table을 선택하지 않습니다.

---

## 3. 후보 상태 보호

패치 8은 Candidate Contract의 상태를 변경하지 않습니다.

```text
candidate status: UNASSESSED 유지
READY 판정: 없음
candidateId 생성: 없음
recipeId 생성: 없음
```

---

## 4. 비용과 보안

패치 8은 OpenAI API를 호출하지 않습니다.

```text
추가 LLM 호출: 0회
추가 토큰 비용: 0원
```

입력과 출력에 원본 행·셀 값·sampleValues를 포함하지 않습니다. 다만 후보
근거에 열 헤더와 의미 정보가 포함될 수 있으므로 production 저장 시에는
기존 암호화 저장 계층을 사용해야 합니다.

권장 파일명:

```text
candidate-resolution.enc
```

---

## 5. 포함 파일

```text
automation/
├─ queryCandidateResolver.js
└─ queryCandidateResolver.schema.json

tests/
├─ queryCandidateResolverTestSupport.js
├─ queryCandidateResolverSmokeTest.js
├─ queryCandidateResolverInferredSemanticSmokeTest.js
├─ queryCandidateResolverSourceScopeSmokeTest.js
├─ queryCandidateResolverConservativeExclusionSmokeTest.js
├─ queryCandidateResolverIntegritySmokeTest.js
├─ queryCandidateResolverPrivacyBoundarySmokeTest.js
├─ queryCandidateResolverSchemaSmokeTest.js
├─ queryCandidateResolverBaselineSmokeTest.js
├─ queryCandidateResolverCapture.js
├─ queryCandidatePatch8SourceIntegritySmokeTest.js
└─ queryCandidatePatch8ManifestSmokeTest.js
```

기존 production route와 기존 패치 1~7 파일은 변경하지 않습니다.

---

## 6. 적용

백엔드 루트에서 실행합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch8_deterministic_candidate_resolver.zip `
  -DestinationPath . `
  -Force
```

파일 확인:

```powershell
Test-Path .\automation\queryCandidateResolver.js
Test-Path .\automation\queryCandidateResolver.schema.json
Test-Path .\tests\queryCandidateResolverCapture.js
```

모두 `True`여야 합니다.

---

## 7. 문법 검사

```powershell
node --check .\automation\queryCandidateResolver.js
node --check .\tests\queryCandidateResolverCapture.js
```

---

## 8. 신규 스모크

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
node .\tests\queryCandidatePatch8ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate resolver smoke
PASS query candidate resolver inferred semantic smoke
PASS query candidate resolver source scope smoke
PASS query candidate resolver conservative exclusion smoke
PASS query candidate resolver integrity smoke
PASS query candidate resolver privacy boundary smoke
PASS query candidate resolver schema smoke
PASS query candidate resolver baseline smoke 6
PASS query candidate patch8 source integrity smoke
PASS query candidate patch8 manifest smoke
```

사용자 저장소의 실제 fixture가 6개이므로 baseline smoke는 `6`이 기대됩니다.

---

## 9. 기준선 작성

```powershell
node .\tests\queryCandidateResolverCapture.js `
  --mode=write
```

각 케이스에 생성됩니다.

```text
tests\fixtures\query-candidate-baseline\<case-id>\candidate-resolution.json
```

인덱스:

```text
tests\fixtures\query-candidate-baseline\candidate-resolution-index.json
```

정상 기준:

```text
[query-candidate-resolver] cases: 6
[query-candidate-resolver] PASS 6/6
errors=0
```

실제 케이스의 `resolved·stillDeferred·excluded` 수는 후보 이름, manifest,
테이블 구조에 따라 달라지므로 특정 개수를 기대하지 않습니다.

반드시 다음 관계가 성립해야 합니다.

```text
resolved + stillDeferred + excluded = total
```

`semanticResolved`는 패치 5에서 `DEFERRED`였지만 패치 7의 의미 profile을
이용해 새로 확정된 후보 수입니다.

---

## 10. 재현성 비교

```powershell
node .\tests\queryCandidateResolverCapture.js `
  --mode=compare
```

정상 기준:

```text
[query-candidate-resolver] PASS 6/6
errors=0
differences=0
```

---

## 11. 이전 단계 회귀

```powershell
node .\tests\querySemanticProfileMergeCapture.js --mode=compare
node .\tests\querySemanticProfilerCapture.js --mode=compare
node .\tests\queryCandidateRetrieverCapture.js --mode=compare
node .\tests\queryCandidateCapabilityCapture.js --mode=compare
node .\tests\queryCandidateContractCapture.js --mode=compare
node .\tests\queryJsonSemanticProfileCapture.js --mode=compare
```

모두 기존 기준선과 `differences=0`이어야 합니다.

---

## 12. 결과 해석

### RESOLVED

```text
병합된 의미로 필수조건 확인
source table 확정
업무영역 충돌 없음
recipe/executor 연결 존재
```

아직 `READY`는 아닙니다.

### STILL_DEFERRED

대표 원인:

```text
UNBOUND capability
다중 테이블 동률
source 참조 불명확
recipe 또는 executor 연결 미확정
낮은 domain confidence
INFERRED 요구조건 불충족
```

### EXCLUDED

```text
기존 패치 5 EXCLUDED 유지
높은 confidence 업무영역 충돌
분석 가능한 테이블 없음
BOUND/PARTIAL 명시 요구조건 불충족
```

---

## 13. 다음 단계

패치 9에서는 `RESOLVED` 후보를 candidate family로 묶고 의미상 중복 후보를
제거합니다. 그 뒤 패치 10 Ranker와 패치 11 Feasibility Gate가 실제 사용자
노출 순위와 `READY·CONDITIONAL·UNSUPPORTED·REJECTED` 상태를 판정합니다.
