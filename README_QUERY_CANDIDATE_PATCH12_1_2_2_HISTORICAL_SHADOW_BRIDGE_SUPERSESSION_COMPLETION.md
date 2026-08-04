# Patch 12.1.2.2 — Historical Shadow Bridge Supersession Completion

## 목적

Patch 12.1.2.1 적용 후 남은 단일 QA 오류를 닫습니다.

```text
historical manifest mismatch: automation/queryCandidatePlannerShadowBridge.js
```

Shadow, Planner, Ranker, Feasibility, Family, Resolver 기준선은 모두 `differences=0`으로 통과했고 Windows PowerShell UTF-8 검증도 완료됐습니다. 따라서 이번 패치는 Shadow Bridge production 코드를 덮어쓰지 않고, 현재 누적 저장소의 검증된 Shadow Bridge를 Patch 12.1 historical manifest의 명시적 superseded 산출물로 처리합니다.

## 변경 경계

```text
Production 코드 변경              없음
Shadow Bridge 파일 덮어쓰기        없음
Resolver/Planner 변경              없음
Fixture/기준선 변경                없음
Production route 변경              없음
Historical manifest 원본값 수정    없음
QA 테스트·validation·문서만 변경   있음
```

## 핵심 정책

`automation/queryCandidatePlannerShadowBridge.js`는 단순히 무조건 제외하지 않습니다. Patch 12.1.2.2 Source Integrity가 현재 파일을 직접 로드하고 mock Shadow 전 체인을 실행해 다음 계약을 검증한 경우에만 historical supersession을 인정합니다.

```text
provider calls                         1
status                                 SHADOW_COMPLETED
invocation                             CALLED
accepted/resolved/ready/ranked         2/2/2/2
Shadow resolution validation           valid
Production candidate merge             false
Production READY assignment            false
Production route changed               false
Source candidate/ranking mutation      없음
```

변경되지 않은 다른 Patch 12.1 파일은 계속 byte size와 SHA-256을 정확히 검사합니다.

## 적용 순서

```text
Patch 12.1.1.2
→ Patch 12.1.2
→ Patch 12.1.2.1
→ Patch 12.1.2.2
→ Patch 13 — Encrypted Hierarchical Cache
```

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch12_1_2_2_historical_shadow_bridge_supersession_completion.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check .\tests\queryCandidatePatch12_1ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2_1ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2_2SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2_2ManifestSmokeTest.js
```

## 누적 QA 검사

```powershell
node .\tests\queryCandidatePatch12_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1_1_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_1_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1_2DocumentationSmokeTest.js
node .\tests\queryCandidatePatch12_1_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1_2_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_2_1ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1_2_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_2_2ManifestSmokeTest.js
```

사용자 누적 저장소의 정상 예상 결과:

```text
PASS query candidate patch12.1 source integrity smoke
PASS query candidate patch12.1 manifest smoke superseded=4
PASS query candidate patch12.1.1.2 source integrity smoke
PASS query candidate patch12.1.1.2 manifest smoke
PASS query candidate patch12.1.2 documentation smoke
PASS query candidate patch12.1.2 manifest smoke superseded=5
PASS query candidate patch12.1.2.1 source integrity smoke
PASS query candidate patch12.1.2.1 manifest smoke superseded=2
PASS query candidate patch12.1.2.2 source integrity smoke
PASS query candidate patch12.1.2.2 manifest smoke
```

깨끗한 조립 저장소에서 Shadow Bridge가 Patch 12.1 원본과 아직 동일하면 첫 historical manifest의 `superseded` 값은 누적 적용 상태에 따라 2~3일 수 있습니다. 사용자 누적 저장소에서는 현재 Shadow Bridge 차이가 확인됐으므로 4가 정상입니다.

## 전체 회귀

기준선을 다시 작성하지 않고 비교만 실행합니다.

```powershell
node .\tests\queryCandidatePlannerShadowCapture.js `
  --mode=compare

node .\tests\queryCandidatePlannerCapture.js `
  --mode=compare

node .\tests\queryCandidateRankerCapture.js `
  --mode=compare

node .\tests\queryCandidateFeasibilityGateCapture.js `
  --mode=compare

node .\tests\queryCandidateFamilyResolverCapture.js `
  --mode=compare

node .\tests\queryCandidateResolverCapture.js `
  --mode=compare
```

정상 조건:

```text
Shadow       PASS 1/1 differences=0
Planner      PASS 6/6 differences=0
Ranker       PASS 6/6 differences=0
Feasibility  PASS 6/6 differences=0
Family       PASS 6/6 differences=0
Resolver     PASS 6/6 differences=0
```

`--mode=write`는 실행하지 않습니다.

## Patch 13 진입 조건

위 누적 QA와 전체 회귀가 전부 통과하면 Patch 13 — 암호화 계층형 캐시로 바로 넘어갑니다. 실제 Provider 호출은 이미 완료됐으므로 추가 유료 호출은 필요하지 않습니다.

Patch 13의 첫 단계는 production route 연결이 아니라 다음 계약부터 고정합니다.

```text
cache namespace/version
결정론적 cache key
업로드 fingerprint
queryJson SHA
semantic profile SHA
Planner input SHA
model/reasoning/schema/policy version
암호화 .enc 저장만 허용
FAILED_SAFE·결제·인증 실패 비캐시
TTL·invalidation reason
첫 실행 CALLED / 두 번째 CACHE_HIT parity
Production merge/route false 유지
```
