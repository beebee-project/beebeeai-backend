# Patch 12.1.2.1 — Historical Manifest Supersession Restore + Windows PowerShell Node Quoting Fix

## 목적

Patch 12.1.2 적용 후 확인된 두 QA 오류만 수정합니다.

1. `queryCandidatePatch12_1ManifestSmokeTest.js`가 누적 supersession 정책을 잃고 `automation/queryCandidateResolver.js`를 Patch 12.1 당시 byte/SHA와 다시 비교한 오류
2. Windows PowerShell에서 단일 인용부호로 감싼 `node -e` JavaScript의 내부 큰따옴표가 제거되어 `require(fs)`, `utf8`, `console.log(PASS`로 손상된 오류

이번 패치는 production 코드, Resolver, Planner, Shadow Bridge, fixture 및 기준선을 변경하지 않습니다.

## 적용 순서

```text
Patch 12.1.1.2
→ Patch 12.1.2
→ Patch 12.1.2.1
→ Patch 13 — Encrypted Hierarchical Cache
```

## 수정 1 — Patch 12.1 historical manifest supersession 복원

Patch 12.1 이후 합법적으로 변경된 파일만 현재 exact hash 비교에서 제외합니다.

```text
automation/queryCandidateResolver.js
README_QUERY_CANDIDATE_PATCH12_1_LIVE_SHADOW_RESOLVER_REENTRY_BRIDGE.md
tests/queryCandidatePatch12_1ManifestSmokeTest.js
```

`automation/queryCandidateResolver.js`는 Patch 12.1.1 계열 validation에서, QA 문서와 manifest 검사는 Patch 12.1.2.1 validation에서 superseded 경로를 가져옵니다.

변경되지 않은 Patch 12.1 파일은 계속 byte size와 SHA-256을 정확히 검증합니다. 과거 `PATCH_MANIFEST_PATCH12_1.json` 값 자체는 수정하지 않습니다.

정상 예상 결과:

```text
PASS query candidate patch12.1 manifest smoke superseded=3
```

## 수정 2 — Patch 12.1.2 historical manifest 처리

Patch 12.1.2.1에서 다시 수정하는 Patch 12.1.2 산출물은 다음과 같습니다.

```text
README_QUERY_CANDIDATE_PATCH12_1_LIVE_SHADOW_RESOLVER_REENTRY_BRIDGE.md
tests/queryCandidatePatch12_1ManifestSmokeTest.js
tests/queryCandidatePatch12_1_2DocumentationSmokeTest.js
tests/queryCandidatePatch12_1_2ManifestSmokeTest.js
README_QUERY_CANDIDATE_PATCH12_1_2_LIVE_SHADOW_UTF8_QA_DOCUMENTATION_HOTFIX.md
```

정상 예상 결과:

```text
PASS query candidate patch12.1.2 manifest smoke superseded=5
```

## 수정 3 — Windows PowerShell-safe Node UTF-8 검증 명령

PowerShell JSON 읽기는 UTF-8을 명시합니다.

```powershell
$result = Get-Content `
  $providerOutput `
  -Raw `
  -Encoding UTF8 |
  ConvertFrom-Json
```

Node 검증은 Windows PowerShell에서 바깥쪽 큰따옴표, JavaScript 문자열 작은따옴표를 사용합니다. 파일 경로 인수도 큰따옴표로 감쌉니다.

```powershell
node -e "const fs=require('fs'); const filePath=process.argv[1]; const result=JSON.parse(fs.readFileSync(filePath,'utf8')); console.log({status:result.status, invocationStatus:result.plannerResolution?.invocation?.status, responseId:result.plannerResolution?.invocation?.responseId, accepted:result.counts?.accepted, resolved:result.counts?.resolved, ready:result.counts?.ready, ranked:result.counts?.ranked, productionCandidateMerge:result.integrity?.productionCandidateMerge, productionRouteChanged:result.integrity?.productionRouteChanged}); console.log('PASS Node UTF-8 JSON parse');" "$providerOutput"
```

다음 형태는 사용하지 않습니다.

```powershell
# 사용 금지: Windows PowerShell native argument quoting 손상 가능
node -e 'const fs=require("fs"); ...' $providerOutput

# 사용 금지: node -e here-string 전달
node -e @'
...
'@ $providerOutput
```

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch12_1_2_1_historical_manifest_supersession_restore_windows_powershell_node_quoting_fix.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check .\tests\queryCandidatePatch12_1ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2DocumentationSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2_1SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2_1ManifestSmokeTest.js
```

## 오류 재검증

```powershell
node .\tests\queryCandidatePatch12_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1_1_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_1_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1_2DocumentationSmokeTest.js
node .\tests\queryCandidatePatch12_1_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1_2_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_2_1ManifestSmokeTest.js
```

정상 예상 결과:

```text
PASS query candidate patch12.1 source integrity smoke
PASS query candidate patch12.1 manifest smoke superseded=3
PASS query candidate patch12.1.1.2 source integrity smoke
PASS query candidate patch12.1.1.2 manifest smoke
PASS query candidate patch12.1.2 documentation smoke
PASS query candidate patch12.1.2 manifest smoke superseded=5
PASS query candidate patch12.1.2.1 source integrity smoke
PASS query candidate patch12.1.2.1 manifest smoke
```

## Node 결과 파일 재검증

실제 Provider를 다시 호출하지 않고 기존 `$providerOutput`을 검증합니다.

```powershell
$result = Get-Content `
  $providerOutput `
  -Raw `
  -Encoding UTF8 |
  ConvertFrom-Json

node -e "const fs=require('fs'); const filePath=process.argv[1]; const result=JSON.parse(fs.readFileSync(filePath,'utf8')); console.log({status:result.status, invocationStatus:result.plannerResolution?.invocation?.status, responseId:result.plannerResolution?.invocation?.responseId, accepted:result.counts?.accepted, resolved:result.counts?.resolved, ready:result.counts?.ready, ranked:result.counts?.ranked, productionCandidateMerge:result.integrity?.productionCandidateMerge, productionRouteChanged:result.integrity?.productionRouteChanged}); console.log('PASS Node UTF-8 JSON parse');" "$providerOutput"
```

정상 출력의 핵심값:

```text
status: SHADOW_COMPLETED
invocationStatus: CALLED
responseId: resp_...
accepted = resolved = ready = ranked = 3
productionCandidateMerge: false
productionRouteChanged: false
PASS Node UTF-8 JSON parse
```

## 전체 회귀

기준선은 다시 작성하지 않고 비교만 실행합니다.

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

# Patch 13으로 넘어가는 시점

## 결론

Patch 12.1.2.1의 위 검사가 전부 통과하면 **바로 Patch 13 — 암호화 계층형 캐시로 넘어가도 됩니다.** 추가 Live Provider 호출은 필요하지 않습니다.

이미 충족된 기능 선행조건:

```text
실제 gpt-5.6-terra 호출                 PASS
Response ID 수신                        PASS
실제 토큰·비용 기록                     PASS
Live proposal 3개                       PASS
Resolver → Family → Feasibility → Ranker PASS 3/3
Production candidate merge              false
Production route changed                false
기존 암호화 단일 Planner cache primitive 존재
```

Patch 13 전에 마지막으로 닫아야 하는 조건은 QA 계약뿐입니다.

```text
Patch 12.1 manifest                 PASS superseded=3
Patch 12.1.2 manifest               PASS superseded=5
Patch 12.1.2 documentation          PASS
Patch 12.1.2.1 source/manifest      PASS
Shadow                              PASS 1/1
결정론적 5개 계층                    PASS 6/6
```

## 왜 그 이후인가

암호화 계층형 캐시는 다음 결과를 장기간 재사용하는 저장 계층입니다.

```text
L1: 동일 프로세스 메모리 cache
L2: 동일 업로드/동일 queryJson 암호화 cache
L3: 동일 의미 profile·Planner input 암호화 cache
L4: 검증된 proposal 및 re-entry 결과 암호화 cache
```

호출·재진입·QA 계약이 안정되기 전에 캐시를 추가하면, 잘못된 결과가 캐시된 것인지 기존 Resolver/문서/manifest 회귀인지 구분하기 어려워집니다. 현재 Live 호출과 재진입 기능은 이미 검증됐으므로 Patch 12.1.2.1에서 QA 회귀만 닫으면 캐시 설계의 기준선이 고정됩니다.

## Patch 13의 권장 첫 범위

Patch 13 첫 단계에서는 production route 연결보다 **암호화 계층형 cache contract와 deterministic key 설계**부터 시작합니다.

```text
cache namespace/version
업로드 fingerprint
queryJson SHA
resolved semantic profile SHA
Planner input SHA
model ID + reasoning effort + schema version
proposal validation policy version
Resolver/Family/Feasibility/Ranker policy version
TTL와 invalidation reason
negative cache 금지 또는 짧은 TTL
FAILED_SAFE·credit_balance_exhausted 비캐시
암호화 .enc 저장만 허용
plaintext 파일 생성 금지
```

그 다음 Shadow 모드에서 `CACHE_HIT`과 실제 재진입 결과의 SHA 동등성을 확인한 뒤 production 연결 여부를 결정합니다.
