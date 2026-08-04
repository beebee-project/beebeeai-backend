# 패치 12.1.2 — Live Shadow UTF-8 QA Documentation Hotfix

## 목적

실제 `gpt-5.6-terra` Live Shadow 검증에서 Provider·Resolver Re-entry는 성공했지만 Windows PowerShell의 기본 문자열 디코딩으로 한국어 JSON이 깨져 `ConvertFrom-Json`이 실패했습니다.

실제 성공 결과:

```text
Status                   SHADOW_COMPLETED
InvocationStatus         CALLED
ProviderCalls            1
Proposed                 3
Accepted                 3
Resolved                 3
Ready                    3
Ranked                   3
ProductionCandidateMerge false
ProductionRouteChanged   false
```

이번 패치는 production 코드와 fixture를 변경하지 않고 검증 문서만 정리합니다.

## 변경 범위

```text
README_QUERY_CANDIDATE_PATCH12_1_LIVE_SHADOW_RESOLVER_REENTRY_BRIDGE.md
  - PowerShell JSON 읽기에 -Encoding UTF8 명시
  - 실패한 node -e here-string 예제 사용 금지
  - 검증된 단일행 node -e 명령으로 교체
  - 실제 Provider 성공·실패 판정 기준 명시

tests/queryCandidatePatch12_1ManifestSmokeTest.js
  - 위 README와 historical manifest 검사 파일의 후속 문서 수정 허용
```

변경하지 않는 항목:

```text
automation 코드
OpenAI adapter
Shadow Bridge
Resolver
Family
Feasibility
Ranker
fixture
production route
기존 baseline
```

# 적용

```powershell
Expand-Archive `
  .\query_candidate_patch12_1_2_live_shadow_utf8_qa_documentation_hotfix.zip `
  -DestinationPath . `
  -Force
```

## 문법·문서 검사

```powershell
node --check .\tests\queryCandidatePatch12_1ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2DocumentationSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch12_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch12_1_2DocumentationSmokeTest.js
node .\tests\queryCandidatePatch12_1_2ManifestSmokeTest.js
```

기대 결과:

```text
PASS query candidate patch12.1 manifest smoke superseded=2
PASS query candidate patch12.1.2 documentation smoke
PASS query candidate patch12.1.2 manifest smoke
```

# PowerShell 결과 JSON 읽기

반드시 UTF-8을 명시합니다.

```powershell
$result = Get-Content `
  $providerOutput `
  -Raw `
  -Encoding UTF8 |
  ConvertFrom-Json
```

인코딩을 생략하면 한국어가 깨지고 정상 JSON도 `':' 또는 '}'가 필요합니다` 오류로 오인될 수 있습니다.

# Node.js 결과 JSON 검증

PowerShell here-string을 `node -e`에 직접 붙이지 않고 다음 단일행 명령을 사용합니다.

```powershell
node -e "const fs=require('fs'); const filePath=process.argv[1]; const result=JSON.parse(fs.readFileSync(filePath,'utf8')); console.log({status:result.status, invocationStatus:result.plannerResolution?.invocation?.status, responseId:result.plannerResolution?.invocation?.responseId, accepted:result.counts?.accepted, resolved:result.counts?.resolved, ready:result.counts?.ready, ranked:result.counts?.ranked, productionCandidateMerge:result.integrity?.productionCandidateMerge, productionRouteChanged:result.integrity?.productionRouteChanged}); console.log('PASS Node UTF-8 JSON parse');" "$providerOutput"
```

다음 방식은 사용하지 않습니다.

```powershell
# 사용 금지
node -e @'
...
'@ $providerOutput
```

정상 예상 출력:

```text
{
  status: 'SHADOW_COMPLETED',
  invocationStatus: 'CALLED',
  responseId: 'resp_...',
  accepted: 3,
  resolved: 3,
  ready: 3,
  ranked: 3,
  productionCandidateMerge: false,
  productionRouteChanged: false
}
PASS Node UTF-8 JSON parse
```

# 회귀 확인

이번 패치는 문서·QA 테스트만 변경하므로 기준선을 다시 작성하지 않습니다.

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
Shadow PASS 1/1 differences=0
Planner PASS 6/6 differences=0
Ranker PASS 6/6 differences=0
Feasibility PASS 6/6 differences=0
Family PASS 6/6 differences=0
Resolver PASS 6/6 differences=0
```
