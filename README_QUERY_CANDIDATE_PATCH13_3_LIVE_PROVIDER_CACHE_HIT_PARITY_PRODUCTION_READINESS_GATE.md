# Query Candidate Patch 13.3

## Live Provider Cache-Hit Parity & Production Readiness Gate

### 적용 순서

Patch 13.2 적용 및 검증 완료 후 적용합니다.

```text
Patch 13
→ Patch 13.1
→ Patch 13.2
→ Patch 13.3
```

## 목적

실제 Provider 호출 1회로 생성한 Shadow 결과와 암호화 계층형 캐시에서 복원한 결과가 완전히 동일한지 검증합니다.

```text
첫 실행
→ 실제 Provider CALLED
→ L3 Planner Provider Result 저장
→ L3 Planner Resolution 저장
→ L4 Shadow Re-entry 저장

새 cache instance 재실행
→ Provider CACHE_HIT
→ Provider 추가 호출 0회
→ Planner Resolution L3_SEMANTIC hit
→ Re-entry L4_REENTRY hit
```

검증 결과를 기반으로 Production 연결 준비 상태를 판정하지만 Production route를 자동으로 변경하지 않습니다.

## 신규 계약

```text
query_candidate_planner_live_cache_parity_audit_v1
query_candidate_planner_production_readiness_gate_v1
candidate_planner_production_readiness_policy_v1
```

Readiness gate의 성공 결정은 다음과 같습니다.

```text
ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW
```

이 값은 Production 자동 승인을 뜻하지 않습니다.

```text
productionPromotionAllowed       false
productionRouteAutoWired         false
productionCandidateMergeAllowed  false
productionReadyAssignmentAllowed false
manualPromotionReviewRequired    true
failClosed                       true
```

## 필수 Live parity 조건

```text
Origin status                         SHADOW_COMPLETED
Origin invocation                     CALLED
Origin providerCallCount              1
실제 관찰 Provider 호출               1
Origin responseId                     존재
Origin totalTokens                    1 이상
Origin failureCode                    빈 값
Origin accepted                       1 이상
Origin accepted=resolved=ready=ranked

Replay status                         SHADOW_COMPLETED
Replay invocation                     CACHE_HIT
Replay providerCallCount              0
Replay Planner Provider cacheHit      true
Replay Planner Resolution source      L3_SEMANTIC
Replay Re-entry source                L4_REENTRY
Replay accepted=resolved=ready=ranked
Origin/Replay replay-safe SHA          동일

영구 캐시 파일                        최소 3개
영구 캐시 확장자                      .enc만 허용
평문 캐시 파일                        0개
Production merge/READY/route          모두 false
```

## 개인정보 및 감사 경계

Parity audit와 readiness gate에는 다음 값을 기록하지 않습니다.

```text
실제 responseId 값
토큰 사용량 수치
tenantId
원본 파일명
rawRows
sampleValues
캐시 파일 전체 경로 목록
```

대신 다음 최소 증거만 기록합니다.

```text
responseId 존재 여부
토큰 사용량 양수 여부
Provider 호출 횟수
Origin/Replay invocation 상태
L3/L4 cache source
암호화 파일 수
평문 파일 수
Replay-safe SHA
Production 격리 상태
```

## 변경 파일

```text
automation/queryCandidatePlannerProductionReadinessGate.js
automation/queryCandidatePlannerProductionReadinessGate.schema.json

tests/queryCandidatePatch13_3ProductionReadinessGateSmokeTest.js
tests/queryCandidatePatch13_3ProductionReadinessFailClosedSmokeTest.js
tests/queryCandidatePatch13_3LiveProviderCacheHitParitySmokeTest.js
tests/queryCandidatePatch13_3SourceIntegritySmokeTest.js
tests/queryCandidatePatch13_3ManifestSmokeTest.js

PATCH_VALIDATION_PATCH13_3.json
PATCH_MANIFEST_PATCH13_3.json
README_QUERY_CANDIDATE_PATCH13_3_LIVE_PROVIDER_CACHE_HIT_PARITY_PRODUCTION_READINESS_GATE.md
```

기존 Production 파일과 route는 변경하지 않습니다.

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch13_3_live_provider_cache_hit_parity_production_readiness_gate.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check `
  .\automation\queryCandidatePlannerProductionReadinessGate.js

node --check `
  .\tests\queryCandidatePatch13_3ProductionReadinessGateSmokeTest.js

node --check `
  .\tests\queryCandidatePatch13_3ProductionReadinessFailClosedSmokeTest.js

node --check `
  .\tests\queryCandidatePatch13_3LiveProviderCacheHitParitySmokeTest.js

node --check `
  .\tests\queryCandidatePatch13_3SourceIntegritySmokeTest.js

node --check `
  .\tests\queryCandidatePatch13_3ManifestSmokeTest.js
```

## 비용 없는 계약 검사

```powershell
node `
  .\tests\queryCandidatePatch13_3ProductionReadinessGateSmokeTest.js

node `
  .\tests\queryCandidatePatch13_3ProductionReadinessFailClosedSmokeTest.js

node `
  .\tests\queryCandidatePatch13_3SourceIntegritySmokeTest.js

node `
  .\tests\queryCandidatePatch13_3ManifestSmokeTest.js
```

정상 결과:

```text
PASS query candidate patch13.3 production readiness gate smoke
PASS query candidate patch13.3 production readiness fail-closed smoke
PASS query candidate patch13.3 source integrity smoke
PASS query candidate patch13.3 manifest smoke
```

## 실제 Provider Cache-Hit Parity 검사

이 검사는 실제 Provider를 정확히 1회 호출합니다. 두 번째 실행은 같은 암호화 캐시를 새 cache instance에서 읽어야 하며 Provider를 다시 호출하면 실패합니다.

### 1. API 키 확인

API 키 원문은 출력하지 않습니다.

```powershell
if ([string]::IsNullOrWhiteSpace($env:OPENAI_API_KEY)) {
  $secureKey = Read-Host "OpenAI API Key" -AsSecureString
  $credential = New-Object System.Net.NetworkCredential("", $secureKey)
  $env:OPENAI_API_KEY = $credential.Password
  $secureKey = $null
  $credential = $null
}

"OPENAI_API_KEY loaded: $($env:OPENAI_API_KEY.Length) characters"
```

### 2. Live parity 환경변수

기존 실호출에서 사용한 접근 가능한 모델 값을 유지합니다.

```powershell
$env:QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY = "1"
$env:QUERY_CANDIDATE_PLANNER_REASONING_EFFORT = "low"

$liveParityOutput = Join-Path `
  $PWD `
  "tests\fixtures\query-candidate-planner-shadow\call_required_group_avg_time_count\candidate-planner-live-cache-parity-readiness.json"

$env:QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_OUTPUT = $liveParityOutput

Remove-Item $liveParityOutput -ErrorAction SilentlyContinue
```

`QUERY_CANDIDATE_PLANNER_MODEL`이 이미 설정돼 있다면 그대로 사용합니다. 설정되지 않았다면 현재 프로젝트에서 실제 접근 가능한 모델 ID를 지정합니다.

```powershell
if ([string]::IsNullOrWhiteSpace($env:QUERY_CANDIDATE_PLANNER_MODEL)) {
  throw "QUERY_CANDIDATE_PLANNER_MODEL을 실제 접근 가능한 모델 ID로 설정하세요."
}
```

### 3. 실행

```powershell
node `
  .\tests\queryCandidatePatch13_3LiveProviderCacheHitParitySmokeTest.js

if ($LASTEXITCODE -ne 0) {
  throw "Patch 13.3 Live Provider cache parity 검증이 실패했습니다."
}
```

정상 콘솔:

```text
PASS query candidate patch13.3 live provider cache-hit parity smoke origin=CALLED replay=CACHE_HIT providerCalls=1 accepted=1이상 eligible=true
```

### 4. UTF-8 결과 확인

```powershell
$liveParityResult = Get-Content `
  $liveParityOutput `
  -Raw `
  -Encoding UTF8 |
  ConvertFrom-Json

[pscustomobject]@{
  OriginStatus                = $liveParityResult.origin.status
  OriginInvocation            = $liveParityResult.origin.invocationStatus
  OriginProviderCalls         = $liveParityResult.origin.providerCallCount
  OriginResponseId            = $liveParityResult.origin.responseId
  OriginTotalTokens           = $liveParityResult.origin.usage.totalTokens
  OriginAccepted              = $liveParityResult.origin.counts.accepted
  ReplayStatus                = $liveParityResult.replay.status
  ReplayInvocation            = $liveParityResult.replay.invocationStatus
  ReplayProviderCalls         = $liveParityResult.replay.providerCallCount
  ReplayPlannerCacheSource    = $liveParityResult.replay.cache.plannerResolution.source
  ReplayReentryCacheSource    = $liveParityResult.replay.cache.reentry.source
  ObservedProviderCalls       = $liveParityResult.parityAudit.observedProviderCallCount
  ParityValid                 = $liveParityResult.parityAudit.valid
  EncryptedPersistentFiles    = $liveParityResult.parityAudit.persistentFiles.encryptedFileCount
  PlaintextPersistentFiles    = $liveParityResult.parityAudit.persistentFiles.plaintextFileCount
  ReadinessEligible           = $liveParityResult.readinessGate.eligible
  ReadinessDecision           = $liveParityResult.readinessGate.decision
  ProductionPromotionAllowed = $liveParityResult.readinessGate.guardrails.productionPromotionAllowed
  ProductionRouteAutoWired    = $liveParityResult.readinessGate.guardrails.productionRouteAutoWired
} | Format-List
```

최종 성공 기준:

```text
OriginStatus                SHADOW_COMPLETED
OriginInvocation            CALLED
OriginProviderCalls         1
OriginResponseId            resp_...
OriginTotalTokens           1 이상
OriginAccepted              1 이상

ReplayStatus                SHADOW_COMPLETED
ReplayInvocation            CACHE_HIT
ReplayProviderCalls         0
ReplayPlannerCacheSource    L3_SEMANTIC
ReplayReentryCacheSource    L4_REENTRY
ObservedProviderCalls       1
ParityValid                 True
EncryptedPersistentFiles    3 이상
PlaintextPersistentFiles    0
ReadinessEligible           True
ReadinessDecision           ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW
ProductionPromotionAllowed False
ProductionRouteAutoWired    False
```

## Patch 13.2·13.1·13 회귀 검사

```powershell
node .\tests\queryCandidatePatch13_2TtlSweepAuditSmokeTest.js
node .\tests\queryCandidatePatch13_2UploadInvalidationSmokeTest.js
node .\tests\queryCandidatePatch13_2KeyRotationSmokeTest.js
node .\tests\queryCandidatePatch13_2ReplaySafeAuditSmokeTest.js
node .\tests\queryCandidatePatch13_2ShadowUploadDeletionBridgeSmokeTest.js
node .\tests\queryCandidatePatch13_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch13_1ShadowReentryCacheIntegrationSmokeTest.js
node .\tests\queryCandidatePatch13_1CorruptFallbackSmokeTest.js
node .\tests\queryCandidatePatch13_1PrivacyBoundarySmokeTest.js
node .\tests\queryCandidatePatch13_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13_1ManifestSmokeTest.js

node .\tests\queryCandidatePatch13HierarchicalCacheKeySmokeTest.js
node .\tests\queryCandidatePatch13EncryptedHierarchySmokeTest.js
node .\tests\queryCandidatePatch13CachePolicySmokeTest.js
node .\tests\queryCandidatePatch13PlannerAdapterSmokeTest.js
node .\tests\queryCandidatePatch13SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch13ManifestSmokeTest.js
```

## 전체 기준선 비교

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

## 환경변수 정리

```powershell
Remove-Item Env:QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY `
  -ErrorAction SilentlyContinue

Remove-Item Env:QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_OUTPUT `
  -ErrorAction SilentlyContinue

Remove-Item Env:QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_CACHE_DIR `
  -ErrorAction SilentlyContinue

Remove-Item Env:QUERY_CANDIDATE_PLANNER_LIVE_CACHE_PARITY_KEEP_CACHE `
  -ErrorAction SilentlyContinue
```

## 완료 판정

Live parity와 전체 회귀가 모두 통과하면 다음을 확인한 상태입니다.

```text
실제 Provider 최초 호출                  검증 완료
새 cache instance 영구 L3 hit            검증 완료
새 cache instance 영구 L4 hit            검증 완료
Provider 추가 호출 0회                   검증 완료
Proposal/Resolver/Family/Feasibility/Rank 동일
암호화 파일만 저장                       검증 완료
Production 격리                          검증 완료
Controlled promotion 검토 준비           적격
Production 자동 연결                     미수행
```
