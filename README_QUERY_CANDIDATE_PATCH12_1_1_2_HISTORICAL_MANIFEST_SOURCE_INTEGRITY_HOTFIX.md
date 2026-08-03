# Patch 12.1.1.2 — Historical Manifest Supersession + Source Integrity Robustness Hotfix

적용 순서:

```text
Patch 12 → Patch 12.1 → Patch 12.1.1 → Patch 12.1.1.1 → Patch 12.1.1.2
```

## 수정 대상

사용자 저장소에서 기능 기준선은 모두 통과했지만 다음 과거 QA 테스트 2개가 실패했다.

```text
queryCandidatePatch12_1_1SourceIntegritySmokeTest.js
→ 일반 structural recipe 목록 정규식 탐색 실패

queryCandidatePatch12_1_1ManifestSmokeTest.js
→ automation/queryCandidateResolver.js byte size mismatch
```

Resolver·Planner·Shadow Bridge·fixture·기준선은 변경하지 않는다.

## 원인

### Source Integrity

기존 테스트는 Resolver 상수 선언의 공백과 정확한 구문 형태를 정규식으로 추출했다. 누적 패치 이후 선언 구조가 바뀌어도 실제 격리 동작은 정상이지만 문자열 정규식만 실패할 수 있었다.

패치 후에는 공개 함수인 다음 두 함수를 호출해 동작을 직접 검증한다.

```text
parsedRecipeOperandSpec
isStructuralGenericCandidate
```

검증 경계:

```text
일반 time_count       → NOT_APPLICABLE / structural false
Planner time_count    → REQUIRED / period operand / structural true
일반 count_rows       → NOT_APPLICABLE
Planner count_rows    → REQUIRED / operand 없음
```

### Historical Manifest

`PATCH_MANIFEST_PATCH12_1_1.json`은 패치 12.1.1 배포 당시의 Resolver 해시를 기록한 역사적 manifest다. 이후 누적 패치가 같은 파일을 수정하면 과거 manifest의 Resolver exact hash가 현재 파일과 달라지는 것은 정상이다.

패치 후 정책:

```text
변경되지 않은 과거 산출물 → 기존 bytes/SHA exact 검증
후속 패치가 대체한 Resolver → 존재·비어 있지 않음·SHA 형식 검증
과거 manifest 기록값       → 수정하지 않고 보존
```

허용되는 superseded path는 후속 12.1.1.1과 이번 패치가 실제로 대체한 다음 4개로 고정한다.

```text
automation/queryCandidateResolver.js
tests/queryCandidatePatch12_1_1SourceIntegritySmokeTest.js
tests/queryCandidatePatch12_1_1ManifestSmokeTest.js
README_QUERY_CANDIDATE_PATCH12_1_1_PLANNER_REENTRY_RESOLVER_ISOLATION_HOTFIX.md
```

## 적용

```powershell
Expand-Archive `
  .\query_candidate_patch12_1_1_2_historical_manifest_source_integrity_hotfix.zip `
  -DestinationPath . `
  -Force
```

## 문법 검사

```powershell
node --check .\tests\queryCandidatePatch12_1_1SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch12_1_1ManifestSmokeTest.js
node --check .\tests\queryCandidatePatch12_1_1_2SourceIntegritySmokeTest.js
node --check .\tests\queryCandidatePatch12_1_1_2ManifestSmokeTest.js
```

## 오류 재검증

```powershell
node .\tests\queryCandidatePlannerResolverLegacyTimeCountCompatibilitySmokeTest.js
node .\tests\queryCandidatePatch12_1_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch12_1_1_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_1_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch12_1_1_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch12_1_1_2ManifestSmokeTest.js
```

정상 출력:

```text
PASS query candidate planner resolver legacy time_count compatibility smoke
PASS query candidate patch12.1.1 source integrity smoke
PASS query candidate patch12.1.1 manifest smoke superseded=4
PASS query candidate patch12.1.1.1 source integrity smoke
PASS query candidate patch12.1.1.1 manifest smoke superseded=2
PASS query candidate patch12.1.1.2 source integrity smoke
PASS query candidate patch12.1.1.2 manifest smoke
```

Resolver가 과거 manifest와 아직 동일한 환경에서는 테스트·README 3개만 대체되므로 `superseded=3`도 정상이다. 사용자 누적 저장소처럼 Resolver도 후속 변경됐다면 `superseded=4`가 정상이다.

## Shadow 회귀

```powershell
node .\tests\queryCandidatePlannerShadowCapture.js `
  --mode=compare

node .\tests\queryCandidatePlannerShadowBaselineSmokeTest.js
node .\tests\queryCandidatePlannerShadowSampleAudit.js
```

정상 조건:

```text
PASS 1/1
accepted 2
resolved 2
ready 2
ranked 2
status SHADOW_COMPLETED
productionCandidateMerge false
productionRouteChanged false
```

## 전체 기준선 비교

```powershell
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

모두 다음 결과여야 한다.

```text
PASS 6/6
differences=0
```

기준선 `--mode=write`는 실행하지 않는다.
