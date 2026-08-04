# Patch 14.5 — API/UI E2E + Rollback Quality Gate

## 목적

Patch 14.5는 Query Candidate Planner의 API·내부 Preview·캐시 수명주기·Controlled Production Gate·Merge Adapter를 하나의 결정론적 통합 시나리오에서 검증한다.

이번 패치는 CI/QA 전용 품질 게이트를 추가한다. 기존 API Route, Controller, 일반 사용자 UI, Production 응답 경로에는 연결하지 않는다.

## 고정 안전 상태

- 기존 `/analysis-candidates` Primary 응답 권한 유지
- Shadow 오류와 Timeout이 HTTP 응답에 전파되지 않음
- Internal UI Preview는 메모리 기반 읽기 전용
- Preview 기록기 실패가 Primary 응답에 전파되지 않음
- 캐시 무효화 실패가 파일 삭제 API 성공을 막지 않음
- 다운로드는 캐시를 유지
- Promotion Gate 기본값 `BLOCKED`
- Merge Adapter는 허용 경로에서도 `dry-run`만 검증
- Global Kill Switch 활성화 직후 다음 판정부터 Shadow·Promotion·Merge 차단
- Production 후보 병합, READY 승격, Production Route 변경 없음
- 실제 Provider 호출 0회
- Railway 환경변수 변경 없음

## 통합 시나리오

1. 기존 Primary API 응답을 Shadow Boundary로 통과시킨다.
2. Primary payload·HTTP status·응답 형식이 유지되는지 확인한다.
3. Shadow 결과와 Primary 후보의 Comparator 결과를 확인한다.
4. Observation을 내부 Preview Store에 기록하고 읽기 전용 HTML 계약을 검사한다.
5. Promotion Gate의 기본 `BLOCKED` 상태를 확인한다.
6. 테스트 전용 SHA-256 Allowlist 조건에서 Gate `ALLOW`와 Merge Adapter `DRY_RUN_READY`를 확인한다.
7. Shadow 예외, Shadow Timeout, Preview Recorder 예외를 각각 주입한다.
8. Cache Invalidation 예외를 주입해 파일 API 장애 격리를 확인한다.
9. 다운로드 시 캐시 `RETAINED`를 확인한다.
10. Runtime Global Kill Switch를 활성화한다.
11. 다음 API 요청부터 Shadow가 `BLOCKED`인지 확인한다.
12. Promotion Gate와 Merge Adapter가 즉시 `BLOCKED`로 복귀하는지 확인한다.
13. Kill Switch를 반복 활성화해 fail-closed 상태가 유지되는지 확인한다.

## 적용

백엔드 루트에서 실행한다.

```powershell
Get-FileHash `
  .\query_candidate_patch14_5_api_ui_e2e_rollback_quality_gate.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch14_5_api_ui_e2e_rollback_quality_gate.zip `
  -DestinationPath . `
  -Force
```

이 ZIP은 저장소 루트 구조로 구성되어 있어 추가 최상위 폴더를 생성하지 않는다.

## 문법 검사

```powershell
node --check `
  .\automation\queryCandidatePlannerApiUiRollbackQualityGate.js

Get-ChildItem .\tests\queryCandidatePatch14_5*.js |
  ForEach-Object {
    node --check $_.FullName
  }
```

## Patch 14.5 QA

```powershell
Get-ChildItem .\tests\queryCandidatePatch14_5*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    node $_.FullName
  }
```

정상적으로 21개 PASS가 출력되어야 한다. Full Quality Gate는 내부 통합 체크 22개를 수행한다.

핵심 출력:

```text
PASS query candidate patch14.5 full quality gate smoke checks=22
PASS query candidate patch14.5 immediate rollback smoke
PASS query candidate patch14.5 production guardrails smoke
PASS query candidate patch14.5 manifest smoke
```

## 선행 패치 호환성

```powershell
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3_1CumulativeIntegrityConvergenceSmokeTest.js
node .\tests\queryCandidatePatch14_4ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_4_1ManifestSmokeTest.js
```

Patch 14.5는 기존 파일을 교체하지 않는 adds-only 패치이므로 선행 Manifest를 변경하지 않는다.

## 환경변수

이번 단계에서는 Railway 환경변수를 추가하거나 변경하지 않는다. Quality Gate는 테스트 내부에 격리된 Feature Control과 SHA-256 Allowlist 환경을 사용하며 `process.env`를 수정하지 않는다.

## 다음 단계

Patch 14.5 통과 이후 Patch 15.0에서 Accuracy Metric과 Evaluation Dataset을 정의한다. 실제 Promotion Gate 최초 활성화는 Patch 15.3 Internal Allowlist Canary까지 계속 금지한다.
