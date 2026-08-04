# Patch 14.3.1 — Cumulative Predecessor Integrity Convergence

## 목적

Patch 14.3 Internal UI Preview 신규 QA는 정상 통과했지만, 누적 회귀 검사에서 다음 두 활성 선행 파일의 로컬 드리프트가 발견된 경우를 복구한다.

- `automation/queryCandidatePlannerApiShadowBoundary.js` — Patch 14.1 활성 보호 파일
- `automation/queryCandidatePlannerApiShadowRunner.js` — Patch 14.2 활성 보호 파일

이 패치는 최초 오류 파일 하나만 복원하지 않는다. Patch 14.3이 대체하지 않은 Patch 14.1/14.2 활성 파일 전체를 한 번에 검사하고, 정식 Manifest의 byte length와 SHA-256이 다른 파일만 백업 후 복원한다.

## 판정

- Patch 14.3 신규 기능 QA 15/15 PASS는 유효하다.
- 실패는 Internal UI Preview 기능 오류가 아니라 누적 선행 파일 정합성 오류다.
- Patch 14.3 ZIP에는 위 Boundary와 Runner가 포함되어 있지 않으므로, 해당 ZIP의 정상 압축 해제만으로 두 파일이 교체되는 구조는 아니다.
- 정확한 로컬 드리프트 발생 원인은 제공된 실행 로그만으로 확정할 수 없다. 편집기 자동 저장·포매터·수동 변경·별도 작업 트리 반영 여부는 백업 파일과 Git diff로 확인할 수 있다.

## 보호 범위

- Patch 14.1 활성 보호 파일: 14개
- Patch 14.2 활성 보호 파일: 28개
- 총 보호 파일: 42개
- 계약 Manifest: Patch 14.1, 14.2, 14.3 총 3개

Patch 14.3이 정식 대체한 다음 파일은 복원 대상에서 제외한다.

- `routes/automationRoutes.js`
- `tests/queryCandidatePatch14_2ManifestSmokeTest.js`

따라서 Internal UI Preview 코드와 Patch 14.3 successor-compatible Manifest 검사는 유지된다.

## 적용

백엔드 루트에서 ZIP을 임시 폴더에 푼다.

```powershell
Get-FileHash `
  .\query_candidate_patch14_3_1_cumulative_integrity_convergence.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch14_3_1_cumulative_integrity_convergence.zip `
  -DestinationPath .\_patch14_3_1 `
  -Force
```

복원기를 실행한다.

```powershell
node `
  .\_patch14_3_1\applyPatch14_3_1CumulativeIntegrityConvergence.js `
  .
```

변경된 기존 파일은 다음 위치에 자동 백업된다.

```text
.patch_backups\query_candidate_patch14_3_1_<timestamp>\
```

## 예상 적용 출력

현재 보고된 상태에서는 최소 Boundary와 Runner가 복원되어 다음과 유사하게 출력될 수 있다.

```text
PASS query candidate patch14.3.1 cumulative integrity convergence apply protected=42 manifests=3 restored=2 created=0 ...
```

`restored`가 2보다 크더라도 오류가 아니다. 다른 활성 선행 파일에도 숨은 드리프트가 있어 함께 복원됐다는 뜻이다.

## 최종 검사

```powershell
node .\tests\queryCandidatePatch14_3_1CumulativeIntegrityConvergenceSmokeTest.js

node .\tests\queryCandidatePatch14_1RouteContractSmokeTest.js
node .\tests\queryCandidatePatch14_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js

node .\tests\queryCandidatePatch14_2RouteContractSmokeTest.js
node .\tests\queryCandidatePatch14_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_2SchemaSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js

node .\tests\queryCandidatePatch14_2_1BoundaryRestoreSmokeTest.js
node .\tests\queryCandidatePatch14_2_2ProtectedIntegrityConvergenceSmokeTest.js

Get-ChildItem .\tests\queryCandidatePatch14_3*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object { node $_.FullName }
```

정상 핵심 출력은 다음과 같다.

```text
PASS query candidate patch14.3.1 cumulative integrity convergence smoke patch14.1Protected=14 patch14.2Protected=28 total=42
PASS query candidate patch14.1 manifest smoke superseded=7
PASS query candidate patch14.2 manifest smoke superseded=2
PASS query candidate patch14.2.1 boundary manifest drift restore smoke
PASS query candidate patch14.2.2 protected integrity convergence smoke protected=14 superseded=7
```

Patch 14.3 신규 테스트는 다시 15개 모두 PASS해야 한다.

## 운영 영향

- Feature Flag 변경 없음
- Railway 환경변수 변경 없음
- Provider 호출 없음
- Primary API 응답 변경 없음
- Production 후보 병합 없음
- READY 승격 없음
- Patch 14.3 내부 Preview 코드 덮어쓰기 없음

검사 완료 후 `_patch14_3_1`은 삭제할 수 있다. `.patch_backups`는 정상 상태를 Git에 커밋한 후 삭제하는 것이 안전하다.
