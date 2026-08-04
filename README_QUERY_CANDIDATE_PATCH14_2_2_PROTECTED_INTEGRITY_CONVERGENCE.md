# Patch 14.2.2 — Protected Integrity Convergence

## 목적

Patch 14.2 적용 후 Patch 14.1 Manifest가 보고한 비의도적 드리프트를 한 파일씩 우회하지 않고, Patch 14.2가 대체하지 않은 Patch 14.1 보호 파일 전체를 정식 기준으로 복원합니다.

확인된 드리프트:

- `automation/queryCandidatePlannerApiShadowBoundary.js`
- `automation/queryCandidatePlannerFeatureControlRuntime.js`

Patch 14.2가 정식으로 대체한 7개 파일은 복원 대상에서 제외됩니다.

## 안전 계약

- Patch 14.2 후속 파일 7개는 덮어쓰지 않습니다.
- 현재 파일이 정식 해시와 다를 때만 백업 후 복원합니다.
- 기존 파일은 `.patch_backups/query_candidate_patch14_2_2_<timestamp>`에 보존합니다.
- 환경변수, Provider, Production route, Primary 응답을 변경하지 않습니다.

## 적용

ZIP을 백엔드 루트의 임시 폴더에 압축 해제합니다.

```powershell
Expand-Archive `
  .\query_candidate_patch14_2_2_protected_integrity_convergence.zip `
  -DestinationPath .\_patch14_2_2 `
  -Force
```

복원기를 실행합니다.

```powershell
node `
  .\_patch14_2_2\applyPatch14_2_2ProtectedIntegrityConvergence.js `
  .
```

## 검증

```powershell
node .\tests\queryCandidatePatch14_2_2ProtectedIntegrityConvergenceSmokeTest.js
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
```

정상 출력:

```text
PASS query candidate patch14.2.2 protected integrity convergence smoke protected=14 superseded=7
PASS query candidate patch14.1 manifest smoke superseded=7
PASS query candidate patch14.2 source integrity smoke
PASS query candidate patch14.2 manifest smoke
```
