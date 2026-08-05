# Patch 15.3.1 — Predecessor Integrity Repair

Patch 15.3 신규 Canary 기능은 정상이며, 이 패치는 선행 Manifest 검사에서 다시 불일치한 네 보호 파일만 검증된 정식 바이트로 복원합니다.

복원 대상:

- `automation/queryCandidatePlannerControlledProductionPromotionGate.js`
- `automation/queryCandidatePlannerApiUiRollbackQualityGate.js`
- `automation/queryCandidatePlannerCostCacheLatencyEvaluator.js`
- `automation/queryCandidatePlannerShadowAccuracyEvaluator.js`

이 패치는 Route, Controller, Canary 설정, 환경변수, Production Merge 활성화, Provider 호출을 변경하지 않습니다.

## 적용

```powershell
Get-FileHash .\query_candidate_patch15_3_1_predecessor_integrity_repair.zip -Algorithm SHA256
Expand-Archive .\query_candidate_patch15_3_1_predecessor_integrity_repair.zip -DestinationPath . -Force
node .\tests\queryCandidatePatch15_3_1PredecessorIntegrityRepairSmokeTest.js
```

## 누적 회귀 검사

```powershell
node .\tests\queryCandidatePatch14_4_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_4_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch14_5ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5FullQualityGateSmokeTest.js
node .\tests\queryCandidatePatch15_0ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch15_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch15_2ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_3SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch15_3ManifestSmokeTest.js
```
