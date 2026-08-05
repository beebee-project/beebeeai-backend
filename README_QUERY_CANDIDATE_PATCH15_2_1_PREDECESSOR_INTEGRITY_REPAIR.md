# Patch 15.2.1 — Predecessor Integrity Repair

Patch 15.2 기능은 정상이며, 이 패치는 누락된 원본 ZIP 때문에 복원되지 못한 세 보호 파일만 정식 바이트로 복원합니다.

복원 대상:

- `automation/queryCandidatePlannerControlledProductionPromotionGate.js`
- `automation/queryCandidatePlannerApiUiRollbackQualityGate.js`
- `automation/queryCandidatePlannerCostCacheLatencyEvaluator.js`

이 패치는 Route, Controller, Production Merge, Provider 호출을 변경하지 않습니다.

## 적용

```powershell
Get-FileHash .\query_candidate_patch15_2_1_predecessor_integrity_repair.zip -Algorithm SHA256
Expand-Archive .\query_candidate_patch15_2_1_predecessor_integrity_repair.zip -DestinationPath . -Force
node .\tests\queryCandidatePatch15_2_1PredecessorIntegrityRepairSmokeTest.js
```

## 회귀 검사

```powershell
node .\tests\queryCandidatePatch14_4_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_5FullQualityGateSmokeTest.js
node .\tests\queryCandidatePatch15_0ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_1SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch15_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch15_2SourceIntegritySmokeTest.js
node .\tests\queryCandidatePatch15_2ManifestSmokeTest.js
```
