# Priority 4A — Actual Flow Evidence 기반 Mixed Section Restore Gate

## 목적

`재고·입출고 흐름 요약`과 같은 혼합 Section을 복원하거나 유지하기 전에 원본에서 생성된 `normalizedQueryTables`에 실제 재고 흐름 증거가 있는지 검사합니다.

허용 증거는 다음 두 가지입니다.

1. 실제 숫자값이 있는 `입고수량` 계열 열과 `출고수량` 계열 열이 모두 존재
2. 실제 방향값이 있는 입출고·이동 구분 열과 숫자 수량 열이 함께 존재

다음만 존재하는 경우에는 흐름 증거로 인정하지 않습니다.

- 사용수량
- 대여수량
- 일반 처리수량
- 방향 열 없이 존재하는 수량 열
- 방향값이 `사용`, `대여` 등인 경우

## Production 변경 파일

- `automation/semanticOutputPlanner.js`

## Production 비변경 항목

- 원본 CSV/XLSX
- 저장된 queryTables
- `metricSemanticRoleEngine.js`
- `businessTemplateExecutor.js`
- 업무별 Builder
- Success7 baseline
- Reference workbook

## 버전

- `semantic_output_planner_common_v2_8_actual_flow_evidence_restore_gate`
- `actual_flow_evidence_restore_gate_v1`

## 판정 예시

- 재고 보유: `입고수량 + 출고수량` → PASS, 혼합 흐름 요약 유지
- 창고 이동: `이동구분 + 이동수량`, 값 `입고/출고/이동` → PASS
- 소모품 사용: `사용수량 + 잔여수량` → FAIL, 혼합 흐름 요약 제거
- 장비 대여: `대여수량 + 대여상태` → FAIL, 혼합 흐름 요약 제거

특정 case ID, 파일명, fixture 숫자를 Production 분기에 사용하지 않습니다.

## 적용

```powershell
Copy-Item `
  .\automation\semanticOutputPlanner.js `
  .\automation\semanticOutputPlanner.js.pre-actual-flow-restore-gate.bak `
  -Force

Expand-Archive `
  .\priority4a_actual_flow_evidence_restore_gate_patch.zip `
  -DestinationPath . `
  -Force
```

## 해시 확인

```powershell
Get-FileHash `
  .\automation\semanticOutputPlanner.js `
  -Algorithm SHA256
```

기대 SHA-256:

```text
8AB0E6BEFBE60710014622FE3579656F82B5A8C82A8C3F5818AA626091BAE615
```

## 문법 및 신규 스모크

```powershell
node --check .\automation\semanticOutputPlanner.js

node .\tests\actualFlowEvidenceRestoreGateCommonSmokeTest.js
node .\tests\actualFlowEvidenceMixedSectionIntegrationSmokeTest.js
node .\tests\actualFlowEvidenceRestoreGateSourceIntegritySmokeTest.js
node .\tests\priority4AActualFlowEvidenceOriginalFixtureSmokeTest.js
```

기대 결과:

```text
PASS actual flow evidence restore gate common smoke
PASS actual flow evidence mixed section integration smoke
PASS actual flow evidence restore gate source integrity
PASS priority4A actual flow evidence original fixture smoke: 7
```

## 기존 호환 스모크

```powershell
node .\tests\actualMixedSectionRowShapeSnapshotAliasSmokeTest.js
node .\tests\actualMixedSectionRowShapeSourceIntegritySmokeTest.js
node .\tests\generalInventorySnapshotAliasCommonSmokeTest.js
node .\tests\inventorySnapshotAliasAggregationBridgeIntegrationSmokeTest.js
node .\tests\inventorySnapshotAliasAggregationBridgeSourceIntegritySmokeTest.js
node .\tests\aggregationAwareContractKpiBridgeCommonSmokeTest.js
node .\tests\mixedSectionContractSnapshotSourceIntegritySmokeTest.js
node .\tests\contractKpiSnapshotBridgeCommonSmokeTest.js
node .\tests\mixedSectionRowContractPrecedenceCommonSmokeTest.js
node .\tests\mixedSectionContractSnapshotIntegrationSmokeTest.js
node .\tests\metricSemanticRoleEngineCommonSmokeTest.js
node .\tests\semanticOutputPlannerMetricRoleAggregationSmokeTest.js
node .\tests\semanticContractPrecedenceCommonSmokeTest.js
node .\tests\semanticContractPrecedenceIntegrationSmokeTest.js
node .\tests\semanticContractPrecedenceSourceIntegritySmokeTest.js
node .\tests\semanticOutputPlannerModuleIntegritySmokeTest.js
node .\tests\semanticOutputPlannerCommonSmokeTest.js
node .\tests\semanticOutputPlannerGenericSectionCleanupSmokeTest.js
node .\tests\priority4AMetricSemanticSourceIntegritySmokeTest.js
node .\tests\manifestMetricIdsCommonSmokeTest.js
node .\tests\success7ReviewClosureCommonSmokeTest.js
node .\tests\statusRateCommonEngineSmokeTest.js
```

## 실제 7개 회귀

서버를 완전히 재시작한 뒤 실행합니다.

```powershell
$caseIds = Get-Content `
  .\tests\priority4a_actual_flow_evidence_restore_gate_cases.txt |
  Where-Object {
    $_.Trim() -and
    -not $_.Trim().StartsWith("#")
  } |
  ForEach-Object {
    "id:$($_.Trim())"
  }

$env:REGRESSION_CASE_FILTER = $caseIds -join ","
$env:REGRESSION_FAIL_FAST = "1"

Remove-Item Env:\REGRESSION_REENTRY_PROBE -ErrorAction SilentlyContinue
Remove-Item Env:\REGRESSION_PROBE_STOP_AFTER -ErrorAction SilentlyContinue
Remove-Item Env:\REGRESSION_BASELINE_MODE -ErrorAction SilentlyContinue

node .\tests\automationRegressionTests.js
```

기대 게이트:

```text
selectedCases: 7
passed: 7/7
artifact audit: 7/7
full artifact gate: OK
```

## 생성본 확인 기준

### 유지되어야 하는 혼합 흐름 요약

- `template_inventory_stock_status`
- `template_warehouse_movement_report`

### 없어야 하는 혼합 흐름 요약

- `template_supply_usage_report`
- `template_equipment_rental_status`

### 재고 핵심값 보존

```text
현재재고 최신 스냅샷 = 211
재고금액 최신 스냅샷 = 3,229,000
총 재고수량 = 211
평균 재고수량 = 70.33333333333333
```

위 숫자는 테스트 fixture 기대값이며 Production 코드에 포함되지 않습니다.

## 검증 범위

격리 환경에서 문법 검사, 신규 4개 테스트, 기존 Planner 호환 테스트를 실행했습니다. 실제 서버 재시작, 실제 7개 회귀, 전체 백엔드 회귀는 이 패키지 생성 환경에서 실행하지 않았습니다.
