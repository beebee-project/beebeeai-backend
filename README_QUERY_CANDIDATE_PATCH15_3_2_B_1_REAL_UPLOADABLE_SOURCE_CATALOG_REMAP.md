# Patch 15.3.2-B.1 — Real Uploadable Source Catalog Remap & Invalid Ledger Guard

## Decision

Patch 15.3.2-B의 기존 10개 `caseId` 중 일부는 label-only seed 또는 물리 원본 미지정 Case여서 실제 플랫폼 업로드 요청을 만들 수 없었다. 이 패치는 가짜 Fixture를 생성하지 않는다. 대신 10개 Case 각각을 실제 업로드 가능한 `REAL_ANONYMIZED` 또는 `PUBLIC_DATASET` 파일에 명시적으로 결합하고, 결합이 완료되기 전에는 Ledger 생성·기록·최종화를 모두 차단한다.

기존 v1 Ledger는 source artifact와 암호학적으로 결합돼 있지 않으므로 자동 승계하지 않는다. 현재 표시상 4/10 Ledger는 백업 후 무효화되고, 최종 Source Catalog를 기준으로 v2 Ledger를 0/10에서 다시 시작한다. `hardcase_two_tables_one_sheet_waste`, `real_world_event_applicant_workshop`, `template_course_evaluation_report`도 Catalog에 실제 파일을 결합한 뒤 다시 실행한다.

## Safety state

- Evidence Collector: OFF 유지
- Internal Canary: OFF 유지
- Production Merge: OFF 유지
- Production Promotion: BLOCKED 유지
- Railway 변수 자동 변경: 없음
- 실제 원본 파일: 패치 ZIP에 포함하지 않음
- Synthetic/Generated Fixture: 실제 Evidence Source로 금지

## Added/updated components

- `automation/queryCandidatePlannerRealShadowUploadableSourceCatalog.js`
- `automation/queryCandidatePlannerRealShadowRegistryFinalization.js` v2
- Source Catalog scaffold/bind/progress/finalize CLI
- Legacy Ledger invalidation CLI
- Source-bound Ledger scaffold/record/progress/finalize CLI
- strict exact-64 fingerprint validation
- request/upload 동일 fingerprint 차단
- source file SHA-256 binding mismatch 차단
- private catalog/ledger Git staging guard

## Source policy

허용:

- `REAL_ANONYMIZED`: 실제 업무 파일에서 개인정보·민감정보를 제거한 파일
- `PUBLIC_DATASET`: 재사용 가능한 공개 데이터 원본

금지:

- `SEED`
- `SYNTHETIC`
- `GENERATED_FIXTURE`
- 임의로 만든 테스트 CSV/XLSX
- Case label과 의미가 다른 파일
- 동일 원본을 두 Case에 중복 결합

## Apply

```powershell
$ErrorActionPreference = "Stop"

Get-FileHash `
  .\query_candidate_patch15_3_2_B_1_real_uploadable_source_catalog_remap.zip `
  -Algorithm SHA256

Expand-Archive `
  .\query_candidate_patch15_3_2_B_1_real_uploadable_source_catalog_remap.zip `
  -DestinationPath . `
  -Force
```

## Syntax check

```powershell
node --check `
  .\automation\queryCandidatePlannerRealShadowUploadableSourceCatalog.js

node --check `
  .\automation\queryCandidatePlannerRealShadowRegistryFinalization.js

Get-ChildItem `
  .\scripts\queryCandidatePlanner*RealShadow*.js |
  ForEach-Object {
    node --check $_.FullName
    if ($LASTEXITCODE -ne 0) {
      throw "Syntax check failed: $($_.Name)"
    }
  }
```

## Patch QA

```powershell
Get-ChildItem `
  .\tests\queryCandidatePatch15_3_2_B_1*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object {
    Write-Host "RUN $($_.Name)"
    node $_.FullName
    if ($LASTEXITCODE -ne 0) {
      throw "Patch 15.3.2-B.1 test failed: $($_.Name)"
    }
  }
```

Expected: `17/17 PASS`.

## Predecessor convergence

```powershell
$tests = @(
  ".\tests\queryCandidatePatch15_3_2_BSourceIntegritySmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2_BManifestSmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2_ASourceIntegritySmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2_AManifestSmokeTest.js",
  ".\tests\queryCandidatePatch15_3SourceIntegritySmokeTest.js",
  ".\tests\queryCandidatePatch15_3ManifestSmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2SourceIntegritySmokeTest.js",
  ".\tests\queryCandidatePatch15_3_2ManifestSmokeTest.js",
  ".\tests\queryCandidatePatch15_3_1PredecessorIntegrityRepairSmokeTest.js"
)

foreach ($test in $tests) {
  node $test
  if ($LASTEXITCODE -ne 0) {
    throw "Predecessor integrity failed: $test"
  }
}
```

## Code commit before private workflow

Patch source and tests should be committed before creating private catalog outputs. `tests` is ignored in the current repository, so use `git add -f` for B.1 tests.

## 1. Scaffold private Source Catalog

```powershell
$sourceCatalogDraft = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowUploadableSourceCatalog.draft.private.json"

node `
  .\scripts\queryCandidatePlannerScaffoldRealShadowUploadableSourceCatalog.js `
  --catalog-id "internal_real_shadow_uploadable_sources_2026_08_v1" `
  --output $sourceCatalogDraft
```

This starts at `0/10`. No source file is fabricated.

## 2. Bind each real uploadable source

Example:

```powershell
node `
  .\scripts\queryCandidatePlannerBindRealShadowUploadableSource.js `
  --catalog $sourceCatalogDraft `
  --case-id "template_course_evaluation_report" `
  --source-file "C:\absolute\path\course_evaluation_report.csv" `
  --source-kind "PUBLIC_DATASET" `
  --confirm-semantic-compatibility true `
  --verified-at "$([DateTime]::UtcNow.ToString('o'))"
```

The script computes and stores file SHA-256, size, extension, and absolute private path. It never logs workbook rows.

### Ten required semantic bindings

1. `hardcase_two_tables_one_sheet_waste`: budget execution, multiple tables, irrelevant rows
2. `real_world_event_applicant_workshop`: event application and attendance
3. `seed_attendance_conditional`: replace seed with a real/public attendance status file
4. `seed_sales_ready`: replace seed with a real/public transaction sales file
5. `template_course_evaluation_report`: `course_evaluation_report.csv` when locally present and verified
6. `seed_unstructured_unsupported`: replace seed with a real/public unsupported or unstructured spreadsheet
7. `ambiguous_mixed_columns_review`: ambiguous mixed-column real/public file requiring review
8. `inventory_stock_movement`: stock movement and reorder threshold file
9. `project_task_tracker`: project task, owner, deadline, status file
10. `expense_claim_review`: expense claim, policy status, approval file

Case IDs remain aligned with Patch 15.0 labels. The remap is the binding from each label Case to a real uploadable source artifact, not a silent change of labels.

## 3. Source Catalog progress

```powershell
node `
  .\scripts\queryCandidatePlannerShowRealShadowUploadableSourceProgress.js `
  --catalog $sourceCatalogDraft
```

Completion criteria:

```text
PROGRESS 10/10
REMAINING 0
COMPLETE true
```

## 4. Finalize Source Catalog

```powershell
$sourceCatalog = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowUploadableSourceCatalog.private.json"

node `
  .\scripts\queryCandidatePlannerFinalizeRealShadowUploadableSourceCatalog.js `
  --catalog $sourceCatalogDraft `
  --output $sourceCatalog `
  --summary-output ".\queryCandidatePlannerRealShadowUploadableSourceCatalog.summary.private.json"
```

Finalization is blocked unless all 10 files still exist and their current SHA-256/size/extension match the catalog.

## 5. Invalidate current v1 Ledger

The existing 4/10 Ledger is not source-bound. Back it up and create v2 at 0/10:

```powershell
$legacyLedger = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowFingerprintLedger.private.json"

$ledgerV2 = Join-Path `
  (Get-Location) `
  "queryCandidatePlannerRealShadowFingerprintLedger.v2.private.json"

node `
  .\scripts\queryCandidatePlannerInvalidateLegacyRealShadowLedger.js `
  --legacy-ledger $legacyLedger `
  --source-catalog $sourceCatalog `
  --registry-id "internal_real_shadow_2026_08_v2" `
  --output $ledgerV2

$ledger = $ledgerV2
```

Expected:

```text
LEGACY_CAPTURES_PRESERVED false
RERUN_ALL_CASES_REQUIRED true
PROGRESS 0/10
```

If no old Ledger remains, scaffold v2 directly:

```powershell
node `
  .\scripts\queryCandidatePlannerScaffoldRealShadowFingerprintLedger.js `
  --source-catalog $sourceCatalog `
  --registry-id "internal_real_shadow_2026_08_v2" `
  --output $ledgerV2
```

## 6. Rerun and record each Case

After uploading the exact catalog-bound file and obtaining fresh Preview fingerprints:

```powershell
node `
  .\scripts\queryCandidatePlannerRecordRealShadowFingerprint.js `
  --ledger $ledger `
  --source-catalog $sourceCatalog `
  --source-file "C:\absolute\path\the_same_catalog_bound_file.csv" `
  --case-id "template_course_evaluation_report" `
  --request-fingerprint $requestFingerprint `
  --upload-fingerprint $uploadFingerprint `
  --capture-source "INTERNAL_PREVIEW" `
  --captured-at "$([DateTime]::UtcNow.ToString('o'))"
```

The source file is rehashed at record time. A different file is rejected with `REAL_SHADOW_CAPTURE_SOURCE_ARTIFACT_MISMATCH`.

Fingerprints must be exactly 64 hexadecimal characters. B.1 no longer truncates suffixes. Request and Upload fingerprints must differ.

## 7. Progress and finalization

```powershell
node `
  .\scripts\queryCandidatePlannerShowRealShadowRegistryProgress.js `
  --ledger $ledger `
  --source-catalog $sourceCatalog
```

After `10/10`:

```powershell
node `
  .\scripts\queryCandidatePlannerFinalizeRealShadowCaseRegistry.js `
  --ledger $ledger `
  --source-catalog $sourceCatalog `
  --output ".\queryCandidatePlannerRealShadowCaseRegistry.private.json" `
  --railway-output ".\queryCandidatePlannerRealShadowCaseRegistry.railway.private.txt" `
  --summary-output ".\queryCandidatePlannerRealShadowCaseRegistry.summary.private.json"
```

## Private file exclusion

Add these patterns to `.git/info/exclude`:

```text
queryCandidatePlannerRealShadowUploadableSourceCatalog*.private.json
queryCandidatePlannerRealShadowFingerprintLedger*.private.json
queryCandidatePlannerRealShadowCaseRegistry*.private.json
queryCandidatePlannerRealShadowCaseRegistry*.private.txt
```

Then run:

```powershell
node `
  .\scripts\queryCandidatePlannerAssertRealShadowPrivateOutputsUntracked.js
```

## Completion criteria

- Uploadable Source Catalog: 10/10
- Source kinds: only REAL_ANONYMIZED or PUBLIC_DATASET
- Duplicate source artifacts: 0
- Synthetic/generated sources: 0
- Legacy v1 Ledger accepted: false
- Source-bound v2 Ledger: 10/10 after rerun
- Exact-64 fingerprint enforcement: PASS
- Request/Upload identical fingerprint: blocked
- Source mismatch at record time: blocked
- Collector/Internal Canary/Production Merge: still OFF
