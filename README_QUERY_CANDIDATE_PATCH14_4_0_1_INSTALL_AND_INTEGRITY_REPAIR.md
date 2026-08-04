# Patch 14.4.0.1 — Install Packaging + Cumulative Integrity Repair

## Purpose

This hotfix corrects two independent issues observed while applying Patch 14.4:

1. The original Patch 14.4 ZIP contained an extra top-level directory, so expanding it into the repository root created a nested folder instead of installing `automation/` and `tests/` at the repository root.
2. Active Patch 14.1–14.3 files in the working tree had drifted from their manifests. Patch 14.4 itself was not installed and did not cause those source changes.

## Safety Contract

- Restores active Patch 14.1 and 14.2 protected files through the already validated Patch 14.3.1 convergence contract.
- Restores the complete Patch 14.3 payload and manifest.
- Installs Patch 14.4 files into the correct repository-root paths.
- Creates backups only for files whose bytes differ.
- Does not change API routes or controllers beyond restoring the accepted Patch 14.3 route state.
- Does not wire the production merge adapter into any API path.
- Does not call an external provider.
- Keeps the production merge adapter default OFF.

## Apply

Extract this ZIP into a temporary directory, not directly over the repository:

```powershell
Expand-Archive `
  .\query_candidate_patch14_4_0_1_install_and_integrity_repair.zip `
  -DestinationPath .\_patch14_4_0_1 `
  -Force

node `
  .\_patch14_4_0_1\applyPatch14_4_0_1InstallAndIntegrityRepair.js `
  .
```

Before applying, add these lines to `.gitignore` so temporary folders and safety backups are not committed:

```gitignore
_patch*/
.patch_backups/
query_candidate_patch*.zip
```

The incorrectly extracted directory from the original Patch 14.4 package can be removed:

```powershell
Remove-Item `
  .\query_candidate_patch14_4_controlled_production_merge_adapter `
  -Recurse `
  -Force `
  -ErrorAction SilentlyContinue
```

## Verification

```powershell
node .\tests\queryCandidatePatch14_4_0_1InstallRepairSmokeTest.js

node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_3_1CumulativeIntegrityConvergenceSmokeTest.js

node --check `
  .\automation\queryCandidatePlannerControlledProductionMergeAdapter.js

Get-ChildItem .\tests\queryCandidatePatch14_4*SmokeTest.js |
  Sort-Object Name |
  ForEach-Object { node $_.FullName }
```

Expected key results:

```text
PASS query candidate patch14.4.0.1 install and integrity repair smoke
PASS query candidate patch14.1 manifest smoke superseded=7
PASS query candidate patch14.2 manifest smoke superseded=2
PASS query candidate patch14.3 manifest smoke
PASS query candidate patch14.3.1 cumulative integrity convergence smoke ...
PASS query candidate patch14.4 manifest smoke
```

## Commit Discipline

The convergence files must be explicitly staged. Do not run `git commit` before checking the staged file list.

```powershell
git add -A
git diff --cached --name-status
git status
```

The staged list should include deletion of the previously committed `.patch_backups/...` files, restored active planner files where necessary, Patch 14.4 adapter/test files, and `.gitignore`.
