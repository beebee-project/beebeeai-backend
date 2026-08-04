# Patch 14.2.1 — API Shadow Boundary Manifest Drift Restore

## Purpose

Restore `automation/queryCandidatePlannerApiShadowBoundary.js` to the exact
Patch 14.1 canonical source after an undeclared local byte-level drift caused
`queryCandidatePatch14_1ManifestSmokeTest.js` to fail.

Patch 14.2 does not intentionally modify this boundary file. Therefore the
safe resolution is restoration, not adding the file to Patch 14.2's
supersession allowlist.

## Canonical file identity

- Path: `automation/queryCandidatePlannerApiShadowBoundary.js`
- Bytes: `3741`
- SHA-256: `2eff9cebc23d8695a3bb27a6007d4f7bc8419b4c6b3e41292f2a1f2f357f8aff`

## Apply

Run from the backend repository root:

```powershell
Copy-Item `
  .\automation\queryCandidatePlannerApiShadowBoundary.js `
  .\automation\queryCandidatePlannerApiShadowBoundary.js.patch14_2_1.bak `
  -Force

Expand-Archive `
  .\query_candidate_patch14_2_1_boundary_manifest_drift_restore.zip `
  -DestinationPath . `
  -Force
```

## Verify

```powershell
Get-FileHash `
  .\automation\queryCandidatePlannerApiShadowBoundary.js `
  -Algorithm SHA256

node .\tests\queryCandidatePatch14_2_1BoundaryRestoreSmokeTest.js
node .\tests\queryCandidatePatch14_1ManifestSmokeTest.js
node .\tests\queryCandidatePatch14_2ManifestSmokeTest.js
```

Expected final lines:

```text
PASS query candidate patch14.2.1 boundary manifest drift restore smoke
PASS query candidate patch14.1 manifest smoke superseded=7
PASS query candidate patch14.2 manifest smoke
```

No environment variable changes and no provider calls are required.
