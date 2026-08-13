# Query Candidate Planner Patch 15.3.3-B-2

## Single-Subject Live Bootstrap Runtime Wiring

This patch wires the existing Patch 15.3.3-A live-bootstrap authorization gate
into the already-live Internal Allowlist Canary boundary.

The patch is deliberately **observe-only** for the bootstrap path:

- the existing Primary HTTP payload remains authoritative;
- bootstrap execution never applies the Controlled Production Merge Adapter;
- no Production READY assignment is made;
- no Production route is enabled;
- no percentage rollout is authorized;
- legacy Patch 15.3 real-shadow evidence is not substituted;
- the existing real-shadow collector bridge is not used by the bootstrap runner;
- Patch 15.3.3-B-2 itself does not change Railway variables and performs no
  Provider call.

## Exact predecessor

Apply only on:

```text
branch main
HEAD d4e3588b70a44f71880f6353230f6b05211357e5
working tree clean
```

Protected predecessor identities:

```text
Internal Allowlist Canary Service
1A61F219ADF49BD863B84C5B8C4DB02158E901E7EDA864AC551656A4A7E75C8F

F.1.6 Approval Binding Gate
ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7

15.3.2-G Final Evaluation Evidence module
439F29AC82D866EEADA3EDFBD8615892904ACD507E4F8D4D5161431E0449440A

15.3.3-A Live Bootstrap Gate
4585B4549B0F756274F47FBB9089E56A07D21C6EFE3C1929214E856B068B5498

15.3.3-A Local Bootstrap Verifier
92B2085B188F3AEFAFC0D84008E3CF8542E0F9CB504F67B53505740719747C8C

automationRoutes.js
2D5390681F3A4306EBE1BE6166FBE9CC875A71C5A94CCDAABE824511EBC4B626
```

## Runtime flow

When `QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_BOOTSTRAP_ENABLED` is absent or
explicitly `0`, the existing boundary behavior remains unchanged.

When it is explicitly `1`, the bootstrap path takes exclusive control of the
new-canary branch:

```text
Primary controller result
        |
legacy Internal Canary preflight
        |
15.3.3-A single-subject bootstrap authorization
        |
        +-- BLOCK -> Primary response only; no bootstrap shadow execution
        |
        +-- ALLOW -> Primary response returned immediately
                       |
                       +-- asynchronous semantic-profiler-only canary shadow
                           max Provider calls = 1
                           observe-only no-merge adapter
                           Primary response remains unchanged
```

While bootstrap mode is active, a legacy preflight that unexpectedly becomes
ALLOW is **not** allowed to fall through to the legacy canary merge path. The
bootstrap authorization gate blocks with
`BOOTSTRAP_NOT_REQUIRED_LEGACY_PREFLIGHT_ALREADY_ALLOWED`, and the boundary
returns Primary only. This prevents an unexpected evidence-state change from
turning the bootstrap stage into a Production merge.

## Collector isolation

`routes/automationRoutes.js` is not modified.

The route currently supplies a real-shadow capture `shadowRunner` and an
`onObservation` callback for the legacy path. Patch B-2 intentionally does not
forward either of those into the bootstrap runner.

The bootstrap runtime calls the canonical
`runQueryCandidatePlannerInternalAllowlistCanary` with its default canonical
API shadow runner and an injected no-merge adapter. Therefore Patch B-2 does
not re-enable or relabel the excluded legacy Patch 15.3.2-E evidence contract.

A dedicated Patch 15.3.3-C operational evidence bundle will be built later
from actual internal-canary telemetry.

## Files

Repository-visible patch files:

```text
PATCH_MANIFEST_PATCH15_3_3_B_2.json
PATCH_VALIDATION_PATCH15_3_3_B_2.json
README_QUERY_CANDIDATE_PATCH15_3_3_B_2_LIVE_BOOTSTRAP_RUNTIME_WIRING.md
automation/queryCandidatePlannerInternalAllowlistCanaryBoundary.js
automation/queryCandidatePlannerInternalCanaryLiveBootstrapRuntime.js
evaluation/queryCandidatePlannerInternalCanaryLiveBootstrapRuntimePolicy.v1.json
scripts/queryCandidatePlannerVerifyInternalCanaryLiveBootstrapRuntimeWiring.js
```

Smoke tests under `tests/` are included in the package but the repository's
current `.gitignore` ignores the whole `tests/` directory, matching the
existing Patch 15.3.2-G / 15.3.3-A local-test policy.

## Apply

Do not push before Patch B-3. GitHub `main` is connected to Railway production
with Auto Deploy enabled.

From the backend repository root, first verify the ZIP SHA supplied with the
package. Then verify the predecessor and extract the ZIP into the repository
root.

After extraction, do **not** change Railway variables yet.

## Local verification

Run:

```powershell
node .\scripts\queryCandidatePlannerVerifyInternalCanaryLiveBootstrapRuntimeWiring.js

Get-ChildItem `
  .\tests\queryCandidatePatch15_3_3_B_2*SmokeTest.js |
Sort-Object Name |
ForEach-Object {
  Write-Host "RUN $($_.Name)"
  node $_.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Patch 15.3.3-B-2 smoke failed: $($_.Name)"
  }
}
```

All tests are local/mock/provider-free.

## B-2 does not authorize activation

After B-2 verification:

```text
Railway variables modified                 false
Git push executed                          false
Provider calls executed by patch           0
Actual internal-user exposure              false
Actual operational telemetry               false
Percentage rollout                         false
Production promotion                       false
```

Next: Patch 15.3.3-B-3 prepares and verifies the exact Railway environment
bindings while the bootstrap kill switch remains active.
