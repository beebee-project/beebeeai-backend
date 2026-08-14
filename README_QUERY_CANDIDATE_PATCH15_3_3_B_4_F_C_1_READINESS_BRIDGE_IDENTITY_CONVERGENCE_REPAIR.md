# Patch 15.3.3-B-4-F-C.1 — Readiness Bridge Identity Convergence Repair

## Purpose

Patch B-4-F-A.2 V2 normalized the sanitized readiness JSON and therefore changed the readiness bridge SHA256 from `97C8F0CE...` to `77DB527F...`. The A Gate remained pinned to the former bridge SHA, so the staged activation pre-deploy gate correctly failed closed with `BOOTSTRAP_READINESS_BRIDGE_SHA_DRIFT`.

## Change

Exactly one production source line changes:

- `automation/queryCandidatePlannerInternalCanaryLiveBootstrapGate.js`
  - `EXPECTED_READINESS_BRIDGE_SHA256`: `97C8F0CE...` -> `77DB527F...`

No feature flags, routing, provider execution, merge execution, readiness content, approval evidence, F.1.6 gate, B-2 runtime, or Feature Control behavior changes.

## Required safety sequence

Before pushing this code, restore the staged Railway activation controls to the fail-closed values with `--skip-deploys`. This is required because a Git push can trigger Railway Auto Deploy and must not combine this code fix with an unverified activation configuration.

## Identities

- predecessor A Gate SHA256: `DFE04C089F0F514FA60026BE9FD3EF4EDA0DD584B4B55ECC6C2AF54FDECECD7D`
- repaired A Gate SHA256: `9386A73BAD4E37C055209AF59B86C5FFB21A62545E26017BFB5D3A109E4EB1D9`
- unchanged readiness bridge SHA256: `77DB527F808BBB61BD63BD61913E01A489AB25E154C5D4C0E67DAC730AB81259`
- unchanged readiness file SHA256: `46D1211AF4F318DAB91D137F0728C3AE6F246CD8B85A2582802CCB6DB1475AC4`

## Boundaries

- Railway mutation by patch application: false
- deploy by patch application: false
- provider calls by patch application: 0
- actual live request by patch application: false
- production route remains disabled until separately authorized
- rollout remains 0 until separately authorized
