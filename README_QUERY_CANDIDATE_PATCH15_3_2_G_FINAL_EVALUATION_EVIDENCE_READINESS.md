# Patch 15.3.2-G — Final Evaluation Evidence Bundle & Internal Canary Readiness

## Purpose

Finalize Patch 15.3.2-F evidence into one immutable, private readiness bundle for the next internal-canary bootstrap step.

G does **not** claim that F's canonical benchmark is real current operational traffic. The existing Patch 15.3 evidence contract requires `source=REAL_SHADOW_TRAFFIC`, `actualTraffic=true`, and `synthetic=false`; G does not forge or replace that contract.

## Decision

A valid bundle may emit:

`READY_FOR_15_3_3_INTERNAL_ALLOWLIST_CANARY_BOOTSTRAP`

This means evaluation evidence and approval/code bindings are ready for Patch 15.3.3 to perform an explicit, allowlist-only bootstrap integration. It does not itself activate runtime traffic.

## Immutable bindings

- Final F baseline SHA-256
- Evaluator SHA-256
- F.1.4 candidate payload/file SHA-256
- F.1.7 rotation-plan file SHA-256
- Rotated allowlist SHA-256
- Reissued F.1.5 approval receipt payload/file SHA-256
- F.1.6 approval binding Gate source SHA-256
- F.1.6.1 composed Service source SHA-256

## Safety boundary

- Legacy Patch 15.3 real-shadow evidence contract satisfied by G: **false**
- Evidence substitution: **forbidden**
- Provider calls: **0**
- Actual operational telemetry: **false**
- Railway/environment/route mutation: **none**
- Runtime auto-activation: **false**
- Internal user actual exposure authorized by G: **false**
- Percentage rollout: **false**
- Production promotion: **false**

Patch 15.3.3 must explicitly consume the G bootstrap-readiness bundle, preserve allowlist-only/0%/kill-switch/fallback constraints, and collect the first actual operational telemetry. Patch 15.3.4 must require actual traffic evidence before any 15.4 entry review.
