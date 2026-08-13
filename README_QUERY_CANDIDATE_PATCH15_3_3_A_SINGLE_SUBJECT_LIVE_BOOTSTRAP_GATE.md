# Patch 15.3.3-A — Single-Subject Live Bootstrap Authorization Gate

This patch is the first substep of Patch 15.3.3. It installs an additive, default-off bootstrap authorization gate only. It does not wire or activate live traffic.

## Why this exists

The existing Patch 15.3 legacy preflight requires a valid real-shadow evidence bundle (`REAL_SHADOW_TRAFFIC`, actualTraffic=true) before a canary request is allowed. Patch 15.3.2-G explicitly forbids treating its pre-canary evaluation bundle as that legacy evidence. Therefore Patch 15.3.3 needs a narrow bridge to collect the first real internal canary evidence without fabricating legacy evidence.

## Authorization rules

Bootstrap authorization is possible only when all are true:

- bootstrap enabled explicitly;
- bootstrap kill switch explicitly OFF;
- legacy preflight is BLOCKED specifically with `READINESS_EVIDENCE_INVALID`;
- legacy evidence remains invalid and is never substituted;
- exact Patch G bundle payload SHA is provided and verifies;
- exact F.1.6 approval binding gate is intact;
- exact approved allowlist subject is used;
- F.1.6 approval binding ALLOWs;
- audience is ALLOWLIST and rollout percent is 0;
- Semantic Profiler only policy and all existing Feature Control / kill-switch checks pass.

## Defaults

- Bootstrap enabled: false
- Bootstrap kill switch: true
- Route change: none
- Railway mutation: none
- Provider calls: 0
- Actual internal user exposure: none
- Actual operational telemetry: false
- Percentage rollout: false
- Production promotion: false

Patch 15.3.3-B will perform explicit runtime integration and a first live internal request only after this patch passes locally.

## Local verification

Use the private Patch G bundle, the reissued F.1.5 approval receipt, and the privacy-safe subject SHA only. Do not pass a raw account ID.

```powershell
node `
  .\scripts\queryCandidatePlannerVerifyInternalCanaryLiveBootstrapReadiness.js `
  --g-bundle .\queryCandidatePlannerPatch15_3_2_G.private\queryCandidatePlannerFinalEvaluationEvidenceBundle.private.json `
  --approval-receipt .\queryCandidatePlannerPatch15_3_2_F.private\queryCandidatePlannerInternalCanaryManualApprovalReceipt.private.json `
  --subject-sha256 <approved-subject-sha256>
```

This verifier builds an isolated in-memory environment object. It does not mutate `process.env`, Railway, routes, or execute a Provider call.
