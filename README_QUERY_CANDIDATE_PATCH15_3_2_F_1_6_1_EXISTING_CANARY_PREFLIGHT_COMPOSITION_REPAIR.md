# Patch 15.3.2-F.1.6.1 — Existing Canary Preflight Composition Repair

## Purpose

Patch F.1.6 correctly introduced the manual approval binding gate, but its first
service integration used:

```text
return approvalBindingGate.preflight;
```

immediately after subject derivation.

That meant an approval ALLOW replaced the existing Patch 15.3 preflight rather
than composing with it.

F.1.6.1 repairs that integration.

## Required starting state

The Internal Allowlist Canary service must already be restored to:

```text
089E260D90625E068769F3D3538FAC198B4EAB3CEC4D864EF8CA9A747123E561
```

The F.1.6 approval binding gate must remain:

```text
ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7
```

Any drift blocks the repair.

## Correct composition

```text
derive immutable subject
        ↓
F.1.6 approval binding gate
        │
        ├─ BLOCK → return blocked preflight immediately
        │
        └─ ALLOW → DO NOT RETURN
                    ↓
              existing Patch 15.3 preflight
                    ↓
              config valid?
                    ↓
              canary enabled?
                    ↓
              internal kill switch off?
                    ↓
              semantic-profiler-only?
                    ↓
              subject complete?
                    ↓
              legacy canary evidence valid?
                    ↓
              audience = ALLOWLIST?
                    ↓
              rollout = 0?
                    ↓
              existing Promotion Gate
                    ↓
              existing Feature Control
                    ↓
              ALLOWLIST_PREFLIGHT_ALLOWED
```

Therefore final runtime authorization is an AND composition:

```text
F.1.6 approval binding ALLOW
AND
existing Patch 15.3 preflight ALLOW
```

## Safety

F.1.6.1:

- does not modify routes;
- does not modify controllers;
- does not set environment variables;
- does not call a provider;
- keeps rollout percent at 0;
- does not authorize Production READY;
- does not authorize Production route changes;
- does not authorize broad Production Promotion.

The apply script backs up the service under `.patch_backups`.

## Obsolete F.1.6 integration

The old early-return integration marker is explicitly forbidden. If it is still
present, F.1.6.1 fails closed and requires restoring the service first.
