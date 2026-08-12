# Patch 15.3.2-F.1.5 — Manual Approval Receipt & Immutable Evidence/Allowlist Hash Binding

F.1.5 binds three things without activating runtime:

```text
F.1.4 candidate payload SHA
F.1.4 candidate physical file SHA
Current allowlist SHA
```

Required candidate payload:

```text
928F6A6E0AA8683D63A5A2CB62199FA460EB84494B119EB7E171000843D484EA
```

Allowlist environment variable name recorded as metadata:

```text
QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256
```

The raw immutable account ID is never stored in the receipt.

Receipt creation requires explicit:

```text
--approve true
```

Successful F.1.5 means:

```text
internalCanaryApprovalGranted true
runtimeGateBindingApplied false
runtimeCanaryAuthorized false
percentageRolloutAuthorized false
productionPromotionAuthorized false
```

The receipt outputs a new:

```text
APPROVAL_RECEIPT_PAYLOAD_SHA256
```

F.1.6 must later bind all three independent hashes:

```text
CANDIDATE_PAYLOAD_SHA256
APPROVAL_RECEIPT_PAYLOAD_SHA256
ALLOWLIST_SHA256
```

F.1.5 does not modify the Gate, environment, feature flags, kill switch,
allowlist, routes, production merge adapter, or provider runtime.
