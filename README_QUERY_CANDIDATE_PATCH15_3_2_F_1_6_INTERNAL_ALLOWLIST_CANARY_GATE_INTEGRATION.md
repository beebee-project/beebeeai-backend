# Patch 15.3.2-F.1.6 — Internal Allowlist Canary Gate Integration

## Purpose

F.1.6 is the first runtime binding step for the approved Internal Allowlist Canary.

It makes the F.1.5 manual approval receipt mandatory before the existing
Internal Allowlist Canary service can return an ALLOW preflight.

The patch does **not** authorize general rollout.

## Immutable binding chain

Runtime ALLOW requires all three independent identities to match:

```text
F.1.4 CANDIDATE_PAYLOAD_SHA256
F.1.5 APPROVAL_RECEIPT_PAYLOAD_SHA256
QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256
```

The fixed F.1.4 candidate payload is:

```text
928F6A6E0AA8683D63A5A2CB62199FA460EB84494B119EB7E171000843D484EA
```

F.1.5 receipt contents are supplied in:

```text
QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_RECEIPT_JSON
```

The immutable F.1.5 receipt payload SHA is supplied separately in:

```text
QUERY_CANDIDATE_PLANNER_CANARY_APPROVAL_BUNDLE_SHA256
```

The Gate recomputes the receipt canonical payload SHA and requires exact equality.

## Existing runtime architecture preserved

Patch 15.3 already routes the analysis-candidates boundary through the Internal
Allowlist Canary service. F.1.6 does not modify routes.

F.1.6 modifies the service through an explicit integration script so its
preflight is superseded by the new manual-approval binding gate.

A backup is created under:

```text
.patch_backups/query_candidate_patch15_3_2_F_1_6_<timestamp>
```

## Runtime ALLOW requirements

All must be true:

```text
Internal Canary Enabled = 1
Internal Canary Kill Switch = 0
Global Kill Switch = 0

Feature Enabled = 1
Shadow Enabled = 1
Provider Enabled = 1
Provider Kill Switch = 0

Production Enabled = 1
Production Candidate Merge Enabled = 1
Production Kill Switch = 0

Production READY Assignment Enabled = 0
Production Route Enabled = 0

Promotion Gate Enabled = 1
Promotion Audience = ALLOWLIST
Promotion Rollout Percent = 0

Internal Canary LLM Mode = SEMANTIC_PROFILER_ONLY

Receipt candidate SHA matches F.1.4
Receipt payload SHA matches approval bundle env SHA
Receipt allowlist SHA is present in runtime allowlist
Request subject SHA equals approved receipt allowlist SHA
Request subject SHA is present in runtime allowlist

Feature Control allows:
SHADOW_EXECUTION
PROVIDER_CALL
PRODUCTION_CANDIDATE_MERGE
```

Any mismatch is fail-closed.

## Evidence semantics

This patch intentionally does **not** claim current operational telemetry.

The approved evidence remains:

```text
canonical benchmark
+
approved actual pricing
+
historical live provider cache parity
+
manual operator approval
```

Therefore:

```text
ACTUAL_OPERATIONAL_TELEMETRY false
CANARY_EVIDENCE_COLLECTION_REQUIRED true
```

The internal Canary run itself is what produces the next real operational
evidence for consideration before Patch 15.4.

## Authorization boundary

A successful Gate can authorize:

```text
Internal Allowlist Canary runtime for the exact approved subject
```

It still does not authorize:

```text
1% or higher percentage rollout
general-user rollout
Production READY assignment
Production route change
broad Production Promotion
```

## Provider calls

Patch installation, integration, and offline verification execute zero Provider calls.

Actual Provider calls occur only later when the user deliberately enables and
executes the Internal Canary runtime under the existing semantic-profiler-only
policy.
