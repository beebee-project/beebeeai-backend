# Patch 15.3.2-F.1.4 — Internal Canary Evidence Candidate Bundle

## Purpose

Patch F.1.3 ended with:

```text
OPERATIONAL_DECISION EVALUATION_PASS
ASSESSMENT_DECISION ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS
FAILED_CHECK_COUNT 0
ABSOLUTE_COST_FAILURE_COUNT 0
CACHE_COST_AVOIDANCE_PASSED true
PRODUCTION_PROMOTION_AUTHORIZED false
```

F.1.4 does **not** turn that evaluation result into production authorization.

It creates a sanitized, deterministic evidence bundle that is only eligible for a later **manual Internal Allowlist Canary review**.

## Result boundary

Maximum result:

```text
REVIEW_DECISION
ELIGIBLE_FOR_INTERNAL_ALLOWLIST_CANARY_REVIEW

internalCanaryReviewEligible   true
manualOperatorApprovalRequired true

internalCanaryAuthorized       false
percentageRolloutAuthorized    false
productionPromotionAuthorized  false
productionMergeAuthorized      false
```

## Inputs

F.1.4 binds hashes for:

```text
APPROVED_ACTUAL pricing policy
Historical Patch 13.3 live parity evidence
F.1.2 actual-pricing-consistent canonical input
Canonical source threshold policy
F.1.3 private recalibrated threshold policy
F.1.3 recalibration evidence
F.1.3 operational report
F.1.3 assessment
F.1.3 final baseline
Current Cost/Cache/Latency evaluator
Current evaluator Git HEAD version
```

Expected evaluator:

```text
2461A48972A8F771E6D49911D70079009E62148658C17EEF986CA3E01972208D
```

Expected final F.1.3 baseline:

```text
0c59e08cead5a81d84abd4159aedd34d21666898d6d637c58aed7616ab62730f
```

## Sanitization

The candidate bundle includes aggregate evidence and integrity hashes only.

It does not include:

```text
responseId
raw execution rows
raw token-usage rows
immutable account IDs
allowlist subjects
environment values
```

## Important

The private recalibrated threshold policy is **not copied into the production configuration**.

The candidate bundle records only the approved candidate contract:

```text
averageCostMicrousdMax                  2600
providerCallAverageCostMicrousdMax      6500
monthlyProjectedCostMicrousdMax         26000000
cacheCostAvoidanceRateMin               0.59
providerCallRateMax                     0.40
warmAverageCostMicrousdMax              0
```

A later patch must separately create a manual approval receipt and bind its immutable bundle SHA before any Internal Allowlist Canary can be authorized.

## No runtime changes

F.1.4 does not modify:

```text
Canary gate
Promotion gate
Feature flags
Kill switch
Environment variables
Allowlist
Routes
Production merge adapter
Provider runtime
```

Provider calls executed by this patch:

```text
0
```
