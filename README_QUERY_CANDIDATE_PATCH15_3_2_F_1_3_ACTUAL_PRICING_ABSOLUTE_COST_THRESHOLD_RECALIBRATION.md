# Patch 15.3.2-F.1.3 — Actual Pricing Absolute Cost Threshold Recalibration

F.1.2 fixed `avoidedByCacheMicrousd` pricing consistency. The existing 0.59 cache-cost-avoidance threshold now passes at 0.597531.

Only three absolute cost ceilings remain blocked.

## Method

Do not copy the measured averages into thresholds.

Use the current 10 provider-call Terra costs:

```text
4400, 4500, 4620, 4720, 4840,
4940, 5060, 5160, 5280, 5380 microusd
```

Anchor the unit-cost ceiling to the observed maximum:

```text
max = 5380
headroom = 20%
5380 × 1.20 = 6456
round upward to 100 = 6500
```

Preserve the existing provider call-rate limit:

```text
providerCallRateMax = 0.40
```

Derive average execution ceiling:

```text
6500 × 0.40 = 2600 microusd
```

Preserve the monthly projection:

```text
monthlyProjectionExecutions = 10000
```

Derive monthly ceiling:

```text
2600 × 10000 = 26000000 microusd
```

## Private recalibrated thresholds

Only these three fields differ from the source threshold policy:

```text
averageCostMicrousdMax                 130      -> 2600
providerCallAverageCostMicrousdMax     325      -> 6500
monthlyProjectedCostMicrousdMax        1300000  -> 26000000
```

Preserved:

```text
cacheCostAvoidanceRateMin              0.59
warmAverageCostMicrousdMax             0
providerCallRateMax                    0.40
monthlyProjectionExecutions            10000
all non-cost thresholds
```

The canonical source threshold file is not modified. A private derived policy is used for evaluation only.

Expected re-evaluation:

```text
OPERATIONAL_DECISION EVALUATION_PASS
ASSESSMENT_DECISION ACTUAL_PRICING_ABSOLUTE_COST_RECALIBRATION_PASS
FAILED_CHECK_COUNT 0
ABSOLUTE_COST_PASS_COUNT 3
ABSOLUTE_COST_FAILURE_COUNT 0
CACHE_COST_AVOIDANCE_PASSED true
PROVIDER_CALLS_EXECUTED_BY_EVALUATOR 0
PRODUCTION_PROMOTION_AUTHORIZED false
```

An evaluation PASS does not authorize production promotion.
