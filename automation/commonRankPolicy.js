"use strict";

const COMMON_RANK_POLICY_VERSION =
  "common_rank_contract_v1";

const DEFAULT_RANK_TIE_POLICY =
  "include-all-ties-then-stable-source-order";

const DEFAULT_RANK_NUMBERING_POLICY =
  "competition";

const SUPPORTED_RANK_NUMBERING_POLICIES =
  Object.freeze([
    "competition",
    "dense",
    "ordinal",
  ]);

function normalizeRankDirection(
  direction = "desc",
) {
  return direction === "asc"
    ? "asc"
    : "desc";
}

function normalizeRankLimit(limit = 1) {
  const parsed = Number(limit);
  return Number.isFinite(parsed) && parsed > 0
    ? Math.max(1, Math.floor(parsed))
    : 1;
}

function normalizeRankNumberingPolicy(
  policy = DEFAULT_RANK_NUMBERING_POLICY,
) {
  return SUPPORTED_RANK_NUMBERING_POLICIES
    .includes(policy)
    ? policy
    : DEFAULT_RANK_NUMBERING_POLICY;
}

function numericRankValue(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed)
    ? parsed
    : null;
}

function rankValuesEqual(left, right) {
  const leftNumber = numericRankValue(left);
  const rightNumber = numericRankValue(right);

  if (
    leftNumber != null &&
    rightNumber != null
  ) {
    return Object.is(
      leftNumber,
      rightNumber,
    );
  }

  return String(left ?? "") ===
    String(right ?? "");
}

function stableRankEntries(
  entries = [],
  direction = "desc",
) {
  const normalizedDirection =
    normalizeRankDirection(direction);
  const multiplier =
    normalizedDirection === "asc"
      ? 1
      : -1;

  return (entries || [])
    .map((entry, originalIndex) => ({
      entry,
      originalIndex,
      numericValue:
        numericRankValue(entry?.value),
      sourceOrder:
        Number.isFinite(
          Number(entry?.sourceOrder),
        )
          ? Number(entry.sourceOrder)
          : originalIndex,
    }))
    .sort((left, right) => {
      const leftValid =
        left.numericValue != null;
      const rightValid =
        right.numericValue != null;

      if (leftValid && rightValid) {
        const delta =
          (
            left.numericValue -
            right.numericValue
          ) * multiplier;

        if (delta) return delta;
      } else if (leftValid !== rightValid) {
        return leftValid ? -1 : 1;
      } else {
        const textDelta = String(
          left.entry?.value ?? "",
        ).localeCompare(
          String(
            right.entry?.value ?? "",
          ),
        );

        if (textDelta) {
          return (
            normalizedDirection === "asc"
              ? textDelta
              : -textDelta
          );
        }
      }

      return (
        left.sourceOrder -
          right.sourceOrder ||
        left.originalIndex -
          right.originalIndex
      );
    })
    .map((item) => item.entry);
}

function includesBoundaryTies(
  tiePolicy = DEFAULT_RANK_TIE_POLICY,
) {
  return String(tiePolicy || "")
    .startsWith("include-all-ties");
}

function selectRankEntries({
  sortedEntries = [],
  limit = 1,
  tiePolicy = DEFAULT_RANK_TIE_POLICY,
} = {}) {
  const normalizedLimit =
    normalizeRankLimit(limit);

  if (
    sortedEntries.length <=
    normalizedLimit
  ) {
    return [...sortedEntries];
  }

  if (!includesBoundaryTies(tiePolicy)) {
    return sortedEntries.slice(
      0,
      normalizedLimit,
    );
  }

  const boundary =
    sortedEntries[
      normalizedLimit - 1
    ]?.value;

  return sortedEntries.filter(
    (entry, index) =>
      index < normalizedLimit ||
      rankValuesEqual(
        entry?.value,
        boundary,
      ),
  );
}

function assignRankNumbers({
  entries = [],
  numberingPolicy =
    DEFAULT_RANK_NUMBERING_POLICY,
} = {}) {
  const normalizedPolicy =
    normalizeRankNumberingPolicy(
      numberingPolicy,
    );

  let previousValue;
  let currentRank = 0;
  let denseRank = 0;

  return entries.map((entry, index) => {
    const sameAsPrevious =
      index > 0 &&
      rankValuesEqual(
        entry?.value,
        previousValue,
      );

    if (normalizedPolicy === "ordinal") {
      currentRank = index + 1;
    } else if (!sameAsPrevious) {
      if (
        normalizedPolicy === "dense"
      ) {
        denseRank += 1;
        currentRank = denseRank;
      } else {
        currentRank = index + 1;
      }
    }

    previousValue = entry?.value;

    return {
      ...entry,
      rank: currentRank,
    };
  });
}

function buildRankValue({
  entries = [],
  sourceMetricId = "",
  direction = "desc",
  limit = 1,
  tiePolicy =
    DEFAULT_RANK_TIE_POLICY,
  rankNumberingPolicy =
    DEFAULT_RANK_NUMBERING_POLICY,
} = {}) {
  const normalizedDirection =
    normalizeRankDirection(direction);
  const normalizedLimit =
    normalizeRankLimit(limit);
  const sortedEntries =
    stableRankEntries(
      entries,
      normalizedDirection,
    );
  const selectedEntries =
    selectRankEntries({
      sortedEntries,
      limit: normalizedLimit,
      tiePolicy,
    });
  const items = assignRankNumbers({
    entries: selectedEntries,
    numberingPolicy:
      rankNumberingPolicy,
  });

  return {
    valueType: "rank",
    sourceMetricId,
    direction: normalizedDirection,
    limit: normalizedLimit,
    items,
  };
}

module.exports = {
  COMMON_RANK_POLICY_VERSION,
  DEFAULT_RANK_TIE_POLICY,
  DEFAULT_RANK_NUMBERING_POLICY,
  SUPPORTED_RANK_NUMBERING_POLICIES,
  normalizeRankDirection,
  normalizeRankLimit,
  normalizeRankNumberingPolicy,
  numericRankValue,
  rankValuesEqual,
  stableRankEntries,
  includesBoundaryTies,
  selectRankEntries,
  assignRankNumbers,
  buildRankValue,
};
