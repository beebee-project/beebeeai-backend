const paymentService = require("./paymentService");

function nowDate(value) {
  const d = value ? new Date(value) : new Date();
  return Number.isNaN(d.getTime()) ? new Date() : d;
}

function normalizeStatus(status = "") {
  return String(status || "INACTIVE").toUpperCase();
}

function hasOwn(obj, key) {
  return Object.prototype.hasOwnProperty.call(obj || {}, key);
}

function ensureUsageShape(user) {
  if (!user) return false;

  let changed = false;
  const current = user.usage || {};

  if (!user.usage) {
    user.usage = {};
    changed = true;
  }

  if (typeof current.templateGenerations !== "number") {
    user.usage.templateGenerations =
      typeof current.formulaConversions === "number"
        ? current.formulaConversions
        : 0;
    changed = true;
  }

  if (typeof current.formulaConversions !== "number") {
    user.usage.formulaConversions = user.usage.templateGenerations || 0;
    changed = true;
  }

  if (typeof current.fileUploads !== "number") {
    user.usage.fileUploads = 0;
    changed = true;
  }

  if (!current.lastReset) {
    user.usage.lastReset = new Date();
    changed = true;
  }

  return changed;
}

function needMonthlyReset(lastReset, now = new Date()) {
  if (!lastReset) return true;
  const base = new Date(lastReset);
  if (Number.isNaN(base.getTime())) return true;

  return (
    base.getUTCFullYear() !== now.getUTCFullYear() ||
    base.getUTCMonth() !== now.getUTCMonth()
  );
}

function resetUsageCounters(user, now = new Date()) {
  if (!user) return false;
  if (!user.usage) user.usage = {};

  user.usage.templateGenerations = 0;
  user.usage.formulaConversions = 0;
  user.usage.fileUploads = 0;
  user.usage.lastReset = now;
  return true;
}

function applyMonthlyUsageReset(user, options = {}) {
  const now = nowDate(options.now);
  let changed = ensureUsageShape(user);

  if (needMonthlyReset(user?.usage?.lastReset, now)) {
    resetUsageCounters(user, now);
    changed = true;
  }

  return { changed, reason: changed ? "MONTHLY_USAGE_RESET_OR_SHAPE" : "NOOP" };
}

function hasRealSubscriptionSignal(user = {}) {
  const sub = user.subscription || {};
  const status = normalizeStatus(sub.status);

  return Boolean(
    sub.billingKey ||
    sub.customerKey ||
    sub.startedAt ||
    sub.trialEndsAt ||
    sub.nextChargeAt ||
    ["ACTIVE", "PAST_DUE", "CANCELED_PENDING"].includes(status),
  );
}

function isExpiredCanceledPending(user = {}, now = new Date()) {
  const sub = user.subscription || {};
  const status = normalizeStatus(sub.status);
  if (status !== "CANCELED_PENDING" || !sub.cancelAtPeriodEnd) return false;
  if (!sub.nextChargeAt) return false;

  const endAt = new Date(sub.nextChargeAt);
  if (Number.isNaN(endAt.getTime())) return false;
  return endAt <= now;
}

function getBetaMode(options = {}) {
  return hasOwn(options, "betaMode")
    ? Boolean(options.betaMode)
    : paymentService.isBetaMode();
}

function hasBetaModeOffResetMarker(user = {}) {
  return Boolean(user?.usage?.betaModeOffResetAt);
}

/**
 * BETA_MODE=true에서는 DB plan이 FREE인 사용자도 paymentService/getEffectivePlan
 * 흐름에서 PRO처럼 사용할 수 있다. 따라서 true -> false 전환 후 리셋 조건을
 * plan=PRO에만 걸면 실제 베타 사용량이 남는다.
 *
 * 이 전환 리셋은 사용자별 1회만 수행한다.
 * - 실제 구독 신호가 있으면 리셋하지 않는다.
 * - 이미 betaModeOffResetAt이 있으면 다시 리셋하지 않는다.
 * - 최초 실서비스 모드 진입 시 marker를 저장해 이후 FREE 사용량은 정상 누적된다.
 */
function isBetaModeTurnedOffUsageResetTarget(user = {}, options = {}) {
  if (getBetaMode(options)) return false;
  if (hasRealSubscriptionSignal(user)) return false;
  if (hasBetaModeOffResetMarker(user)) return false;
  return true;
}

function isBetaModeTurnedOffOrphanPro(user = {}, options = {}) {
  if (getBetaMode(options)) return false;
  if (normalizeStatus(user.plan) !== "PRO") return false;
  return !hasRealSubscriptionSignal(user);
}

function resetPlanToFreeWithoutDeletingFiles(
  user,
  subscriptionPatch = {},
  now = new Date(),
) {
  user.plan = "FREE";
  resetUsageCounters(user, now);
  user.subscription = {
    ...(user.subscription || {}),
    ...subscriptionPatch,
  };
  return true;
}

function applyBetaModeOffUsageReset(user, options = {}) {
  if (!isBetaModeTurnedOffUsageResetTarget(user, options)) {
    return { changed: false, reason: "NOOP" };
  }

  const now = nowDate(options.now);
  resetPlanToFreeWithoutDeletingFiles(
    user,
    {
      status: "INACTIVE",
      cancelAtPeriodEnd: false,
      nextChargeAt: null,
      endedAt: now,
    },
    now,
  );

  user.usage.betaModeOffResetAt = now;
  return { changed: true, reason: "BETA_MODE_OFF_USAGE_RESET" };
}

function applyUsageStateTransitions(user, options = {}) {
  if (!user) return { changed: false, reason: "NO_USER" };

  const now = nowDate(options.now);
  ensureUsageShape(user);

  // 1) 구독 해지 예약 상태에서 이용기간이 끝난 경우: FREE 전환 + 사용량 리셋
  if (isExpiredCanceledPending(user, now)) {
    resetPlanToFreeWithoutDeletingFiles(
      user,
      {
        status: "CANCELED",
        cancelAtPeriodEnd: false,
        endedAt: now,
        nextChargeAt: null,
      },
      now,
    );

    user.usage.subscriptionPeriodEndedResetAt = now;
    return { changed: true, reason: "SUBSCRIPTION_PERIOD_ENDED" };
  }

  // 2) BETA_MODE=true -> false 전환 후 실제 구독 신호가 없는 사용자 1회 리셋
  const betaReset = applyBetaModeOffUsageReset(user, { ...options, now });
  if (betaReset.changed) return betaReset;

  return { changed: false, reason: "NOOP" };
}

function applyBetaUsageResetAfterRealMode(user, options = {}) {
  const result = applyBetaModeOffUsageReset(user, options);
  return Boolean(result.changed);
}

function applyUsageResetPolicies(user, options = {}) {
  const now = nowDate(options.now);

  const monthly = applyMonthlyUsageReset(user, { now });
  const transition = applyUsageStateTransitions(user, { ...options, now });

  return {
    changed: Boolean(monthly.changed || transition.changed),
    reasons: [monthly.reason, transition.reason].filter(
      (reason) => reason && reason !== "NOOP",
    ),
  };
}

module.exports = {
  applyBetaModeOffUsageReset,
  applyBetaUsageResetAfterRealMode,
  applyMonthlyUsageReset,
  applyUsageResetPolicies,
  applyUsageStateTransitions,
  ensureUsageShape,
  hasRealSubscriptionSignal,
  isBetaModeTurnedOffOrphanPro,
  isBetaModeTurnedOffUsageResetTarget,
  isExpiredCanceledPending,
  needMonthlyReset,
  resetUsageCounters,
};
