const paymentService = require("./paymentService");

function nowDate(value) {
  const d = value ? new Date(value) : new Date();
  return Number.isNaN(d.getTime()) ? new Date() : d;
}

function normalizeStatus(status = "") {
  return String(status || "INACTIVE").toUpperCase();
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

function isBetaModeTurnedOffOrphanPro(user = {}) {
  if (paymentService.isBetaMode()) return false;
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
    return { changed: true, reason: "SUBSCRIPTION_PERIOD_ENDED" };
  }

  // 2) BETA_MODE=true 때 PRO였던 유저가 BETA_MODE=false에서 실제 구독 신호가 없는 경우
  if (isBetaModeTurnedOffOrphanPro(user)) {
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
    return { changed: true, reason: "BETA_MODE_OFF_ORPHAN_PRO" };
  }

  return { changed: false, reason: "NOOP" };
}

function applyUsageResetPolicies(user, options = {}) {
  const now = nowDate(options.now);

  const monthly = applyMonthlyUsageReset(user, { now });
  const transition = applyUsageStateTransitions(user, { now });

  return {
    changed: Boolean(monthly.changed || transition.changed),
    reasons: [monthly.reason, transition.reason].filter(
      (reason) => reason && reason !== "NOOP",
    ),
  };
}

function applyBetaUsageResetAfterRealMode(user, options = {}) {
  const betaMode =
    typeof options.betaMode === "boolean"
      ? options.betaMode
      : paymentService.isBetaMode();

  if (betaMode) return false;

  const now = nowDate(options.now);
  const changed = Boolean(options.force) || isBetaModeTurnedOffOrphanPro(user);

  if (!changed) return false;

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

  return true;
}

module.exports = {
  applyBetaUsageResetAfterRealMode,
  applyMonthlyUsageReset,
  applyUsageResetPolicies,
  applyUsageStateTransitions,
  ensureUsageShape,
  hasRealSubscriptionSignal,
  isBetaModeTurnedOffOrphanPro,
  isExpiredCanceledPending,
  needMonthlyReset,
  resetUsageCounters,
};
