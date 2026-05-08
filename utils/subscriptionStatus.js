function getSubscriptionExpiresAt(user) {
  const raw = user?.subscription?.nextChargeAt;
  if (!raw) return null;

  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return null;

  return date;
}

function isSubscriptionActive(user, now = new Date()) {
  const sub = user?.subscription;
  if (!sub) return false;

  const status = String(sub.status || "").toUpperCase();

  const activeLikeStatuses = new Set([
    "ACTIVE",
    "CANCELED_PENDING",
    "PAST_DUE",
  ]);

  if (!activeLikeStatuses.has(status)) return false;

  const expiresAt = getSubscriptionExpiresAt(user);
  if (!expiresAt) return false;

  return expiresAt > now;
}

function getEffectivePlan(user, now = new Date()) {
  return isSubscriptionActive(user, now) ? "PRO" : "FREE";
}

module.exports = {
  getSubscriptionExpiresAt,
  isSubscriptionActive,
  getEffectivePlan,
};
