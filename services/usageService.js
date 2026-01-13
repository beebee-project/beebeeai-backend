const User = require("../models/User");
const paymentService = require("./paymentService");

function hasSubscriptionSignal(sub = {}) {
  return Boolean(
    sub?.billingKey ||
      sub?.customerKey ||
      sub?.startedAt ||
      sub?.trialEndsAt ||
      sub?.nextChargeAt
  );
}

function isSubscriptionActiveStrict(sub = {}) {
  const status = String(sub.status || "NONE").toUpperCase();
  return (
    hasSubscriptionSignal(sub) &&
    ["ACTIVE", "PAST_DUE", "CANCELED_PENDING"].includes(status)
  );
}

function getEffectivePlanFromUser(user) {
  const base = paymentService.getEffectivePlan(user.plan || "FREE");
  // ✅ 베타모드일 때만 plan(PRO)을 그대로 인정
  if (paymentService.isBetaMode() && base === "PRO") return "PRO";
  // ✅ 베타모드가 아니면 "구독 시그널 + status"일 때만 PRO
  if (isSubscriptionActiveStrict(user.subscription || {})) return "PRO";

  return "FREE";
}

const LIMITS = {
  FREE: { formulaConversions: 10, fileUploads: 1 },
  PRO: { formulaConversions: null, fileUploads: null },
};

function getLimits(plan = "FREE") {
  return LIMITS[plan] || LIMITS.FREE;
}

function needMonthlyReset(lastReset, now = new Date()) {
  if (!lastReset) return true;
  return (
    lastReset.getUTCFullYear() !== now.getUTCFullYear() ||
    lastReset.getUTCMonth() !== now.getUTCMonth()
  );
}

async function getUsageSummary(userId) {
  const user = await User.findById(userId).select("plan usage subscription");
  if (!user) throw new Error("User not found");

  let changed = false;
  if (!user.usage) {
    user.usage = {
      formulaConversions: 0,
      fileUploads: 0,
      lastReset: new Date(),
    };
    changed = true;
  }
  if (needMonthlyReset(user.usage.lastReset)) {
    user.usage.formulaConversions = 0;
    user.usage.fileUploads = 0;
    user.usage.lastReset = new Date();
    changed = true;
  }
  if (changed) await user.save();

  const plan = getEffectivePlanFromUser(user);
  return {
    plan,
    usage: {
      formulaConversions: user.usage.formulaConversions,
      fileUploads: user.usage.fileUploads,
    },
    limits: getLimits(plan),
  };
}

async function bumpUsage(userId, field, delta) {
  const user = await User.findById(userId).select(
    "plan usage subscription isDeleted"
  );
  if (!user) throw new Error("User not found");

  if (user.isDeleted) return { skipped: true };
  if (!user.usage)
    user.usage = {
      formulaConversions: 0,
      fileUploads: 0,
      lastReset: new Date(),
    };
  if (needMonthlyReset(user.usage.lastReset)) {
    user.usage.formulaConversions = 0;
    user.usage.fileUploads = 0;
    user.usage.lastReset = new Date();
  }
  user.usage[field] = Math.max(0, (user.usage[field] || 0) + delta);
  await user.save();
  const plan = getEffectivePlanFromUser(user);
  return { usage: user.usage, limits: getLimits(plan), plan };
}

async function assertCanUse(userId, field, amount = 1) {
  const user = await User.findById(userId).select(
    "plan subscription usage isDeleted purgeAt"
  );
  if (!user) throw new Error("User not found");

  if (user.isDeleted) {
    const err = new Error("ACCOUNT_DELETED");
    err.status = 403;
    err.code = "ACCOUNT_DELETED";
    throw err;
  }

  // 월 리셋 동일 로직
  if (!user.usage) {
    user.usage = {
      formulaConversions: 0,
      fileUploads: 0,
      lastReset: new Date(),
    };
  }
  if (needMonthlyReset(user.usage.lastReset)) {
    user.usage.formulaConversions = 0;
    user.usage.fileUploads = 0;
    user.usage.lastReset = new Date();
    await user.save();
  }

  const plan = getEffectivePlanFromUser(user);
  const limits = getLimits(plan);

  const used = user.usage[field] || 0;
  const limit = limits[field];

  if (limit == null) return { plan, used, limit: null };

  // PRO는 사실상 무제한(또는 limit이 크니 통과)
  if (typeof limit === "number" && used + amount > limit) {
    const err = new Error("LIMIT_EXCEEDED");
    err.status = 429;
    err.code = "LIMIT_EXCEEDED";
    err.meta = { field, used, limit, plan };
    throw err;
  }

  return { plan, used, limit };
}

module.exports = { getUsageSummary, bumpUsage, getLimits, assertCanUse };
