const User = require("../models/User");
const paymentService = require("./paymentService");
const {
  getEffectivePlan,
  isSubscriptionActive,
} = require("../utils/subscriptionStatus");
const { getPlanLimits } = require("../config/planLimits");

function isLocalBypassMode() {
  return process.env.LOCAL_DEV === "1" && process.env.DEV_BYPASS_AUTH === "1";
}

function hasSubscriptionSignal(sub = {}) {
  return Boolean(
    sub?.billingKey ||
    sub?.customerKey ||
    sub?.startedAt ||
    sub?.trialEndsAt ||
    sub?.nextChargeAt,
  );
}

function getEffectivePlanFromUser(user) {
  // 베타모드일 때만 DB plan(PRO)을 그대로 인정
  if (paymentService.isBetaMode()) {
    return paymentService.getEffectivePlan(user?.plan || "FREE");
  }

  return getEffectivePlan(user);
}

function getLimits(plan = "FREE") {
  return getPlanLimits(plan);
}

function needMonthlyReset(lastReset, now = new Date()) {
  if (!lastReset) return true;
  return (
    lastReset.getUTCFullYear() !== now.getUTCFullYear() ||
    lastReset.getUTCMonth() !== now.getUTCMonth()
  );
}

function ensureUsageShape(user) {
  if (!user.usage) {
    user.usage = {
      templateGenerations: 0,
      fileUploads: 0,
      lastReset: new Date(),
    };
    return true;
  }

  let changed = false;

  if (typeof user.usage.templateGenerations !== "number") {
    user.usage.templateGenerations =
      typeof user.usage.formulaConversions === "number"
        ? user.usage.formulaConversions
        : 0;
    changed = true;
  }

  if (typeof user.usage.fileUploads !== "number") {
    user.usage.fileUploads = 0;
    changed = true;
  }

  if (!user.usage.lastReset) {
    user.usage.lastReset = new Date();
    changed = true;
  }

  return changed;
}

function resetMonthlyUsage(user) {
  user.usage.templateGenerations = 0;
  user.usage.fileUploads = 0;
  user.usage.lastReset = new Date();
}

async function getUsageSummary(userId) {
  if (isLocalBypassMode()) {
    return {
      plan: "PRO",
      usage: {
        templateGenerations: 0,
        fileUploads: 0,
      },
      limits: getLimits("PRO"),
    };
  }

  const user = await User.findById(userId).select("plan usage subscription");
  if (!user) throw new Error("User not found");

  let changed = false;
  changed = ensureUsageShape(user) || changed;
  if (needMonthlyReset(user.usage.lastReset)) {
    resetMonthlyUsage(user);
    changed = true;
  }
  if (changed) await user.save();

  const plan = getEffectivePlanFromUser(user);
  return {
    plan,
    usage: {
      templateGenerations: user.usage.templateGenerations,
      fileUploads: user.usage.fileUploads,
    },
    limits: getLimits(plan),
  };
}

async function bumpUsage(userId, field, delta) {
  if (isLocalBypassMode()) {
    return {
      skipped: true,
      usage: {
        templateGenerations: 0,
        fileUploads: 0,
      },
      limits: getLimits("PRO"),
      plan: "PRO",
    };
  }

  const user = await User.findById(userId).select(
    "plan usage subscription isDeleted",
  );
  if (!user) throw new Error("User not found");

  if (user.isDeleted) return { skipped: true };
  ensureUsageShape(user);
  if (needMonthlyReset(user.usage.lastReset)) {
    resetMonthlyUsage(user);
  }
  user.usage[field] = Math.max(0, (user.usage[field] || 0) + delta);
  await user.save();
  const plan = getEffectivePlanFromUser(user);
  return { usage: user.usage, limits: getLimits(plan), plan };
}

async function assertCanUse(userId, field, amount = 1) {
  if (isLocalBypassMode()) {
    return {
      plan: "PRO",
      used: 0,
      limit: null,
    };
  }

  const user = await User.findById(userId).select(
    "plan subscription usage isDeleted purgeAt",
  );
  if (!user) throw new Error("User not found");

  if (user.isDeleted) {
    const err = new Error("ACCOUNT_DELETED");
    err.status = 403;
    err.code = "ACCOUNT_DELETED";
    throw err;
  }

  // 월 리셋 동일 로직
  let changed = ensureUsageShape(user);
  if (needMonthlyReset(user.usage.lastReset)) {
    resetMonthlyUsage(user);
    changed = true;
  }
  if (changed) await user.save();

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
