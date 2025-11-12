const User = require("../models/User");

const LIMITS = {
  FREE: { formulaConversions: 20, fileUploads: 1 },
  PRO: { formulaConversions: 5000, fileUploads: 5 },
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
  const user = await User.findById(userId).select("plan usage");
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

  const plan = user.plan || "FREE";
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
  const user = await User.findById(userId).select("plan usage");
  if (!user) throw new Error("User not found");
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
  return { usage: user.usage, limits: getLimits(user.plan || "FREE") };
}

module.exports = { getUsageSummary, bumpUsage, getLimits };
