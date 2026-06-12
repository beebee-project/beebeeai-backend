const PLAN_LIMITS = {
  FREE: {
    formulaConversions: 10,
    fileUploads: 1,
  },
  PRO: {
    formulaConversions: null,
    fileUploads: null,
  },
};

function getPlanLimits(plan = "FREE") {
  return PLAN_LIMITS[plan] || PLAN_LIMITS.FREE;
}

module.exports = {
  PLAN_LIMITS,
  getPlanLimits,
};
