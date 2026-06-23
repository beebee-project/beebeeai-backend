const PLAN_LIMITS = {
  FREE: {
    templateGenerations: 5,
    fileUploads: 1,
  },
  PRO: {
    templateGenerations: null,
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
