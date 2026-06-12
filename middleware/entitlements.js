const User = require("../models/User");
const { getLimits } = require("../services/usageService");
const { getEffectivePlan } = require("../utils/subscriptionStatus");

module.exports = function requireEntitlement(feature) {
  // feature: "formulaConversion" | "fileUpload"
  return async (req, res, next) => {
    const user = await User.findById(req.user.id, "plan usage subscription");
    if (!user) return res.status(401).json({ error: "Unauthorized" });

    const plan = getEffectivePlan(user);
    const limits = getLimits(plan);
    const usage = {
      formulaConversions: user.usage?.formulaConversions || 0,
      fileUploads: user.usage?.fileUploads || 0,
    };

    if (plan !== "PRO") {
      if (
        feature === "formulaConversion" &&
        usage.formulaConversions >= limits.formulaConversions
      ) {
        return res
          .status(403)
          .json({ error: "FREE 플랜은 월 10회까지만 생성이 가능합니다." });
      }
      if (feature === "fileUpload" && usage.fileUploads >= limits.fileUploads) {
        return res
          .status(403)
          .json({ error: "FREE 플랜은 파일 업로드 1회만 가능합니다." });
      }
    }
    next();
  };
};
