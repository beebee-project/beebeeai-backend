const User = require("../models/User");
const { getLimits } = require("../services/usageService");

module.exports = function requireEntitlement(feature) {
  // feature: "formulaConversion" | "fileUpload"
  return async (req, res, next) => {
    const user = await User.findById(req.user.id, "plan usage");
    if (!user) return res.status(401).json({ error: "Unauthorized" });

    const plan = user.plan || "FREE";
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
          .json({ error: "FREE 플랜은 월 20회까지만 수식 생성이 가능합니다." });
      }
      if (feature === "fileUpload" && usage.fileUploads >= limits.fileUploads) {
        return res
          .status(403)
          .json({ error: "FREE 플랜은 파일 업로드 1개까지만 가능합니다." });
      }
    }
    next();
  };
};
