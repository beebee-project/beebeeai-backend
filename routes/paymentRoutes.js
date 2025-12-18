const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const ctrl = require("../controllers/paymentController");

const isBetaMode = () => String(process.env.BETA_MODE).toLowerCase() === "true";

// ✅ 베타모드일 때는 "결제 차단"이 아니라 "프리모드(프로 허용)"로 바꿀 거라면
// 결제 라우트 자체는 막지 않는 게 맞음. (아래에서 remove)
const blockInBeta = (req, res, next) =>
  isBetaMode()
    ? res
        .status(403)
        .json({ error: "베타 기간에는 결제가 비활성화되어 있습니다." })
    : next();

// 결제/플랜
router.get("/plans", ctrl.getPlans);
router.get("/usage", protect, ctrl.getUsage);
router.post("/checkout", protect, ctrl.createCheckout);
router.post("/confirm", protect, ctrl.confirmPayment);

// router.post("/webhook", ctrl.webhook); // 필요 시 나중에 활성화

module.exports = router;
