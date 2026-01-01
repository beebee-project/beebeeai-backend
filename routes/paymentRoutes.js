const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const ctrl = require("../controllers/paymentController");

const isBeta = String(process.env.BETA_MODE || "").toLowerCase() === "true";
const blockInBeta = (req, res, next) =>
  isBeta
    ? res
        .status(403)
        .json({ error: "베타 기간에는 결제가 비활성화되어 있습니다." })
    : next();

// ✅ 기존(프론트가 이미 호출 중인) 라우트 복구
router.get("/plans", ctrl.getPlans);
router.get("/usage", protect, ctrl.getUsage);
router.post("/checkout", protect, blockInBeta, ctrl.createCheckout);
router.post("/confirm", protect, blockInBeta, ctrl.confirmPayment);

// ✅ 정기결제(빌링키) 라우트 유지
router.post(
  "/subscription/start",
  protect,
  blockInBeta,
  ctrl.startSubscription
);
router.post(
  "/subscription/complete",
  protect,
  blockInBeta,
  ctrl.completeSubscription
);
router.post(
  "/subscription/cancel",
  protect,
  blockInBeta,
  ctrl.cancelSubscription
);

// cron
router.post("/cron/charge", ctrl.cronCharge);

module.exports = router;
