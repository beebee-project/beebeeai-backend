const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const ctrl = require("../controllers/paymentController");

// 결제/플랜
router.get("/plans", ctrl.getPlans);
router.get("/usage", protect, ctrl.getUsage);
router.post("/checkout", protect, ctrl.createCheckout);
router.post("/confirm", protect, ctrl.confirmPayment);
// ✅ 구독(빌링키) start/complete
router.post("/subscription/start", protect, ctrl.startSubscription);
router.post("/subscription/complete", protect, ctrl.completeSubscription);
// ✅ Cron: nextChargeAt 지난 구독자 자동 청구
router.post("/cron/charge", ctrl.cronCharge);

// router.post("/webhook", ctrl.webhook); // 필요 시 나중에 활성화

module.exports = router;
