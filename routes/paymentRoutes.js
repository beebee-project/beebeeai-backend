const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const ctrl = require("../controllers/paymentController");

// 결제/플랜
router.get("/plans", ctrl.getPlans);
router.get("/usage", protect, ctrl.getUsage);
router.post("/checkout", protect, ctrl.createCheckout);
router.post("/confirm", protect, ctrl.confirmPayment);
router.post("/subscription/start", protect, ctrl.startSubscription);
router.post("/subscription/complete", protect, ctrl.completeSubscription);
router.post("/subscription/cancel", protect, ctrl.cancelSubscription);
router.post("/cron/charge", ctrl.cronCharge);

router.post("/webhook", ctrl.webhook);

module.exports = router;
