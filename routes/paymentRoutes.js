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

function requireCronSecret(req, res, next) {
  const secret = process.env.CRON_SECRET;
  if (!secret) return res.status(500).json({ error: "CRON_SECRET is not set" });

  const provided =
    req.get("x-cron-secret") ||
    (req.get("authorization") || "").replace(/^Bearer\s+/i, "");

  if (provided !== secret) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

router.post("/cron/charge", requireCronSecret, ctrl.cronCharge);

router.post("/webhook", ctrl.webhook);

module.exports = router;
