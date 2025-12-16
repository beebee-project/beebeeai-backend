const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const ctrl = require("../controllers/paymentController");

const isBeta = String(process.env.BETA_MODE).toLowerCase() === "true";
const blockInBeta = (req, res, next) =>
  isBeta
    ? res
        .status(403)
        .json({ error: "베타 기간에는 결제가 비활성화되어 있습니다." })
    : next();

// 결제/플랜
router.get("/plans", ctrl.getPlans);
router.get("/usage", protect, ctrl.getUsage);
router.post("/checkout", protect, ctrl.createCheckout);
router.post("/confirm", protect, ctrl.confirmPayment);
// router.post("/checkout", protect, blockInBeta, ctrl.createCheckout);
// router.post("/confirm", protect, blockInBeta, ctrl.confirmPayment);

// router.post("/webhook", ctrl.webhook); // 필요 시 나중에 활성화

module.exports = router;
