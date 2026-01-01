const express = require("express");
const router = express.Router();
const { protect } = require("../middleware/authMiddleware");
const paymentController = require("../controllers/paymentController");

// 사용량/플랜 (프론트에서 계속 호출 중)
router.get("/usage", protect, paymentController.getUsage);
router.get("/plans", paymentController.getPlans);

// (현재 프론트 결제 버튼이 호출하는 엔드포인트)
router.post("/checkout", protect, paymentController.createCheckout);

// 정기결제 시작: customerKey 기반으로 Toss 결제창(빌링) 오픈
router.post(
  "/subscription/start",
  protect,
  paymentController.startSubscription
);

// 정기결제 완료: success 페이지에서 authKey 받아 billingKey 발급/첫 결제/구독 활성화
router.post(
  "/subscription/complete",
  protect,
  paymentController.completeSubscription
);

// 구독 해지(기간말 해지)
router.post(
  "/subscription/cancel",
  protect,
  paymentController.cancelSubscription
);

// cron: nextChargeAt 도래한 ACTIVE 대상 월 과금
router.post("/cron/charge", paymentController.cronCharge); // CRON_SECRET 등으로 보호

module.exports = router;
