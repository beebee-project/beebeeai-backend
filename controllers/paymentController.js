const paymentService = require("../services/paymentService");
const User = require("../models/User");

// 사용량 조회
exports.getUsage = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select("plan usage");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    res.json({
      plan: user.plan,
      usage: {
        formulaConversions: user.usage.formulaConversions,
        fileUploads: user.usage.fileUploads,
      },
      limits:
        user.plan === "PRO"
          ? { formulaConversions: 5000, fileUploads: 5 }
          : { formulaConversions: 20, fileUploads: 1 },
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "사용 현황 조회 실패" });
  }
};

// 결제 세션 생성
exports.createCheckout = async (req, res) => {
  try {
    // 금액은 서버가 결정 (price.js/price.json 등 한 곳에서만 관리)
    const amount = 9000; // 또는 price.js에서 읽어오도록 통일
    const successUrl = `${process.env.PUBLIC_ORIGIN}/price.html?pg=success&provider=sprite`;
    const failUrl = `${process.env.PUBLIC_ORIGIN}/price.html?pg=fail&provider=sprite`;

    const session = await paymentService.createCheckoutSession({
      userId: req.user.id,
      amount,
      successUrl,
      failUrl,
      meta: { plan: "PRO" },
    });

    // 프런트는 checkoutUrl로 리다이렉트만 하면 됨
    return res.json({
      provider: "sprite",
      orderId: session.orderId,
      amount,
      orderName: "BeeBee AI PRO (월결제)",
      customerName: req.user.name || "회원",
      checkoutUrl: session.checkoutUrl,
      successUrl: session.successUrl,
      failUrl: session.failUrl,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "결제 세션 생성 실패" });
  }
};

// 결제 승인
exports.confirmPayment = async (req, res) => {
  try {
    const amount = 9000; // 서버 금액과 반드시 동일
    const { provider = "sprite", orderId } = req.body;

    // Sprite 확인 (Toss는 더 이상 사용 안 함)
    await paymentService.confirmPayment({
      provider,
      orderId,
      expectedAmount: amount,
    });

    // 결제 성공 → 사용자 플랜 승격 + 사용량 초기화
    const user = await User.findById(req.user.id);
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    user.plan = "PRO";
    user.usage = {
      formulaConversions: 0,
      fileUploads: 0,
      lastReset: new Date(),
    };
    await user.save();

    res.json({ ok: true, message: "PRO 플랜 활성화 완료" });
  } catch (e) {
    console.error(e);
    res.status(400).json({ error: e.message || "결제 승인 실패" });
  }
};

// 플랜 목록
exports.getPlans = async (req, res) => {
  const isBeta = String(process.env.BETA_MODE).toLowerCase() === "true";
  res.json({
    beta: isBeta,
    plans: [
      {
        code: "FREE_BETA",
        price: 0,
        interval: "month",
        features: ["NL→Sheet"],
        available: true,
      },
      {
        code: "PRO",
        price: 9000,
        interval: "month",
        features: ["우선지원", "고급기능"],
        available: !isBeta,
      },
    ],
  });
};
