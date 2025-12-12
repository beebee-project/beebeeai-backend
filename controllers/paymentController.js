const paymentService = require("../services/paymentService");
const User = require("../models/User");

const PROVIDER = String(process.env.PG_PROVIDER || "toss").toLowerCase();
const CURRENCY = process.env.CURRENCY || "KRW";
const isBeta = String(process.env.BETA_MODE).toLowerCase() === "true";
const PUBLIC_ORIGIN = process.env.PUBLIC_ORIGIN || "https://beebeeai.kr";

// 사용량 조회
exports.getUsage = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select("plan usage");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    res.json({
      plan: user.plan,
      usage: {
        formulaConversions: user?.usage?.formulaConversions ?? 0,
        fileUploads: user?.usage?.fileUploads ?? 0,
      },
      // 필요하면 추후 limits 확장
      limits: {
        formulaConversions: user?.limits?.formulaConversions ?? null,
        fileUploads: user?.limits?.fileUploads ?? null,
      },
    });
  } catch (err) {
    console.error("getUsage error:", err);
    res.status(500).json({ error: "사용량 조회 실패" });
  }
};

// 플랜 목록
exports.getPlans = (req, res) => {
  res.json({
    provider: PROVIDER,
    currency: CURRENCY,
    betaMode: isBeta,
    plans: [
      {
        code: "FREE",
        price: 0,
        interval: "month",
        features: ["NL→Sheet"],
        available: true,
      },
      {
        code: "PRO",
        price: 4900,
        interval: "month",
        features: ["우선지원", "고급기능"],
        available: !isBeta, // 베타 중에는 결제 비활성화
      },
    ],
  });
};

// 결제 시작: 프론트에서 Toss Payment Widget 띄우기 전에 호출
exports.createCheckout = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select("name email plan");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    // 현재는 PRO 월 구독 9,000원만 있다고 가정
    const amount = 4900;

    const successUrl = `${PUBLIC_ORIGIN}/success.html?provider=${PROVIDER}`;
    const failUrl = `${PUBLIC_ORIGIN}/fail.html?provider=${PROVIDER}`;

    const session = await paymentService.createCheckoutSession({
      userId: String(user._id),
      amount,
      successUrl,
      failUrl,
      meta: {
        orderName: "BeeBee AI PRO (월 정기 결제)",
        customerName: user.name || user.email,
      },
    });

    // 프론트에서 Toss Payment Widget에 넘길 데이터
    return res.json({
      provider: session.provider, // "toss"
      orderId: session.orderId,
      amount: session.amount,
      currency: session.currency,
      orderName: session.orderName,
      customerName: session.customerName,
      successUrl: session.successUrl,
      failUrl: session.failUrl,
    });
  } catch (err) {
    console.error("createCheckout error:", err);
    res.status(500).json({
      error: "결제 세션 생성 실패",
      code: err.code,
    });
  }
};

// 결제 승인: successUrl로 리다이렉트된 후, 프론트에서 호출
// body: { paymentKey, orderId, amount }
exports.confirmPayment = async (req, res) => {
  try {
    const { paymentKey, orderId, amount } = req.body || {};

    if (!paymentKey || !orderId || !amount) {
      return res
        .status(400)
        .json({ error: "paymentKey, orderId, amount는 필수입니다." });
    }

    const expectedAmount = 4900;
    const numericAmount = Number(amount);

    if (Number.isNaN(numericAmount)) {
      return res
        .status(400)
        .json({ error: "amount 형식이 올바르지 않습니다." });
    }
    if (numericAmount !== expectedAmount) {
      return res.status(400).json({
        error: "요청 금액이 서버 설정 금액과 일치하지 않습니다.",
      });
    }

    // Toss Payments에 실제 승인 요청
    const result = await paymentService.confirmPayment({
      paymentKey,
      orderId,
      amount: expectedAmount,
    });

    // 유저 플랜 업데이트 (간단히 PRO로 승급)
    const user = await User.findById(req.user.id);
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    user.plan = "PRO";
    await user.save();

    return res.json({
      ok: true,
      provider: result.provider, // "toss"
      orderId: result.orderId,
      amount: result.amount,
      paymentKey: result.paymentKey,
    });
  } catch (err) {
    console.error("confirmPayment error:", err);

    const tossError = err.response?.data;
    if (tossError) {
      console.error("Toss error response:", tossError);
    }

    res.status(500).json({
      error: "결제 승인 실패",
      code: err.code,
      toss: tossError,
    });
  }
};

// 웹훅은 나중에 필요해지면 구현
// exports.webhook = async (req, res) => { ... };
