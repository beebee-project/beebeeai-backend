const paymentService = require("../services/paymentService");
const User = require("../models/User");
const Payment = require("../models/Payment");

const PROVIDER = String(process.env.PG_PROVIDER || "toss").toLowerCase();
const CURRENCY = process.env.CURRENCY || "KRW";
const isBetaMode = () => String(process.env.BETA_MODE).toLowerCase() === "true";
const PUBLIC_ORIGIN = process.env.PUBLIC_ORIGIN || "https://beebeeai.kr";

// 사용량 조회
exports.getUsage = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select("plan usage");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    const plan = paymentService.getEffectivePlan(user.plan);

    const limits =
      plan === "PRO"
        ? { formulaConversions: 5000, fileUploads: 5 }
        : { formulaConversions: 20, fileUploads: 1 };

    res.json({
      plan,
      usage: {
        formulaConversions: user?.usage?.formulaConversions ?? 0,
        fileUploads: user?.usage?.fileUploads ?? 0,
      },
      limits,
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
    betaMode: isBetaMode(),
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
        available: !isBetaMode(),
      },
    ],
  });
};

// 결제 시작: 프론트에서 Toss Payment Widget 띄우기 전에 호출
exports.createCheckout = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select("name email plan");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    const amount = 4900;

    const successUrl = `${PUBLIC_ORIGIN}/success.html`;
    const failUrl = `${PUBLIC_ORIGIN}/fail.html`;

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

    await Payment.updateOne(
      { orderId: session.orderId },
      {
        $setOnInsert: {
          userId: String(user._id),
          orderId: session.orderId,
          createdAt: new Date(),
        },
        $set: {
          amount: session.amount,
          currency: session.currency,
          provider: session.provider,
          status: "READY",
          updatedAt: new Date(),
        },
      },
      { upsert: true }
    );
    return res.json({
      provider: session.provider,
      orderId: session.orderId,
      amount: session.amount,
      currency: session.currency,
      orderName: session.orderName,
      customerName: session.customerName,
      successUrl: session.successUrl,
      failUrl: session.failUrl,
      customerKey: String(user._id),
      status: "READY",
    });
  } catch (err) {
    console.error("createCheckout error:", err);
    res.status(500).json({ error: "결제 세션 생성 실패", code: err.code });
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

    const numericAmount = Number(amount);
    if (Number.isNaN(numericAmount)) {
      return res
        .status(400)
        .json({ error: "amount 형식이 올바르지 않습니다." });
    }

    // ✅ 1) 우리 DB에서 orderId 검증
    const pay = await Payment.findOne({ orderId, userId: String(req.user.id) });
    if (!pay)
      return res.status(404).json({ error: "존재하지 않는 orderId 입니다." });

    // ✅ 2) 본인 결제인지 확인 (protect 쓰는 구조면 userId 매칭 필수)
    if (String(pay.userId) !== String(req.user.id)) {
      return res
        .status(403)
        .json({ error: "본인의 결제만 승인할 수 있습니다." });
    }

    // ✅ 3) 금액 검증(서버 기준)
    if (numericAmount !== pay.amount) {
      return res
        .status(400)
        .json({ error: "요청 금액이 서버 금액과 일치하지 않습니다." });
    }

    // ✅ 4) 멱등 처리: 이미 승인 완료면 그대로 성공 반환
    if (pay.status === "DONE" && pay.paymentKey === paymentKey) {
      return res.json({
        ok: true,
        duplicated: true,
        orderId: pay.orderId,
        paymentKey: pay.paymentKey,
        amount: pay.amount,
      });
    }

    // ✅ 5) Toss 승인 호출
    const result = await paymentService.confirmPayment({
      paymentKey,
      orderId,
      amount: pay.amount,
    });

    // ✅ 6) Payment 업데이트
    pay.status = "DONE";
    pay.paymentKey = result.paymentKey;
    pay.raw = result.raw;
    pay.approvedAt = new Date(result.raw?.approvedAt || Date.now());
    await pay.save();

    // ✅ 7) User 구독 반영 (일단 30일 운영)
    const user = await User.findById(req.user.id);
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    const now = new Date();
    const expiresAt = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

    user.plan = "PRO";
    user.subscription = {
      status: "active",
      startedAt: user.subscription?.startedAt || now,
      expiresAt,
      lastPaymentKey: result.paymentKey,
      lastOrderId: orderId,
    };

    await user.save();

    return res.json({
      ok: true,
      provider: result.provider,
      orderId,
      amount: pay.amount,
      paymentKey: result.paymentKey,
      subscription: user.subscription,
    });
  } catch (err) {
    console.error("confirmPayment error:", err);
    return res.status(500).json({ error: "결제 승인 실패" });
  }
};

// 웹훅은 나중에 필요해지면 구현
// exports.webhook = async (req, res) => { ... };
