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

    if (user.subscription?.billingKey) {
      return res.status(409).json({
        error: "이미 구독 설정된 계정입니다. 구독 결제 흐름을 사용해주세요.",
      });
    }

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

exports.startSubscription = async (req, res) => {
  try {
    // 베타모드면 결제 없이 PRO
    if (paymentService.isBetaMode()) {
      await User.findByIdAndUpdate(req.user.id, { plan: "PRO" });
      return res.json({
        ok: true,
        beta: true,
        plan: "PRO",
        message: "BETA_MODE=true: 결제 없이 PRO가 활성화되었습니다.",
      });
    }

    const customerKey = String(req.user.id);

    return res.json({
      ok: true,
      beta: false,
      customerKey,
      successUrl:
        process.env.SUBSCRIPTION_SUCCESS_URL || "https://beebeeai.kr/success", // 네 운영 주소에 맞춰 조정
      failUrl: process.env.SUBSCRIPTION_FAIL_URL || "https://beebeeai.kr/fail",
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "구독 시작 실패" });
  }
};

exports.completeSubscription = async (req, res) => {
  try {
    if (paymentService.isBetaMode()) {
      // 베타모드면 complete를 호출해도 PRO 유지
      await User.findByIdAndUpdate(req.user.id, { plan: "PRO" });
      return res.json({ ok: true, beta: true, plan: "PRO" });
    }

    const { customerKey, authKey } = req.body;
    if (!customerKey || !authKey) {
      return res
        .status(400)
        .json({ error: "customerKey/authKey가 필요합니다." });
    }

    // billingKey 발급
    const issued = await paymentService.issueBillingKey({
      customerKey,
      authKey,
    });
    if (!issued?.billingKey) {
      return res.status(500).json({ error: "billingKey 발급 실패" });
    }

    // 7일 무료체험 -> 체험 종료 시점에 첫 과금
    const now = new Date();
    const trialEndsAt = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
    const nextChargeAt = trialEndsAt;

    const user = await User.findById(req.user.id);
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    user.plan = "PRO";
    user.subscription = {
      ...(user.subscription || {}),
      customerKey,
      billingKey: issued.billingKey,
      status: "TRIAL",
      trialEndsAt,
      nextChargeAt,
      lastChargedAt: null,
    };

    await user.save();

    return res.json({
      ok: true,
      plan: "PRO",
      trialEndsAt,
      nextChargeAt,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "구독 완료 처리 실패" });
  }
};

exports.cronCharge = async (req, res) => {
  try {
    // ✅ 보안: CRON_SECRET
    const secret = req.headers["x-cron-secret"];
    if (!process.env.CRON_SECRET || secret !== process.env.CRON_SECRET) {
      return res.status(401).json({ error: "Unauthorized cron" });
    }

    // ✅ 베타모드면 청구 자체를 하지 않음
    if (paymentService.isBetaMode()) {
      return res.json({ ok: true, skipped: true, reason: "BETA_MODE=true" });
    }

    const now = new Date();

    // 금액/상품명은 env로 빼두는 걸 추천 (없으면 기본값 사용)
    const amount = Number(process.env.SUBSCRIPTION_AMOUNT || 4900);
    const orderName = process.env.SUBSCRIPTION_ORDER_NAME || "BeeBee AI PRO";

    // ✅ 청구 대상 조회
    const targets = await User.find(
      {
        plan: "PRO",
        "subscription.billingKey": { $exists: true, $ne: null },
        "subscription.nextChargeAt": { $lte: now },
        "subscription.status": { $ne: "CANCELED" },
      },
      "_id subscription"
    ).lean();

    let successCount = 0;
    let failCount = 0;

    for (const u of targets) {
      const userId = String(u._id);
      const customerKey = u.subscription?.customerKey || userId;
      const billingKey = u.subscription?.billingKey;

      if (!billingKey) continue;

      // ✅ orderId는 매 청구마다 유니크해야 함
      const orderId = `sub-${userId}-${Date.now()}`;

      try {
        await paymentService.chargeBillingKey({
          customerKey,
          billingKey,
          amount,
          orderId,
          orderName,
        });

        const nextChargeAt = paymentService.addMonths(now, 1);

        await User.updateOne(
          { _id: u._id },
          {
            $set: {
              "subscription.status": "ACTIVE",
              "subscription.lastChargedAt": now,
              "subscription.nextChargeAt": nextChargeAt,
              // customerKey가 비어있던 케이스 정리
              "subscription.customerKey": customerKey,
            },
          }
        );

        successCount += 1;
      } catch (e) {
        failCount += 1;

        await User.updateOne(
          { _id: u._id },
          {
            $set: {
              "subscription.status": "PAST_DUE",
              "subscription.lastChargeError": String(
                e?.response?.data?.message || e?.message || e
              ).slice(0, 500),
            },
          }
        );
      }
    }

    return res.json({
      ok: true,
      now,
      targets: targets.length,
      successCount,
      failCount,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Cron charge failed" });
  }
};

// 웹훅은 나중에 필요해지면 구현
// exports.webhook = async (req, res) => { ... };
