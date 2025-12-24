const paymentService = require("../services/paymentService");
const User = require("../models/User");
const Payment = require("../models/Payment");
const bcrypt = require("bcryptjs");

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
    const user = await User.findById(req.user.id).select(
      "name email plan subscription"
    );
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    // ✅ 이미 구독/체험/해지예약 중이면 checkout 막기
    if (paymentService.isSubscriptionActive(user.subscription)) {
      return res.status(409).json({
        error: "이미 구독(또는 무료체험) 진행 중입니다.",
        code: "SUBSCRIPTION_ALREADY_ACTIVE",
        status: user.subscription?.status,
      });
    }

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
      status: "ACTIVE",
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
    const secret = req.headers["x-cron-secret"];
    if (!process.env.CRON_SECRET || secret !== process.env.CRON_SECRET) {
      return res.status(401).json({ error: "Unauthorized cron" });
    }

    // 베타모드면 청구/정리 모두 스킵(원하면 정리만 수행하도록 바꿔도 됨)
    if (paymentService.isBetaMode()) {
      return res.json({ ok: true, skipped: true, reason: "BETA_MODE=true" });
    }

    const now = new Date();

    // (1) 만료 처리: TRIAL 해지(CANCELED)인데 체험 종료됨 -> FREE로 다운그레이드
    const trialEnded = await User.updateMany(
      {
        plan: "PRO",
        "subscription.status": "CANCELED",
        "subscription.trialEndsAt": { $ne: null, $lte: now },
      },
      {
        $set: {
          plan: "FREE",
          "subscription.endedAt": now,
        },
      }
    );

    // (2) 만료 처리: 유료 해지 예약(CANCELED_PENDING)인데 nextChargeAt(=이용 만료일) 지남 -> FREE로 다운그레이드
    const paidEnded = await User.updateMany(
      {
        plan: "PRO",
        "subscription.status": "CANCELED_PENDING",
        "subscription.nextChargeAt": { $ne: null, $lte: now },
      },
      {
        $set: {
          plan: "FREE",
          "subscription.status": "CANCELED", // 최종 종료 상태로 정리
          "subscription.endedAt": now,
          "subscription.nextChargeAt": null,
        },
      }
    );

    // 구독 청구 금액/상품명
    const amount = Number(process.env.SUBSCRIPTION_AMOUNT || 4900);
    const orderName = process.env.SUBSCRIPTION_ORDER_NAME || "BeeBee AI PRO";

    // (3) 청구 대상 조회: ACTIVE/PAST_DUE만 청구하고, CANCELED/CANCELED_PENDING은 제외
    const targets = await User.find(
      {
        plan: "PRO",
        "subscription.billingKey": { $exists: true, $ne: null },
        "subscription.nextChargeAt": { $ne: null, $lte: now },
        "subscription.status": { $in: ["TRIAL", "ACTIVE", "PAST_DUE"] },
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
              "subscription.customerKey": customerKey,
            },
            $unset: { "subscription.lastChargeError": "" },
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
      cleaned: {
        trialEnded: trialEnded?.modifiedCount ?? trialEnded?.nModified ?? 0,
        paidEnded: paidEnded?.modifiedCount ?? paidEnded?.nModified ?? 0,
      },
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

exports.cancelSubscription = async (req, res) => {
  try {
    const { password } = req.body || {};

    // password 비교하려면 password 필드를 select해야 함 (스키마에서 select:false인 경우)
    const user = await User.findById(req.user.id).select(
      "+password plan subscription"
    );
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    const sub = user.subscription || {};
    const status = String(sub.status || "").toUpperCase();
    const now = new Date();

    // ✅ 이미 해지 완료 상태: 멱등 처리 + 분류 메시지
    if (status === "CANCELED") {
      return res.json({
        ok: true,
        code: "ALREADY_CANCELED",
        status: "CANCELED",
        message: "이미 구독 해지가 완료된 상태입니다.",
      });
    }

    // ✅ 이미 해지 예약(기간말 해지) 상태: 멱등 처리 + 분류 메시지
    if (status === "CANCELED_PENDING") {
      return res.json({
        ok: true,
        code: "ALREADY_CANCELED_PENDING",
        status: "CANCELED_PENDING",
        expiresAt: sub.nextChargeAt || null, // 만료일(= nextChargeAt 유지)
        message:
          "이미 구독 해지가 접수되었습니다. 이용 만료일까지 사용 가능합니다.",
      });
    }

    // ✅ 여기서부터는 “해지 처리”를 실제로 해야 하는 상태(TRIAL/ACTIVE/PAST_DUE 등)
    // 소셜 로그인 등으로 password가 없을 수 있음
    if (!user.password) {
      return res.status(400).json({
        error:
          "비밀번호가 설정되지 않은 계정입니다. 비밀번호 설정 후 해지할 수 있습니다.",
        code: "PASSWORD_NOT_SET",
      });
    }

    if (!password) {
      return res.status(400).json({ error: "password is required" });
    }

    const ok = await bcrypt.compare(password, user.password);
    if (!ok) {
      return res.status(401).json({ error: "비밀번호가 올바르지 않습니다." });
    }

    // ✅ 무료 체험 중 해지: 체험은 끝까지 사용, 과금은 절대 발생하면 안 됨
    const inTrial =
      status === "TRIAL" && sub.trialEndsAt && new Date(sub.trialEndsAt) > now;

    if (inTrial) {
      user.subscription = {
        ...sub,
        status: "CANCELED",
        canceledAt: now,
        endedAt: now,
        nextChargeAt: null, // ✅ 체험 끝에 과금 절대 방지
        cancelAtPeriodEnd: false,
      };

      await user.save();
      return res.json({
        ok: true,
        code: "CANCELED_TRIAL",
        status: "CANCELED",
        expiresAt: sub.trialEndsAt || null, // 체험 만료일까지 사용 가능(표시용)
        message:
          "무료 체험 해지가 완료되었습니다. 체험 종료일까지 이용 가능합니다.",
      });
    }

    // ✅ 유료/그 외: 기간말 해지(=다음 결제일부터 자동결제 중단, 만료일까지 사용)
    user.subscription = {
      ...sub,
      status: "CANCELED_PENDING",
      canceledAt: now,
      cancelAtPeriodEnd: true,
      // nextChargeAt 유지 (이용 만료일 역할)
    };

    await user.save();

    return res.json({
      ok: true,
      code: "CANCELED_PENDING",
      status: "CANCELED_PENDING",
      expiresAt: user.subscription.nextChargeAt || null,
      message: "구독 해지가 접수되었습니다. 이용 만료일까지 사용 가능합니다.",
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "구독 해지 실패" });
  }
};
