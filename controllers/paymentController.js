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

// 결제 시작
exports.createCheckout = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select(
      "name email plan subscription"
    );
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    // 구독 플로우는 checkout/confirm(일반결제)로 못 타게 강제
    if (user.plan === "PRO") {
      return res.status(400).json({
        error:
          "PRO는 정기구독(billingKey) 플로우로만 가능합니다. /api/payments/subscription/start 를 사용하세요.",
        code: "SUBSCRIPTION_ONLY",
      });
    }

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

// 결제 승인
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
    const user = await User.findById(req.user.id).select("plan subscription");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    // ❗ 구독 계정은 confirmPayment 금지
    if (user.subscription?.billingKey || user.plan === "PRO") {
      return res.status(400).json({
        error: "PRO 구독은 subscription/complete(authKey)로만 처리됩니다.",
        code: "USE_BILLING_SUBSCRIPTION",
      });
    }

    if (user.subscription?.billingKey) {
      return res.status(409).json({
        error: "이미 구독 설정된 계정입니다. 구독 결제 흐름을 사용해주세요.",
      });
    }

    const now = new Date();
    const expiresAt = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

    user.subscription = {
      startedAt: user.subscription?.startedAt || now,
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
    // ✅ 0) CRON 보호
    const secret = req.headers["x-cron-secret"];
    if (!process.env.CRON_SECRET || secret !== process.env.CRON_SECRET) {
      return res.status(401).json({ ok: false, error: "Unauthorized" });
    }

    const now = new Date();

    // ✅ 1) 과금 대상: billingKey 있고, 상태가 TRIAL/ACTIVE/PAST_DUE 인 유저
    // (CANCELED_PENDING은 기간말 해지라 과금하면 안 됨)
    const targets = await User.find({
      "subscription.billingKey": { $exists: true, $ne: null },
      "subscription.status": { $in: ["TRIAL", "ACTIVE", "PAST_DUE"] },
    }).select("email plan subscription");

    let successCount = 0;
    let failCount = 0;
    let skippedCount = 0;

    for (const user of targets) {
      const sub = user.subscription || {};
      const status = String(sub.status || "").toUpperCase();

      // 안전장치: cancelAtPeriodEnd=true 인 경우 과금 스킵
      if (sub.cancelAtPeriodEnd === true) {
        skippedCount++;
        continue;
      }

      // ✅ TRIAL → trialEndsAt 도래 시 1회 과금 후 ACTIVE 전환
      if (status === "TRIAL") {
        if (!sub.trialEndsAt) {
          // trialEndsAt 없으면 비정상 데이터 → 스킵
          skippedCount++;
          continue;
        }
        if (new Date(sub.trialEndsAt) > now) {
          // 아직 체험중
          skippedCount++;
          continue;
        }

        // 체험 끝났는데도 nextChargeAt이 null이면 과금해야 함
        const amount = 4900;

        try {
          const orderId = `beebeeai-${Date.now()}-${Math.random()
            .toString(16)
            .slice(2)}`;

          // ✅ billingKey 과금 (네 구현에 맞게 호출명 변경)
          const chargeResult = await paymentService.chargeBillingKey({
            billingKey: sub.billingKey,
            customerKey: sub.customerKey,
            amount,
            orderId,
            orderName: "BeeBee AI PRO (월 정기 결제)",
          });

          // ✅ 결제 로그(옵션) - Payment 컬렉션에 기록
          await Payment.updateOne(
            { orderId },
            {
              $setOnInsert: {
                userId: String(user._id),
                orderId,
                createdAt: now,
              },
              $set: {
                amount,
                currency: "KRW",
                provider: "toss",
                status: "DONE",
                paymentKey: chargeResult.paymentKey || null,
                raw: chargeResult.raw || chargeResult,
                approvedAt: now,
                updatedAt: now,
              },
            },
            { upsert: true }
          );

          user.plan = "PRO";
          user.subscription = {
            ...sub,
            status: "ACTIVE",
            startedAt: sub.startedAt || now,
            lastOrderId: orderId,
            lastPaymentKey: chargeResult.paymentKey || sub.lastPaymentKey,
            // ✅ 첫 과금 시점부터 한 달 뒤 재결제
            nextChargeAt: paymentService.addMonths(now, 1),
            // trial 흔적은 남겨도 되고 지워도 됨. 남기는 편이 운영에 유리.
            // trialEndsAt: sub.trialEndsAt,
          };

          await user.save();
          successCount++;
        } catch (e) {
          console.error("[cronCharge][TRIAL] charge failed:", e);

          user.subscription = {
            ...sub,
            status: "PAST_DUE",
            // nextChargeAt은 그대로 두거나, 재시도 전략을 위해 now+1일 같은 값을 줄 수도 있음
          };
          await user.save();
          failCount++;
        }

        continue;
      }

      // ✅ ACTIVE/PAST_DUE → nextChargeAt 도래 시 정기 과금
      if (!sub.nextChargeAt) {
        // nextChargeAt 없으면 스킵 (비정상)
        skippedCount++;
        continue;
      }
      if (new Date(sub.nextChargeAt) > now) {
        // 아직 결제일 아님
        skippedCount++;
        continue;
      }

      const amount = 4900;

      try {
        const orderId = `beebeeai-${Date.now()}-${Math.random()
          .toString(16)
          .slice(2)}`;

        const chargeResult = await paymentService.chargeBillingKey({
          billingKey: sub.billingKey,
          customerKey: sub.customerKey,
          amount,
          orderId,
          orderName: "BeeBee AI PRO (월 정기 결제)",
        });

        await Payment.updateOne(
          { orderId },
          {
            $setOnInsert: {
              userId: String(user._id),
              orderId,
              createdAt: now,
            },
            $set: {
              amount,
              currency: "KRW",
              provider: "toss",
              status: "DONE",
              paymentKey: chargeResult.paymentKey || null,
              raw: chargeResult.raw || chargeResult,
              approvedAt: now,
              updatedAt: now,
            },
          },
          { upsert: true }
        );

        user.plan = "PRO";
        user.subscription = {
          ...sub,
          status: "ACTIVE",
          lastOrderId: orderId,
          lastPaymentKey: chargeResult.paymentKey || sub.lastPaymentKey,
          nextChargeAt: paymentService.addMonths(now, 1),
        };

        await user.save();
        successCount++;
      } catch (e) {
        console.error("[cronCharge][ACTIVE/PAST_DUE] charge failed:", e);

        user.subscription = {
          ...sub,
          status: "PAST_DUE",
        };
        await user.save();
        failCount++;
      }
    }

    return res.json({
      ok: true,
      now,
      targets: targets.length,
      successCount,
      failCount,
      skippedCount,
    });
  } catch (e) {
    console.error("cronCharge error:", e);
    return res.status(500).json({ ok: false, error: "cronCharge failed" });
  }
};

// 웹훅은 나중에 필요해지면 구현
// exports.webhook = async (req, res) => { ... };

exports.cancelSubscription = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select("plan subscription");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    const sub = user.subscription || {};
    const status = String(sub.status || "NONE").toUpperCase();

    const hasAnySubscriptionSignal = !!(
      sub.billingKey ||
      sub.customerKey ||
      sub.startedAt ||
      sub.trialEndsAt ||
      sub.nextChargeAt
    );

    if (!hasAnySubscriptionSignal && status === "NONE") {
      return res.status(409).json({
        ok: false,
        code: "NO_SUBSCRIPTION",
        message: "현재 구독 중이 아닙니다.",
        status: "NONE",
      });
    }

    // ✅ 이미 완전 해지라면 idempotent
    if (status === "CANCELED") {
      return res.json({
        ok: true,
        code: "ALREADY_CANCELED",
        message: "이미 구독 해지가 완료된 상태입니다.",
        status,
      });
    }

    // ✅ 기간말 해지(이미 접수됨)도 idempotent
    if (status === "CANCELED_PENDING") {
      return res.json({
        ok: true,
        code: "ALREADY_CANCELED_PENDING",
        message:
          "이미 구독 해지가 접수되었습니다. 이용 만료일까지 사용 가능합니다.",
        status,
        expiresAt: sub.expiresAt || sub.nextChargeAt || null,
      });
    }

    const now = new Date();

    // ✅ 무료 체험 중 해지: 체험은 끝까지 사용, 과금만 막기
    const inTrial =
      status === "TRIAL" && sub.trialEndsAt && new Date(sub.trialEndsAt) > now;

    if (inTrial) {
      user.subscription = {
        ...sub,
        status: "CANCELED",
        canceledAt: now,
        endedAt: now,
        nextChargeAt: null,
        cancelAtPeriodEnd: false,
        expiresAt: sub.trialEndsAt,
      };
      await user.save();

      return res.json({
        ok: true,
        code: "CANCELED_TRIAL",
        status: "CANCELED",
        expiresAt: sub.trialEndsAt || null,
        message:
          "무료 체험 해지가 완료되었습니다. 체험 종료일까지 이용 가능합니다.",
      });
    }

    // ✅ 유료/기타: 기간말 해지(만료일까지 사용)
    user.subscription = {
      ...sub,
      status: "CANCELED_PENDING",
      canceledAt: now,
      cancelAtPeriodEnd: true,
      // nextChargeAt 유지 = 만료일까지 사용
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
    return res.status(500).json({ error: "구독 해지 실패" });
  }
};
