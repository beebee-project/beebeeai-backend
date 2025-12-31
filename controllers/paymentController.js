const crypto = require("crypto");
const User = require("../models/User");
const Payment = require("../models/Payment"); // 너 프로젝트 Payment 모델에 맞춰 경로 수정
const {
  addMonths,
  isSubscriptionLocked,
  isAlreadyCanceled,
} = require("../services/paymentService");

// 게이트웨이 (toss만 쓸 거면 바로 require해도 됨)
const toss = require("../services/paymentGateway/toss");

const PUBLIC_ORIGIN = process.env.PUBLIC_ORIGIN || "https://www.beebeeai.kr";
const CRON_SECRET = process.env.CRON_SECRET || "";

function makeOrderId(prefix = "beebeeai") {
  // 충돌 적게
  return `${prefix}-${Date.now()}-${crypto.randomBytes(2).toString("hex")}`;
}

/**
 * POST /api/payments/subscription/start
 * - 결제창을 열기 위한 세션(필요값)을 만들어서 프론트에 전달
 * - "이미 구독중"이면 막음
 */
exports.startSubscription = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select(
      "name email plan subscription"
    );
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    // ✅ 중복 구독 시작 방지
    if (isSubscriptionLocked(user.subscription)) {
      return res.status(409).json({
        error: "이미 구독(또는 결제 실패/해지예약) 상태입니다.",
        code: "SUBSCRIPTION_LOCKED",
        status: user.subscription?.status,
      });
    }

    const amount = 4900;
    const orderId = makeOrderId();

    // success/fail은 네가 만들어둔 success.html/fail.html로
    const successUrl = `${PUBLIC_ORIGIN}/success.html`;
    const failUrl = `${PUBLIC_ORIGIN}/fail.html`;

    const session = await toss.createBillingCheckoutSession({
      customerKey: String(user._id),
      orderId,
      amount,
      successUrl,
      failUrl,
      orderName: "BeeBee AI PRO (월 정기 결제)",
      customerName: user.name || user.email,
    });

    // Payment READY 저장(선택이지만 추천)
    await Payment.updateOne(
      { orderId },
      {
        $setOnInsert: {
          userId: String(user._id),
          orderId,
          createdAt: new Date(),
        },
        $set: {
          amount,
          currency: "KRW",
          provider: "toss",
          status: "READY",
          updatedAt: new Date(),
        },
      },
      { upsert: true }
    );

    return res.json(session);
  } catch (e) {
    console.error("startSubscription error:", e);
    return res.status(500).json({ error: "구독 결제 시작 실패" });
  }
};

/**
 * POST /api/payments/subscription/complete
 * - success.html에서 authKey를 받아서 호출
 * - billingKey 발급 -> billingKey로 첫 결제 -> User 구독 ACTIVE 세팅
 */
exports.completeSubscription = async (req, res) => {
  try {
    const { authKey, orderId } = req.body || {};
    if (!authKey || !orderId) {
      return res.status(400).json({ error: "authKey, orderId는 필수입니다." });
    }

    const user = await User.findById(req.user.id).select(
      "name email plan subscription"
    );
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    // ✅ 이미 구독이면 complete도 막아야 “중복 결제”가 안남
    if (isSubscriptionLocked(user.subscription)) {
      return res.status(409).json({
        error: "이미 구독 상태입니다.",
        code: "SUBSCRIPTION_LOCKED",
        status: user.subscription?.status,
      });
    }

    // ✅ Payment 검증(본인 orderId인지)
    const pay = await Payment.findOne({ orderId, userId: String(user._id) });
    if (!pay)
      return res.status(404).json({ error: "존재하지 않는 orderId 입니다." });
    if (pay.status === "DONE") {
      return res.json({
        ok: true,
        duplicated: true,
        subscription: user.subscription,
      });
    }

    // 1) billingKey 발급
    const issued = await toss.issueBillingKey({
      customerKey: String(user._id),
      authKey,
    });
    const billingKey = issued.billingKey;

    // 2) billingKey로 첫 결제 (즉시 1회 결제)
    const charge = await toss.chargeWithBillingKey({
      billingKey,
      customerKey: String(user._id),
      orderId, // 첫 결제 orderId는 start에서 만든 걸 그대로 사용
      amount: pay.amount,
      orderName: "BeeBee AI PRO (월 정기 결제)",
      customerEmail: user.email,
      customerName: user.name || user.email,
    });

    // 3) Payment DONE 업데이트
    pay.status = "DONE";
    pay.paymentKey = charge.paymentKey || charge.raw?.paymentKey;
    pay.raw = charge.raw || charge;
    pay.approvedAt = new Date(charge.raw?.approvedAt || Date.now());
    pay.updatedAt = new Date();
    await pay.save();

    // 4) User 구독 활성화 + 다음 결제일 1개월
    const now = new Date();
    user.plan = "PRO";
    user.subscription = {
      customerKey: String(user._id),
      billingKey,
      status: "ACTIVE",
      startedAt: now,
      lastChargedAt: now,
      nextChargeAt: addMonths(now, 1),
      lastPaymentKey: pay.paymentKey,
      lastOrderId: orderId,
      cancelAtPeriodEnd: false,
      canceledAt: null,
      endedAt: null,
    };

    await user.save();

    return res.json({ ok: true, subscription: user.subscription });
  } catch (e) {
    console.error("completeSubscription error:", e);
    return res.status(500).json({ error: "구독 결제 완료 처리 실패" });
  }
};

/**
 * POST /api/payments/subscription/cancel
 * - “기간말 해지(CANCELED_PENDING)” 기본
 * - 이미 해지 상태면 "이미 해지됨"으로 멱등 처리
 */
exports.cancelSubscription = async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select("plan subscription");
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    const sub = user.subscription || {};
    const status = String(sub.status || "NONE").toUpperCase();

    // ✅ 구독 자체가 없는 경우
    if (status === "NONE" || !sub.billingKey) {
      return res.status(400).json({
        ok: false,
        code: "NOT_SUBSCRIBED",
        message: "현재 구독 중이 아닙니다.",
        status,
      });
    }

    // ✅ 이미 해지된 경우(멱등)
    if (isAlreadyCanceled(sub)) {
      return res.json({
        ok: true,
        code: "ALREADY_CANCELED",
        message: "이미 구독 해지(또는 해지 예약) 상태입니다.",
        status,
      });
    }

    const now = new Date();

    // ✅ 기간말 해지: nextChargeAt(다음 결제일) 이후 자동 종료되도록
    user.subscription = {
      ...sub,
      status: "CANCELED_PENDING",
      cancelAtPeriodEnd: true,
      canceledAt: now,
      // nextChargeAt은 "이용 만료일"처럼 유지(너가 원한 UX)
    };

    await user.save();

    return res.json({
      ok: true,
      status: user.subscription.status,
      nextChargeAt: user.subscription.nextChargeAt,
      message: "구독 해지가 접수되었습니다. 이용 만료일까지 사용 가능합니다.",
    });
  } catch (e) {
    console.error("cancelSubscription error:", e);
    return res.status(500).json({ error: "구독 해지 실패" });
  }
};

/**
 * POST /api/payments/cron/charge
 * - ACTIVE 이고 nextChargeAt <= now 인 대상 월 과금
 * - CRON_SECRET 헤더로 보호 권장
 */
exports.cronCharge = async (req, res) => {
  try {
    const secret = req.headers["x-cron-secret"];
    if (CRON_SECRET && secret !== CRON_SECRET) {
      return res.status(401).json({ ok: false, error: "Unauthorized" });
    }

    const now = new Date();

    // 대상: ACTIVE + billingKey 있고 nextChargeAt 도래
    const targets = await User.find({
      "subscription.status": "ACTIVE",
      "subscription.billingKey": { $exists: true, $ne: "" },
      "subscription.nextChargeAt": { $lte: now },
    }).select("_id email name subscription plan");

    let successCount = 0;
    let failCount = 0;

    for (const user of targets) {
      const sub = user.subscription || {};
      try {
        const orderId = makeOrderId();

        // 결제 시도
        const charge = await toss.chargeWithBillingKey({
          billingKey: sub.billingKey,
          customerKey: String(user._id),
          orderId,
          amount: 4900,
          orderName: "BeeBee AI PRO (월 정기 결제)",
          customerEmail: user.email,
          customerName: user.name || user.email,
        });

        // Payment 기록(선택)
        await Payment.updateOne(
          { orderId },
          {
            $setOnInsert: {
              userId: String(user._id),
              orderId,
              createdAt: new Date(),
            },
            $set: {
              amount: 4900,
              currency: "KRW",
              provider: "toss",
              status: "DONE",
              paymentKey: charge.paymentKey || charge.raw?.paymentKey,
              raw: charge.raw || charge,
              approvedAt: new Date(charge.raw?.approvedAt || Date.now()),
              updatedAt: new Date(),
            },
          },
          { upsert: true }
        );

        // 구독 갱신: 다음 결제일 1개월
        const chargedAt = new Date();
        user.subscription = {
          ...sub,
          status: "ACTIVE",
          lastChargedAt: chargedAt,
          nextChargeAt: addMonths(chargedAt, 1),
          lastOrderId: orderId,
          lastPaymentKey: charge.paymentKey || charge.raw?.paymentKey,
        };
        user.plan = "PRO";
        await user.save();

        successCount++;
      } catch (err) {
        console.error("cronCharge user fail:", String(user._id), err);

        // 실패 처리
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
    });
  } catch (e) {
    console.error("cronCharge error:", e);
    return res.status(500).json({ ok: false, error: "cronCharge failed" });
  }
};
