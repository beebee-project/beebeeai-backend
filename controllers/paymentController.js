const paymentService = require("../services/paymentService");
const User = require("../models/User");

const PROVIDER = String(process.env.PG_PROVIDER || "toss").toLowerCase();
const CURRENCY = process.env.CURRENCY || "KRW";
const isBetaMode = () => String(process.env.BETA_MODE).toLowerCase() === "true";

function ensureAbsoluteUrl(url, fallbackOrigin) {
  // url이 "www.xxx" 같이 스킴 없이 들어오면 fallbackOrigin 붙여서 보정
  // url이 "/path"면 fallbackOrigin + url
  if (!url) return null;

  const trimmed = String(url).trim();

  // 이미 절대 URL이면 그대로
  if (/^https?:\/\//i.test(trimmed)) return trimmed;

  // "www.beebeeai.kr/..." 형태면 https:// 붙여주기
  if (/^www\./i.test(trimmed)) return `https://${trimmed}`;

  // "/success.html" 형태면 origin 붙여주기
  if (trimmed.startsWith("/")) return `${fallbackOrigin}${trimmed}`;

  // 그 외는 origin + "/" + url
  return `${fallbackOrigin}/${trimmed}`;
}

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
    return res.status(410).json({
      error:
        "PRO 결제는 정기결제(구독)로만 가능합니다. 구독 시작을 이용해주세요.",
      code: "CHECKOUT_DEPRECATED",
    });
  } catch (err) {
    console.error("createCheckout error:", err);
    res.status(500).json({ error: "결제 세션 생성 실패", code: err.code });
  }
};

// 결제 승인
exports.confirmPayment = async (req, res) => {
  try {
    return res.status(410).json({
      error:
        "PRO 결제는 정기결제(구독)로만 가능합니다. 구독 시작을 이용해주세요.",
      code: "CHECKOUT_DEPRECATED",
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

    // 기준 origin (환경변수로 통일 권장)
    const origin =
      (process.env.PUBLIC_ORIGIN &&
        ensureAbsoluteUrl(process.env.PUBLIC_ORIGIN, "https://beebeeai.kr")) ||
      "https://beebeeai.kr";

    // env에서 받은 URL을 '절대 URL'로 강제 보정
    const successUrl = ensureAbsoluteUrl(
      process.env.SUBSCRIPTION_SUCCESS_URL || `${origin}/success.html`,
      origin
    );
    const failUrl = ensureAbsoluteUrl(
      process.env.SUBSCRIPTION_FAIL_URL || `${origin}/fail.html`,
      origin
    );

    // 최종 형식 검증: 여기서 걸리면 Toss로 보내기 전에 서버가 막아줌
    try {
      new URL(successUrl);
      new URL(failUrl);
    } catch (e) {
      console.error("Invalid subscription URLs:", { successUrl, failUrl });
      return res.status(500).json({
        error: "Invalid subscription success/fail URL",
        successUrl,
        failUrl,
      });
    }

    return res.json({
      ok: true,
      beta: false,
      customerKey,
      successUrl,
      failUrl,
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

    // 구독 등록 완료: 즉시 ACTIVE, 다음 청구일은 1개월 후
    const now = new Date();
    const nextChargeAt = paymentService.addMonths(now, 1);

    const user = await User.findById(req.user.id);
    if (!user) return res.status(404).json({ error: "사용자 없음" });

    user.plan = "PRO";
    user.subscription = {
      ...(user.subscription || {}),
      customerKey,
      billingKey: issued.billingKey,
      status: "ACTIVE",
      startedAt: user.subscription?.startedAt || now,
      trialEndsAt: null,
      nextChargeAt,
      lastChargedAt: null,
    };

    await user.save();

    return res.json({
      ok: true,
      plan: "PRO",
      trialEndsAt: null,
      nextChargeAt,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "구독 완료 처리 실패" });
  }
};

exports.cronCharge = async (req, res) => {
  try {
    console.log("[cronCharge] hit", new Date().toISOString());

    const secret = req.headers["x-cron-secret"];
    if (!process.env.CRON_SECRET || secret !== process.env.CRON_SECRET) {
      return res.status(401).json({ error: "Unauthorized cron" });
    }

    // 베타모드면 청구/정리 모두 스킵(원하면 정리만 수행하도록 바꿔도 됨)
    if (paymentService.isBetaMode()) {
      return res.json({ ok: true, skipped: true, reason: "BETA_MODE=true" });
    }

    const now = new Date();

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
        "subscription.status": { $in: ["ACTIVE", "PAST_DUE"] },
      },
      "_id subscription"
    ).lean();

    console.log("[cronCharge] targets", targets.length);

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
