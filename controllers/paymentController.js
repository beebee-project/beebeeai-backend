const paymentService = require("../services/paymentService");
const User = require("../models/User");
const tossClient = require("../config/tossClient");

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
        ? { formulaConversions: null, fileUploads: 5 }
        : { formulaConversions: 10, fileUploads: 1 };

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
    // 0) 베타모드면 결제 없이 PRO
    if (paymentService.isBetaMode()) {
      await User.findByIdAndUpdate(req.user.id, { plan: "PRO" });
      return res.json({
        ok: true,
        beta: true,
        plan: "PRO",
        message: "BETA_MODE=true: 결제 없이 PRO가 활성화되었습니다.",
      });
    }

    // ✅ [추가] 1) 유저/구독 상태 확인 (중복 구독 방지 + 해지 후 기간 종료 시 재구독 허용)
    const user = await User.findById(req.user.id).select("plan subscription");
    if (!user) return res.status(404).json({ error: "User not found" });

    const now = new Date();
    const sub = user.subscription || {};

    // 1-1) 이미 활성/연체면 재구독(중복) 차단
    if (["ACTIVE", "PAST_DUE"].includes(sub.status)) {
      return res.status(409).json({
        ok: false,
        error: "이미 구독 중입니다.",
        code: "SUBSCRIPTION_ALREADY_ACTIVE",
        status: sub.status,
      });
    }

    // 1-2) 해지 예약 상태면:
    // - 아직 기간 남아있으면 차단
    // - 기간 끝났으면 만료 확정 처리 후 재구독 허용
    if (sub.status === "CANCELED_PENDING" && sub.cancelAtPeriodEnd) {
      const endAt = sub.nextChargeAt ? new Date(sub.nextChargeAt) : null;

      // 기간이 남아있으면 차단
      if (endAt && now < endAt) {
        return res.status(409).json({
          ok: false,
          error:
            "해지 예약 상태입니다. 남은 이용기간 종료 후 재구독할 수 있습니다.",
          code: "SUBSCRIPTION_CANCEL_PENDING",
          status: sub.status,
          endsAt: endAt.toISOString(),
        });
      }

      // ✅ 기간이 끝났으면 만료 확정 (billingKey는 유지)
      await User.updateOne(
        { _id: user._id },
        {
          $set: {
            plan: "FREE",
            "subscription.status": "CANCELED",
            "subscription.cancelAtPeriodEnd": false,
          },
          // billingKey/customerKey는 유지 (A 정책)
        }
      );
    }

    // ✅ [기존] 2) customerKey 생성
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

    // ✅ [기존] 3) 프론트가 위젯 호출에 필요한 값 반환
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
  // ✅ 30일 만료 후 완전 삭제(Soft delete -> Hard delete)
  const cutoff = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);

  const purgeResult = await User.deleteMany({
    isDeleted: true,
    deletedAt: { $ne: null, $lte: cutoff },
  });

  console.log("[cronCharge] purged users:", purgeResult.deletedCount ?? 0);

  // 베타모드면 청구/정리 모두 스킵(원하면 정리만 하게 변경 가능)
  if (paymentService.isBetaMode()) {
    return res.json({ ok: true, skipped: true, reason: "BETA_MODE=true" });
  }

  const now = new Date();

  function kstYYYYMMDD(date) {
    const dtf = new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Seoul",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    });
    return dtf.format(date).replaceAll("-", ""); // "YYYYMMDD"
  }

  try {
    console.log("[cronCharge] hit", now.toISOString());

    // 1) 만료 확정 처리 (✅ 중복 제거 + 보안검증 이후 실행)
    // - 해지 예약인데 기간이 끝난 유저를 FREE + CANCELED로 확정
    // - A정책: billingKey/customerKey는 유지
    const expiredCanceled = await User.updateMany(
      {
        "subscription.status": "CANCELED_PENDING",
        "subscription.cancelAtPeriodEnd": true,
        "subscription.nextChargeAt": { $ne: null, $lte: now },
      },
      {
        $set: {
          plan: "FREE",
          "subscription.status": "CANCELED",
          "subscription.cancelAtPeriodEnd": false,
          "subscription.endedAt": now,
          // 정책적으로 nextChargeAt은 "더 이상 청구 없음"이므로 null로 정리 추천
          "subscription.nextChargeAt": null,
        },
      }
    );

    // 2) 구독 청구 금액/상품명
    const amount = Number(process.env.SUBSCRIPTION_AMOUNT || 4900);
    const orderName = process.env.SUBSCRIPTION_ORDER_NAME || "BeeBee AI PRO";

    // 3) 청구 대상 조회: ACTIVE/PAST_DUE + nextChargeAt 도래 + billingKey 존재
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

      // 이번 청구 회차 기준: nextChargeAt
      const due = u.subscription?.nextChargeAt
        ? new Date(u.subscription.nextChargeAt)
        : now;

      const periodKey = kstYYYYMMDD(due);
      const orderId = `sub-${userId}-${periodKey}`;

      // ✅ 이미 같은 회차 성공 처리된 경우 스킵 (DB 기반 2중 방지)
      if (
        u.subscription?.lastChargeKey === periodKey &&
        u.subscription?.lastChargedAt
      ) {
        continue;
      }

      // ✅ 동시 실행 방지 락
      const lock = await User.updateOne(
        {
          _id: u._id,
          // 같은 periodKey로 이미 락이 잡혀있으면 못 들어오게
          "subscription.chargeLockKey": { $ne: periodKey },
          "subscription.nextChargeAt": { $ne: null, $lte: now },
          "subscription.status": { $in: ["ACTIVE", "PAST_DUE"] },
        },
        {
          $set: {
            "subscription.chargeLockKey": periodKey,
            "subscription.lastOrderId": orderId,
            "subscription.lastChargeAttemptAt": now,
          },
        }
      );

      if (!lock?.modifiedCount) continue;

      try {
        await paymentService.chargeBillingKey({
          customerKey,
          billingKey,
          amount,
          orderId,
          orderName,
          idempotencyKey: orderId, // ✅ toss.js에서 Idempotency-Key 헤더로 사용
        });

        // ✅ 다음 청구일은 now가 아니라 due 기준으로 (드리프트 방지)
        const nextChargeAt = paymentService.addMonths(due, 1);

        await User.updateOne(
          { _id: u._id },
          {
            $set: {
              "subscription.status": "ACTIVE",
              "subscription.lastChargedAt": now,
              "subscription.nextChargeAt": nextChargeAt,
              "subscription.customerKey": customerKey,
              "subscription.lastChargeKey": periodKey,
            },
            $unset: {
              "subscription.lastChargeError": "",
              "subscription.chargeLockKey": "",
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
            $unset: {
              "subscription.chargeLockKey": "",
            },
          }
        );
      }
    }

    return res.json({
      ok: true,
      now,
      cleaned: {
        expiredCanceled:
          expiredCanceled?.modifiedCount ?? expiredCanceled?.nModified ?? 0,
      },
      targets: targets.length,
      successCount,
      failCount,
      purgedDeletedUsers: purgeResult.deletedCount ?? 0,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Cron charge failed" });
  }
};

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

exports.webhook = async (req, res) => {
  try {
    // 1) 빠르게 ACK (PG 재시도 방지)
    //    -> 여기선 바로 응답하지 않고, 아래 로직까지 처리 후 200 줘도 됨.
    const body = req.body || {};

    // 2) 토스 웹훅 payload는 형태가 다양할 수 있어서 방어적으로 추출
    const paymentKey =
      body.paymentKey || body.data?.paymentKey || body.resource?.paymentKey;
    const orderId =
      body.orderId || body.data?.orderId || body.resource?.orderId;
    const customerKey =
      body.customerKey || body.data?.customerKey || body.resource?.customerKey;

    // billingKey 이벤트가 따로 오는 경우 대비 (없으면 무시)
    const billingKey =
      body.billingKey || body.data?.billingKey || body.resource?.billingKey;

    // paymentKey도 billingKey도 없으면 일단 OK
    if (!paymentKey && !billingKey) {
      return res.json({ ok: true, ignored: true });
    }

    // 3) 사용자 식별
    let userId = null;

    if (customerKey) userId = String(customerKey);

    // cronCharge에서 만든 orderId는 sub-<userId>-<ts> 형태 :contentReference[oaicite:6]{index=6}
    if (!userId && typeof orderId === "string" && orderId.startsWith("sub-")) {
      const parts = orderId.split("-");
      if (parts.length >= 3) userId = parts[1];
    }

    if (!userId) {
      // 식별 못 해도 200은 주되 로그만 남김
      console.log("[webhook] cannot resolve userId", { orderId, customerKey });
      return res.json({ ok: true, unresolved: true });
    }

    const user = await User.findById(userId).select("plan subscription");
    if (!user) return res.json({ ok: true, noUser: true });

    // 4) 중복 방지(idempotent): 이미 처리한 paymentKey면 바로 OK
    if (paymentKey && user.subscription?.lastPaymentKey === paymentKey) {
      return res.json({ ok: true, duplicate: true });
    }

    // 5) 토스에 조회해서 상태 확정 (payload를 믿지 않음)
    let payment = null;
    if (paymentKey) {
      const r = await tossClient.get(`/v1/payments/${paymentKey}`);
      payment = r.data;
    }

    // 6) DB 반영 규칙 (MVP)
    // - DONE: PRO 유지/ACTIVE 유지 + lastPaymentKey/lastOrderId 갱신
    // - 그 외: PAST_DUE로 내려서 결제 실패 상태 표시(운영 정책에 따라 조정 가능)
    if (payment) {
      const status = String(payment.status || "").toUpperCase();

      user.subscription = { ...(user.subscription || {}) };
      user.subscription.lastPaymentKey =
        paymentKey || user.subscription.lastPaymentKey;
      user.subscription.lastOrderId =
        payment.orderId || orderId || user.subscription.lastOrderId;

      if (status === "DONE") {
        user.plan = "PRO";
        user.subscription.status = "ACTIVE";
        user.subscription.lastChargedAt = new Date();
      } else {
        user.subscription.status = "PAST_DUE";
        user.subscription.lastChargeError = `webhook status=${status}`;
      }

      await user.save();
    }

    return res.json({ ok: true });
  } catch (e) {
    console.error("[webhook] error", e);
    // PG는 2xx 아니면 재시도할 수 있으니, MVP에서는 200으로 받고 내부 로깅으로 확인하는 것도 방법
    return res.status(200).json({ ok: true, error: true });
  }
};
