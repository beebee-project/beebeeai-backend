const toss = require("../tossClient");

const CURRENCY = process.env.CURRENCY || "KRW";

/**
 * 결제 시작용 세션 생성
 * - 여기서는 토스 API 호출하지 않고, 우리 쪽 기준 orderId만 생성해서 프론트에 넘겨줌
 */
exports.createCheckoutSession = async ({
  userId,
  amount,
  orderId,
  successUrl,
  failUrl,
  meta = {},
}) => {
  const oid = orderId || `bb_${userId}_${Date.now()}`;

  return {
    provider: "toss",
    orderId: oid,
    amount,
    currency: CURRENCY,
    orderName: meta.orderName || "BeeBee AI PRO (월 정기 결제)",
    customerName: meta.customerName || `user:${userId}`,
    successUrl,
    failUrl,
  };
};

/**
 * 토스 결제 승인
 * - 프론트에서 받은 paymentKey, orderId, amount를 이용해
 *   Toss Payments /v1/payments/confirm 호출
 */
exports.confirmPayment = async ({ paymentKey, orderId, amount }) => {
  if (!paymentKey || !orderId || !amount) {
    const e = new Error("paymentKey, orderId, amount는 필수입니다.");
    e.code = "INVALID_CONFIRM_PARAMS";
    throw e;
  }

  const res = await toss.post("/v1/payments/confirm", {
    paymentKey,
    orderId,
    amount,
  }); // :contentReference[oaicite:2]{index=2}

  const payment = res.data;

  // 정상 완료된 결제만 허용
  if (payment.status !== "DONE") {
    const e = new Error(`결제 상태가 완료가 아닙니다: ${payment.status}`);
    e.code = "PAYMENT_NOT_DONE";
    e.payment = payment;
    throw e;
  }

  return {
    ok: true,
    provider: "toss",
    paymentKey: payment.paymentKey,
    orderId: payment.orderId,
    amount: payment.totalAmount ?? payment.amount ?? amount,
    raw: payment, // 필요하면 컨트롤러에서 DB 저장에 사용
  };
};
