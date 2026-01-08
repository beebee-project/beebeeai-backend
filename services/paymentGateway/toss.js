const tossClient = require("../../config/tossClient");
const CURRENCY = process.env.CURRENCY || "KRW";

exports.createCheckoutSession = async ({
  userId,
  amount,
  successUrl,
  failUrl,
  meta,
}) => {
  const orderId = `beebeeai-${Date.now()}-${String(userId).slice(-4)}`;

  return {
    provider: "toss",
    orderId,
    amount,
    currency: CURRENCY,
    orderName: meta?.orderName,
    customerName: meta?.customerName,
    successUrl,
    failUrl,
  };
};

exports.confirmPayment = async ({ paymentKey, orderId, amount }) => {
  const res = await tossClient.post("/v1/payments/confirm", {
    paymentKey,
    orderId,
    amount,
  });
  return res.data;
};

// ✅ billingKey 발급: authKey -> billingKey
// Toss Billing API: POST /v1/billing/authorizations/{authKey}
exports.issueBillingKey = async ({ customerKey, authKey }) => {
  const res = await tossClient.post(`/v1/billing/authorizations/${authKey}`, {
    customerKey,
  });
  return res.data; // { billingKey, customerKey, ... }
};

exports.getPaymentByOrderId = async (orderId) => {
  // Toss Payments: 주문ID로 결제 조회
  // GET /v1/payments/orders/{orderId}
  const url = "/v1/payments/orders/" + encodeURIComponent(orderId);
  const res = await tossClient.get(url);
  return res.data;
};

// ✅ billingKey로 청구: POST /v1/billing/{billingKey}
exports.chargeBillingKey = async ({
  billingKey,
  customerKey,
  amount,
  orderId,
  orderName,
  idempotencyKey,
}) => {
  const url = "/v1/billing/" + encodeURIComponent(billingKey);

  const payload = {
    customerKey,
    amount: Number(amount),
    orderId: String(orderId),
    orderName: String(orderName),
    currency: CURRENCY,
  };

  const config = {
    headers: {},
  };

  if (idempotencyKey) {
    config.headers["Idempotency-Key"] = String(idempotencyKey);
  }

  try {
    const res = await tossClient.post(url, payload, config);
    return res.data;
  } catch (e) {
    // ✅ 중복/재시도/타임아웃 등으로 "이미 처리됨"일 수 있으니,
    // orderId로 결제 조회해서 DONE이면 성공으로 확정한다.
    const status = e?.response?.status;
    const data = e?.response?.data;

    // Toss가 중복 요청을 어떤 코드로 주든(환경/버전에 따라 다를 수 있음) 안전하게 조회로 확정
    if (
      orderId &&
      (status === 409 || status === 400 || status === 429 || status >= 500)
    ) {
      try {
        const found = await exports.getPaymentByOrderId(orderId);
        // DONE이면 이번 청구는 "이미 성공" 처리로 봐도 됨
        if (found?.status === "DONE") return found;
      } catch (_) {
        // 조회도 실패하면 원래 에러 던짐
      }
    }

    throw e;
  }
};
