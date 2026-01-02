const tossClient = require("../../config/tossClient");
const CURRENCY = process.env.CURRENCY || "KRW";

exports.createCheckoutSession = async ({
  userId,
  amount,
  successUrl,
  failUrl,
  meta,
}) => {
  // 위젯은 “세션 생성 API”가 따로 있는 게 아니라,
  // 우리가 orderId/amount/successUrl/failUrl을 만들어서 프론트에 내려주면 됨.
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

// ✅ billingKey로 청구(묶음 C에서 사용): POST /v1/billing/{billingKey}
exports.chargeBillingKey = async ({
  customerKey,
  billingKey,
  amount,
  orderId,
  orderName,
}) => {
  const res = await tossClient.post(`/v1/billing/${billingKey}`, {
    customerKey,
    amount,
    orderId,
    orderName,
  });
  return res.data;
};
