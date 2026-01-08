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
    amount,
    orderId,
    orderName,
    currency: CURRENCY,
  };

  const config = {};
  if (idempotencyKey) {
    config.headers = {
      "Idempotency-Key": String(idempotencyKey),
    };
  }

  const res = await tossClient.post(url, payload, config);
  return res.data;
};
