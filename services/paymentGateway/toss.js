const axios = require("axios");
const tossClient = require("../../config/tossClient");

const SECRET_KEY = process.env.TOSS_SECRET_KEY;
const CURRENCY = process.env.CURRENCY || "KRW";

function getAuthHeader() {
  if (!SECRET_KEY) throw new Error("TOSS_SECRET_KEY is missing");
  const encoded = Buffer.from(`${SECRET_KEY}:`, "utf8").toString("base64");
  return `Basic ${encoded}`;
}

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
  const r = await axios.post(
    "https://api.tosspayments.com/v1/payments/confirm",
    { paymentKey, orderId, amount },
    {
      headers: {
        Authorization: getAuthHeader(),
        "Content-Type": "application/json",
      },
      timeout: 15000,
    }
  );

  const data = r.data;

  // DONE이 아니면 실패 처리(원하면 더 세밀하게)
  if (data.status && data.status !== "DONE") {
    const err = new Error(`Payment not DONE: ${data.status}`);
    err.code = "TOSS_NOT_DONE";
    err.response = { data };
    throw err;
  }

  return {
    provider: "toss",
    orderId: data.orderId,
    amount: data.totalAmount ?? amount,
    paymentKey: data.paymentKey,
    raw: data,
  };
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
