const axios = require("axios");
const crypto = require("crypto");

const BASE = process.env.SPRITE_API_BASE;
const MID = process.env.SPRITE_MERCHANT_ID;
const SK = process.env.SPRITE_SECRET_KEY;
const WH = process.env.SPRITE_WEBHOOK_SECRET;
const CURRENCY = process.env.CURRENCY || "KRW";

exports.createCheckoutSession = async ({
  userId,
  amount,
  orderId,
  successUrl,
  failUrl,
  meta = {},
}) => {
  const oid = orderId || `u${userId}_${Date.now()}`;
  const r = await axios.post(
    `${BASE}/v1/payments`,
    {
      merchantId: MID,
      orderId: oid,
      amount,
      currency: CURRENCY,
      successUrl,
      failUrl,
      metadata: { userId, ...meta },
    },
    { headers: { Authorization: `Bearer ${SK}` } }
  );

  return {
    provider: "sprite",
    orderId: oid,
    checkoutUrl: r.data.checkoutUrl,
    successUrl,
    failUrl,
  };
};

exports.confirmPayment = async ({ orderId, expectedAmount }) => {
  const r = await axios.get(
    `${BASE}/v1/payments/by-order/${encodeURIComponent(orderId)}`,
    {
      headers: { Authorization: `Bearer ${SK}` },
    }
  );
  const p = r.data; // { status:'paid', amount, currency, paymentId, ... }
  if (p.status !== "paid") {
    const e = new Error("결제 미완료");
    e.code = "NOT_PAID";
    throw e;
  }
  if (Number(p.amount) !== Number(expectedAmount)) {
    const e = new Error("금액 불일치");
    e.code = "AMOUNT_MISMATCH";
    throw e;
  }
  if (p.currency !== CURRENCY) {
    const e = new Error("통화 불일치");
    e.code = "CURRENCY_MISMATCH";
    throw e;
  }
  return { ok: true, providerPaymentId: p.paymentId };
};

exports.parseAndVerifyWebhook = (req) => {
  const sig = req.header("Sprite-Signature");
  const raw = JSON.stringify(req.body);
  const expect = crypto.createHmac("sha256", WH).update(raw).digest("hex");
  if (sig !== expect) {
    const e = new Error("invalid signature");
    e.code = "INVALID_SIGNATURE";
    throw e;
  }
  return req.body;
};
