const tossClient = require("../../config/tossClient");

// 1) 빌링 결제 시작(결제창 열기용 payload 생성)
async function createBillingCheckoutSession({
  customerKey,
  orderId,
  amount,
  successUrl,
  failUrl,
  orderName,
  customerName,
}) {
  // 여기서 Toss 결제창으로 넘길 값 반환(프론트에서 toss SDK로 호출)
  return {
    provider: "toss",
    paymentType: "BILLING",
    customerKey,
    orderId,
    amount,
    orderName,
    customerName,
    successUrl,
    failUrl,
  };
}

// 2) authKey -> billingKey 발급
async function issueBillingKey({ customerKey, authKey }) {
  // 예시: tossClient.issueBillingKey(...)
  // return { billingKey, customerKey, raw }
  return tossClient.issueBillingKey({ customerKey, authKey });
}

// 3) billingKey로 첫 결제(또는 정기결제) 승인(자동결제)
async function chargeWithBillingKey({
  billingKey,
  customerKey,
  orderId,
  amount,
  orderName,
  customerEmail,
  customerName,
}) {
  return tossClient.chargeBillingKey({
    billingKey,
    customerKey,
    orderId,
    amount,
    orderName,
    customerEmail,
    customerName,
  });
}

module.exports = {
  createBillingCheckoutSession,
  issueBillingKey,
  chargeWithBillingKey,
};
