const free = require("./paymentGateway/freeBeta");
const gateway = require("./paymentGateway/toss");

function getProvider() {
  if (process.env.LOCAL_DEV === "1") return "freebeta";
  return String(process.env.PG_PROVIDER || "toss").toLowerCase();
}

function getCurrency() {
  return process.env.CURRENCY || "KRW";
}

function isBetaMode() {
  return String(process.env.BETA_MODE).toLowerCase() === "true";
}

function getEffectivePlan(userPlan) {
  // 베타면 결제 없이 PRO
  if (isBetaMode()) return "PRO";
  return userPlan || "FREE";
}

function addMonths(date, months) {
  const d = new Date(date);
  const day = d.getDate();

  d.setMonth(d.getMonth() + months);

  // 말일 보정 (예: 1/31 → 2/28 or 2/29)
  if (d.getDate() < day) d.setDate(0);
  return d;
}

function selectGateway() {
  switch (getProvider()) {
    case "toss":
      return gateway;
    default:
      return free;
  }
}

function isSubscriptionActive(sub = {}, now = new Date()) {
  const status = String(sub?.status || "").toUpperCase();
  const lockedStatuses = ["ACTIVE", "PAST_DUE", "CANCELED_PENDING"];

  if (!lockedStatuses.includes(status)) return false;

  if (!sub?.nextChargeAt) return false;

  const expiresAt = new Date(sub.nextChargeAt);
  if (Number.isNaN(expiresAt.getTime())) return false;

  return expiresAt > now;
}

exports.isSubscriptionActive = isSubscriptionActive;
exports.createCheckoutSession = (args) =>
  selectGateway().createCheckoutSession(args);
exports.confirmPayment = (args) => selectGateway().confirmPayment(args);
exports.cancelPayment = (args) => selectGateway().cancelPayment?.(args);
exports.parseAndVerifyWebhook = (req) =>
  selectGateway().parseAndVerifyWebhook?.(req);
exports.getProvider = getProvider;
exports.getCurrency = getCurrency;
exports.isBetaMode = isBetaMode;
exports.getEffectivePlan = getEffectivePlan;
exports.addMonths = addMonths;

// 구독(빌링키) 기능
exports.issueBillingKey = (args) => selectGateway().issueBillingKey?.(args);
exports.chargeBillingKey = async ({
  billingKey,
  customerKey,
  amount,
  orderId,
  orderName,
  idempotencyKey,
}) => {
  return gateway.chargeBillingKey({
    billingKey,
    customerKey,
    amount,
    orderId,
    orderName,
    idempotencyKey,
  });
};
