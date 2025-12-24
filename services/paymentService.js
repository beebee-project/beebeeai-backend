const free = require("./paymentGateway/freeBeta");
const toss = require("./paymentGateway/toss");

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
  switch (String(process.env.PG_PROVIDER).toLowerCase()) {
    case "toss":
      return toss;
    default:
      return free;
  }
}

function isSubscriptionActive(sub = {}, now = new Date()) {
  const status = String(sub?.status || "").toUpperCase();

  // 결제 시작/재구독을 막아야 하는 상태들
  const lockedStatuses = ["TRIAL", "ACTIVE", "PAST_DUE", "CANCELED_PENDING"];

  if (lockedStatuses.includes(status)) return true;

  // (선택) status가 비어있더라도 날짜가 미래면 잠금 처리하고 싶으면 아래 활성화
  // if (sub?.trialEndsAt && new Date(sub.trialEndsAt) > now) return true;
  // if (sub?.nextChargeAt && new Date(sub.nextChargeAt) > now) return true;

  return false;
}

exports.isSubscriptionActive = isSubscriptionActive;
exports.createCheckoutSession = (args) =>
  selectGateway().createCheckoutSession(args);
exports.confirmPayment = (args) => selectGateway().confirmPayment(args);
exports.cancelPayment = (args) => selectGateway().cancelPayment?.(args);
exports.parseAndVerifyWebhook = (req) =>
  selectGateway().parseAndVerifyWebhook?.(req);
exports.isBetaMode = isBetaMode;
exports.getEffectivePlan = getEffectivePlan;
exports.addMonths = addMonths;

// ✅ 구독(빌링키) 기능
exports.issueBillingKey = (args) => selectGateway().issueBillingKey?.(args);
exports.chargeBillingKey = (args) => selectGateway().chargeBillingKey?.(args);
