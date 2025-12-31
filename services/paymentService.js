const toss = require("./paymentGateway/toss");
const freeBeta = require("./paymentGateway/freeBeta");

// BETA_MODE=true면 결제 게이트웨이 대신 freeBeta 사용
function selectGateway() {
  const beta =
    String(process.env.BETA_MODE || "false").toLowerCase() === "true";
  const provider = String(process.env.PG_PROVIDER || "toss").toLowerCase();
  if (beta) return freeBeta;
  if (provider === "toss") return toss;
  return toss;
}

// ✅ 운영/테스트에서 날짜 강제이동용(선택)
function getNow() {
  const v = process.env.PAYMENTS_NOW;
  if (!v) return new Date();
  const d = new Date(v);
  return isNaN(d.getTime()) ? new Date() : d;
}

function addMonths(date, months) {
  const d = new Date(date);
  const day = d.getDate();
  d.setMonth(d.getMonth() + months);

  // 월말 보정(예: 1/31 + 1month => 3/03 같은 문제 방지)
  if (d.getDate() < day) d.setDate(0);
  return d;
}

// 구독 중복 락 판정: "다시 구독 시작" 막는 기준
function isSubscriptionLocked(sub = {}) {
  const status = String(sub.status || "NONE").toUpperCase();
  return ["ACTIVE", "PAST_DUE", "CANCELED_PENDING"].includes(status);
}

// 해지 불필요(이미 해지 상태) 판정
function isAlreadyCanceled(sub = {}) {
  const status = String(sub.status || "NONE").toUpperCase();
  return ["CANCELED", "CANCELED_PENDING"].includes(status);
}

module.exports = {
  selectGateway,
  getNow,
  isSubscriptionLocked,
  isSubscriptionActive,
  addMonths,
  isAlreadyCanceled,

  // 아래는 게이트웨이로 그대로 위임
  createCheckoutSession: (args) => selectGateway().createCheckoutSession(args),
  confirmPayment: (args) => selectGateway().confirmPayment(args),

  // ✅ 정기결제(빌링키) 플로우
  startSubscription: (args) => selectGateway().startSubscription(args),
  completeSubscription: (args) => selectGateway().completeSubscription(args),
  chargeBillingKey: (args) => selectGateway().chargeBillingKey(args),
};
