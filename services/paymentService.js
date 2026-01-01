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

function addMonths(date, months) {
  const d = new Date(date);
  const day = d.getDate();
  d.setMonth(d.getMonth() + months);

  // 월말 보정(예: 1/31 + 1month => 3/03 같은 문제 방지)
  if (d.getDate() < day) d.setDate(0);
  return d;
}

// "이미 구독/체험/해지예약 중" 여부 (프론트/컨트롤러에서 쓰기 좋게 별칭 제공)
function isSubscriptionActive(sub = {}) {
  const status = String(sub.status || "NONE").toUpperCase();
  return ["TRIAL", "ACTIVE", "PAST_DUE", "CANCELED_PENDING"].includes(status);
}

// 해지 불필요(이미 해지 상태) 판정
function isAlreadyCanceled(sub = {}) {
  const status = String(sub.status || "NONE").toUpperCase();
  return ["CANCELED", "CANCELED_PENDING"].includes(status);
}

module.exports = {
  selectGateway,
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
