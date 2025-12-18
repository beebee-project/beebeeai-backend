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

exports.createCheckoutSession = (args) =>
  selectGateway().createCheckoutSession(args);
exports.confirmPayment = (args) => selectGateway().confirmPayment(args);
exports.cancelPayment = (args) => selectGateway().cancelPayment?.(args);
exports.parseAndVerifyWebhook = (req) =>
  selectGateway().parseAndVerifyWebhook?.(req);
exports.isBetaMode = isBetaMode;
exports.getEffectivePlan = getEffectivePlan;
exports.addMonths = addMonths;
