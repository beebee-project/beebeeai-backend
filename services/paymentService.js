const free = require("./paymentGateway/freeBeta"); // 베타/free 모드용
const toss = require("./paymentGateway/toss"); // 토스 결제 게이트웨이

function selectGateway() {
  switch (String(process.env.PG_PROVIDER).toLowerCase()) {
    case "toss":
      return toss;
    default:
      // 기본값은 무료 게이트웨이 (또는 필요 없으면 여기서도 toss로)
      return free;
  }
}

const pg = selectGateway();

exports.createCheckoutSession = (args) => pg.createCheckoutSession(args);
exports.confirmPayment = (args) => pg.confirmPayment(args);
exports.cancelPayment = (args) => pg.cancelPayment?.(args);
exports.parseAndVerifyWebhook = (req) => pg.parseAndVerifyWebhook?.(req);
