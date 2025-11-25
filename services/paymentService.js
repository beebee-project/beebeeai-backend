const free = require("./paymentGateway/freeBeta");
const toss = require("./paymentGateway/toss");

function selectGateway() {
  switch (String(process.env.PG_PROVIDER).toLowerCase()) {
    case "toss":
      return toss;
    default:
      return free;
  }
}

const pg = selectGateway();

exports.createCheckoutSession = (args) => pg.createCheckoutSession(args);
exports.confirmPayment = (args) => pg.confirmPayment(args);
exports.cancelPayment = (args) => pg.cancelPayment?.(args);
exports.parseAndVerifyWebhook = (req) => pg.parseAndVerifyWebhook?.(req);
