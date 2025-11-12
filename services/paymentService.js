const free = require("./paymentGateway/freeBeta");
const sprite = require("./paymentGateway/sprite");

function selectGateway() {
  switch (String(process.env.PG_PROVIDER).toLowerCase()) {
    case "sprite":
      return sprite;
    default:
      return free;
  }
}
const pg = selectGateway();

exports.createCheckoutSession = (args) => pg.createCheckoutSession(args);
exports.confirmPayment = (args) => pg.confirmPayment(args);
exports.cancelPayment = (args) => pg.cancelPayment?.(args);
exports.parseAndVerifyWebhook = (req) => pg.parseAndVerifyWebhook?.(req);
