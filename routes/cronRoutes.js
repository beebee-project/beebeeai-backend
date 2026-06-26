const router = require("express").Router();

// Legacy DailySummary cron route removed.
// Subscription billing cron is handled by routes/paymentRoutes.js and scripts/runSubscriptionChargeCron.js.
module.exports = router;
