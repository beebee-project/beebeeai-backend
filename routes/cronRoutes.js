const router = require("express").Router();
const cronMiddleware = require("../middleware/cronMiddleware");
const cronController = require("../controllers/cronController");

router.post("/daily-summary", cronMiddleware, cronController.runDailySummary);

module.exports = router;
