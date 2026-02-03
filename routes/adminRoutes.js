const router = require("express").Router();
const adminAuth = require("../middleware/adminAuth");
const adminController = require("../controllers/adminController");

router.get("/summary", adminAuth, adminController.getAdminSummary);
router.get("/trace/:traceId", adminAuth, adminController.getTraceDetail);
router.get("/daily-summary", adminAuth, adminController.getDailySummary);
router.get("/trends", adminAuth, adminController.getTrends);

module.exports = router;
