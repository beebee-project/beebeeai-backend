const router = require("express").Router();
const adminAuth = require("../middleware/adminAuth");
const adminController = require("../controllers/adminController");

// ✅ GET /admin/summary?from=2026-01-01&to=2026-01-28&limit=20&reasonTopN=10
router.get("/summary", adminAuth, adminController.getAdminSummary);

router.get("/trace/:traceId", adminAuth, adminController.getTraceDetail);

module.exports = router;
