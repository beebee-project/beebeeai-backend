const express = require("express");
const router = express.Router();

const adminAuth = require("../middleware/adminAuth");
const adminController = require("../controllers/adminController");

// 기존 관리자 요약(요청 로그 기반)
router.get("/summary", adminAuth, adminController.getAdminSummary);

// DailySummary 조회/트렌드
router.get("/daily-summaries", adminAuth, adminController.getDailySummaries);
router.get(
  "/daily-summaries/:day",
  adminAuth,
  adminController.getDailySummaryByDay,
);
router.get("/daily-trends", adminAuth, adminController.getDailyTrends);

module.exports = router;
