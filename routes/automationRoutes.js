const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const automationController = require("../controllers/automationController");

router.use(protect);

router.post("/query-preview", automationController.previewQueryTables);
router.post("/query-save", automationController.saveQueryTables);
router.post("/query-analyze", automationController.analyzeQueryIntent);
router.post("/query-execute", automationController.executeQuery);
router.post("/export-xlsx", automationController.exportXlsx);
router.post("/summary-sheet", automationController.createSummarySheet);

module.exports = router;
