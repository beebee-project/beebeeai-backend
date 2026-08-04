const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const automationController = require("../controllers/automationController");
const {
  executeBusinessTemplateObserved,
} = require("../automation/semanticExecutionRouteBridge");
const {
  createQueryCandidatePlannerApiShadowBoundary,
} = require("../automation/queryCandidatePlannerApiShadowBoundary");
const {
  createQueryCandidatePlannerDownloadRetentionBoundary,
} = require("../automation/queryCandidatePlannerFileLifecycleBoundary");

router.use(protect);

const getAnalysisCandidatesShadowObserved =
  createQueryCandidatePlannerApiShadowBoundary({
    handler: automationController.getAnalysisCandidates,
  });
const downloadGeneratedFileCacheRetained =
  createQueryCandidatePlannerDownloadRetentionBoundary({
    handler: automationController.downloadGeneratedFile,
    action: "GENERATED_DOWNLOAD",
  });

router.post("/query-preview", automationController.previewQueryTables);
router.post("/query-save", automationController.saveQueryTables);
router.post("/analysis-candidates", getAnalysisCandidatesShadowObserved);
router.post("/query-analyze", automationController.analyzeQueryIntent);
router.post("/query-execute", automationController.executeQuery);
router.post("/export-xlsx", automationController.exportXlsx);
router.post("/summary-sheet", automationController.createSummarySheet);
router.get("/download", downloadGeneratedFileCacheRetained);
router.post("/export-report-json", automationController.exportReportJson);
router.post(
  "/export-analysis-report",
  automationController.exportAnalysisReport,
);
router.post("/export-pptx", automationController.exportPptx);
router.post(
  "/execute-analysis-candidate",
  automationController.executeAnalysisCandidate,
);
router.post("/execute-business-template", executeBusinessTemplateObserved);

module.exports = router;
