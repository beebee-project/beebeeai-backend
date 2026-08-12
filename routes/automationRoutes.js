const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const automationController = require("../controllers/automationController");
const {
  executeBusinessTemplateObserved,
} = require("../automation/semanticExecutionRouteBridge");
const {
  defaultObservationLogger,
} = require("../automation/queryCandidatePlannerApiShadowBoundary");
const {
  createQueryCandidatePlannerInternalAllowlistCanaryBoundary,
} = require("../automation/queryCandidatePlannerInternalAllowlistCanaryBoundary");
const {
  createQueryCandidatePlannerDownloadRetentionBoundary,
} = require("../automation/queryCandidatePlannerFileLifecycleBoundary");
const {
  recordQueryCandidatePlannerInternalPreviewObservation,
} = require("../automation/queryCandidatePlannerInternalPreviewStore");
const {
  runQueryCandidatePlannerApiShadowWithRealEvidenceCapture,
} = require("../automation/queryCandidatePlannerRealShadowCaptureBridge");
const {
  recordQueryCandidatePlannerRealShadowObservation,
  recordQueryCandidatePlannerRealShadowLifecycleObservation,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceCollector");
const {
  requireQueryCandidatePlannerInternalPreviewAccess,
} = require("../automation/queryCandidatePlannerInternalPreviewAccess");
const {
  internalPreviewPage,
  internalPreviewStatus,
  internalPreviewObservations,
} = require("../automation/queryCandidatePlannerInternalPreviewController");

router.use(protect);

function observeQueryCandidatePlannerForInternalPreview(observation, context) {
  defaultObservationLogger(observation, context);
  recordQueryCandidatePlannerInternalPreviewObservation(observation);
  // void recordQueryCandidatePlannerRealShadowObservation(observation, context);
  void recordQueryCandidatePlannerRealShadowObservation(
    observation,
    context,
  ).then((result) => {
    console.log("[real-shadow-evidence]", {
      kind: "EXECUTION",
      stored: result?.stored === true,
      reason: String(result?.reason || "UNKNOWN"),
    });
  });
}

const getAnalysisCandidatesShadowObserved =
  createQueryCandidatePlannerInternalAllowlistCanaryBoundary({
    handler: automationController.getAnalysisCandidates,
    onObservation: observeQueryCandidatePlannerForInternalPreview,
    shadowRunner: runQueryCandidatePlannerApiShadowWithRealEvidenceCapture,
  });
const downloadGeneratedFileCacheRetained =
  createQueryCandidatePlannerDownloadRetentionBoundary({
    handler: automationController.downloadGeneratedFile,
    action: "GENERATED_DOWNLOAD",
    onObservation: (observation, context) => {
      // void recordQueryCandidatePlannerRealShadowLifecycleObservation(
      //   observation,
      //   context,
      // );
      void recordQueryCandidatePlannerRealShadowLifecycleObservation(
        observation,
        context,
      ).then((result) => {
        console.log("[real-shadow-evidence]", {
          kind: "LIFECYCLE",
          stored: result?.stored === true,
          reason: String(result?.reason || "UNKNOWN"),
        });
      });
    },
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

router.get("/internal/query-candidate-shadow-preview", internalPreviewPage);
router.get(
  "/internal/query-candidate-shadow-preview/status",
  requireQueryCandidatePlannerInternalPreviewAccess,
  internalPreviewStatus,
);
router.get(
  "/internal/query-candidate-shadow-preview/observations",
  requireQueryCandidatePlannerInternalPreviewAccess,
  internalPreviewObservations,
);

module.exports = router;
