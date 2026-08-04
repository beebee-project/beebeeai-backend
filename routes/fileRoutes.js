const express = require("express");
const multer = require("multer");
const fileController = require("../controllers/fileController");
const { protect } = require("../middleware/authMiddleware");
const {
  createQueryCandidatePlannerMutationBoundary,
  createQueryCandidatePlannerDownloadRetentionBoundary,
} = require("../automation/queryCandidatePlannerFileLifecycleBoundary");
const router = express.Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
});

router.use(protect);

const uploadFileCacheObserved =
  createQueryCandidatePlannerMutationBoundary({
    handler: fileController.uploadFile,
    action: "UPLOAD_REPLACEMENT",
  });
const downloadFileCacheRetained =
  createQueryCandidatePlannerDownloadRetentionBoundary({
    handler: fileController.downloadFile,
    action: "SOURCE_DOWNLOAD",
  });
const deleteFileCacheObserved =
  createQueryCandidatePlannerMutationBoundary({
    handler: fileController.deleteFile,
    action: "DELETE",
  });

router.route("/").get(fileController.getFiles);
router.route("/upload").post(upload.single("file"), uploadFileCacheObserved);
router.route("/download/:originalName").get(downloadFileCacheRetained);
router.route("/:originalName").delete(deleteFileCacheObserved);

module.exports = router;
