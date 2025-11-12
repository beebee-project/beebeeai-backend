const express = require("express");
const multer = require("multer");
const fileController = require("../controllers/fileController");
const { protect } = require("../middleware/authMiddleware");
const router = express.Router();

// Multer 설정: 파일을 메모리에 저장
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
});

// 모든 라우트에 protect 미들웨어를 적용하여 로그인된 사용자만 접근 가능하도록 함
router.use(protect);

router.route("/").get(fileController.getFiles);

router.route("/upload").post(upload.single("file"), fileController.uploadFile); // 'file'은 FormData의 키 이름

router.route("/download/:originalName").get(fileController.downloadFile);

router.route("/:originalName").delete(fileController.deleteFile);

module.exports = router;
