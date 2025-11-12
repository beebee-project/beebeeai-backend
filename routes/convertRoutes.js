const express = require("express");
const router = express.Router();
const convertController = require("../controllers/convertController");
const { protect } = require("../middleware/authMiddleware");

router.post("/", protect, convertController.handleConversion);

router.post("/feedback", convertController.handleFeedback);

module.exports = router;
