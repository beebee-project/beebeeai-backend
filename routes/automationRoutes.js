const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const automationController = require("../controllers/automationController");

router.use(protect);

router.post("/query-preview", automationController.previewQueryTables);

module.exports = router;
