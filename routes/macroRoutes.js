const router = require("express").Router();
const { protect } = require("../middleware/authMiddleware");
const macroController = require("../controllers/macroController");

router.post("/generate", protect, macroController.generateMacro);

module.exports = router;
