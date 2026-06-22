const express = require("express");
const router = express.Router();

function directConversionRemoved(req, res) {
  return res.status(410).json({
    ok: false,
    code: "DIRECT_CONVERSION_REMOVED",
    message:
      "직접 수식 생성 API는 제거되었습니다. 템플릿 생성 기능을 이용해 주세요.",
  });
}

router.all("/", directConversionRemoved);
router.all("/feedback", directConversionRemoved);

module.exports = router;
