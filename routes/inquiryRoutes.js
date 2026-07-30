const express = require("express");
const { sendInquiryEmail } = require("../services/emailService");

const router = express.Router();

const ALLOWED_INQUIRY_TYPES = new Set([
  "service",
  "output-improvement",
  "account",
  "payment",
  "feature",
  "other",
]);

function normalizeText(value = "", maxLength = 1000) {
  return String(value ?? "")
    .trim()
    .slice(0, maxLength);
}

function isValidEmail(value = "") {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ""));
}

router.post("/", async (req, res) => {
  const type = normalizeText(req.body?.type, 40);
  const email = normalizeText(req.body?.email, 120).toLowerCase();
  const subject = normalizeText(req.body?.subject, 100);
  const message = normalizeText(req.body?.message, 1000);
  const source = normalizeText(req.body?.source || "web", 30);
  const pageUrl = normalizeText(req.body?.pageUrl, 500);
  const consent = req.body?.consent === true;

  if (!ALLOWED_INQUIRY_TYPES.has(type)) {
    return res.status(400).json({
      ok: false,
      code: "INQUIRY_TYPE_INVALID",
      message: "문의 유형을 선택해주세요.",
    });
  }

  if (!isValidEmail(email)) {
    return res.status(400).json({
      ok: false,
      code: "INQUIRY_EMAIL_INVALID",
      message: "답변받을 이메일을 정확히 입력해주세요.",
    });
  }

  if (!subject) {
    return res.status(400).json({
      ok: false,
      code: "INQUIRY_SUBJECT_REQUIRED",
      message: "문의 제목을 입력해주세요.",
    });
  }

  if (!message) {
    return res.status(400).json({
      ok: false,
      code: "INQUIRY_MESSAGE_REQUIRED",
      message: "문의 내용을 입력해주세요.",
    });
  }

  if (!consent) {
    return res.status(400).json({
      ok: false,
      code: "INQUIRY_CONSENT_REQUIRED",
      message: "문의 처리에 필요한 정보 수집에 동의해주세요.",
    });
  }

  try {
    const delivery = await sendInquiryEmail({
      type,
      email,
      subject,
      message,
      source,
      pageUrl,
    });

    return res.status(201).json({
      ok: true,
      message:
        "문의가 정상적으로 접수되었습니다. 확인 후 이메일로 답변드리겠습니다.",
      delivery: {
        messageId: delivery.messageId,
        recipient: delivery.recipient,
      },
    });
  } catch (error) {
    console.error("[inquiry route failed]", {
      code: error?.code || "",
      status: error?.status || 500,
      message: error?.message || String(error),
    });

    return res.status(error?.status || 500).json({
      ok: false,
      code: error?.code || "INQUIRY_SUBMIT_FAILED",
      message:
        error?.message ||
        "문의 전송에 실패했습니다. 잠시 후 다시 시도해주세요.",
    });
  }
});

module.exports = router;
