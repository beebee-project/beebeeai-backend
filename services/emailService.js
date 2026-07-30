const nodemailer = require("nodemailer");

const EMAIL_ENABLED = Boolean(process.env.EMAIL_USER && process.env.EMAIL_PASS);
const INQUIRY_RECIPIENT = process.env.INQUIRY_RECIPIENT || "hello@beebeeai.kr";

const transporter = EMAIL_ENABLED
  ? nodemailer.createTransport({
      host: "smtp.gmail.com",
      port: 587,
      secure: false,
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    })
  : null;

const sendVerificationEmail = async (to, token) => {
  if (!EMAIL_ENABLED || !transporter) {
    console.log("[email disabled] sendVerificationEmail", { to, token });
    return { skipped: true };
  }
  const verificationLink = `${process.env.FRONTEND_URL}/verify.html?token=${token}`;

  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: to,
    subject: "BeeBee AI 회원가입 이메일 인증",
    html: `
      <div style="font-family: Pretendard, sans-serif; max-width: 600px; margin: auto;">
  <h2 style="font-weight: 700;">BeeBee AI에 가입해 주셔서 감사합니다!</h2>
  <p style="font-size: 16px; margin-top: 12px;">
      아래 버튼을 클릭하여 이메일 인증을 완료해 주세요.
  </p>

  <div style="margin-top: 32px; display:flex; justify-content:center;">
    <a href="${verificationLink}"
       style="
         background-color: #FFC800;
         padding: 14px 26px;
         border-radius: 8px;
         color: black;
         text-decoration: none;
         font-size: 18px;
         font-weight: 600;
       ">
      이메일 인증하기
    </a>
  </div>

  <p style="font-size: 14px; margin-top: 40px; color: #777;">
    인증 버튼이 작동하지 않으면 고객센터(hello@beebeeai.kr)로 문의해 주세요.
  </p>
</div>
    `,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Verification email sent to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
    throw new Error("이메일 발송에 실패했습니다.");
  }
};

const sendPasswordResetEmail = async (to, token) => {
  if (!EMAIL_ENABLED || !transporter) {
    console.log("[email disabled] sendPasswordResetEmail", { to, token });
    return { skipped: true };
  }
  const resetLink = `${process.env.FRONTEND_URL}/reset-password.html?token=${token}`;

  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: to,
    subject: "BeeBee AI 비밀번호 재설정",
    html: `
      <h2>BeeBee AI 비밀번호 재설정 요청</h2>
      <p>비밀번호를 재설정하려면 아래 버튼을 클릭하세요. 이 링크는 10분간 유효합니다.</p>
      <a href="${resetLink}" style="background-color: #ffc800; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">비밀번호 재설정하기</a>
      <p>만약 위 버튼이 동작하지 않으면, 아래 링크를 브라우저에 복사하여 붙여넣어 주세요:</p>
      <p>${resetLink}</p>
    `,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Password reset email sent to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
    throw new Error("이메일 발송에 실패했습니다.");
  }
};

function escapeHtml(value = "") {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

const INQUIRY_TYPE_LABELS = Object.freeze({
  service: "서비스 이용 문의",
  "output-improvement": "생성 결과 개선 요청",
  account: "계정 및 로그인 문의",
  payment: "결제 및 구독 문의",
  feature: "기능 제안",
  other: "기타 문의",
});

const sendInquiryEmail = async ({
  type,
  email,
  subject,
  message,
  source = "web",
  pageUrl = "",
} = {}) => {
  if (!EMAIL_ENABLED || !transporter) {
    const error = new Error("문의 메일 발송 설정이 완료되지 않았습니다.");
    error.code = "EMAIL_DISABLED";
    error.status = 503;
    throw error;
  }

  const typeLabel =
    INQUIRY_TYPE_LABELS[String(type || "").trim()] || "기타 문의";
  const safeEmail = String(email || "").trim();
  const safeSubject = String(subject || "").trim();
  const safeMessage = String(message || "").trim();
  const safeSource = String(source || "web").trim();
  const safePageUrl = String(pageUrl || "").trim();
  const receivedAt = new Date().toISOString();

  const mailOptions = {
    from: `"BeeBee AI 문의" <${process.env.EMAIL_USER}>`,
    to: INQUIRY_RECIPIENT,
    replyTo: safeEmail,
    subject: `[BeeBee AI 문의][${typeLabel}] ${safeSubject}`,
    text: [
      `문의 유형: ${typeLabel}`,
      `답변 이메일: ${safeEmail}`,
      `접수 시각: ${receivedAt}`,
      `접수 경로: ${safeSource}`,
      safePageUrl ? `페이지: ${safePageUrl}` : "",
      "",
      safeMessage,
    ]
      .filter(Boolean)
      .join("\n"),
    html: `
      <div style="font-family: Arial, Pretendard, sans-serif; max-width: 680px; margin: 0 auto; color: #272522;">
        <h2 style="margin-bottom: 24px;">BeeBee AI 문의 접수</h2>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 24px;">
          <tbody>
            <tr>
              <th style="width: 130px; padding: 10px; border: 1px solid #dddddd; text-align: left; background: #f7f7f7;">문의 유형</th>
              <td style="padding: 10px; border: 1px solid #dddddd;">${escapeHtml(typeLabel)}</td>
            </tr>
            <tr>
              <th style="padding: 10px; border: 1px solid #dddddd; text-align: left; background: #f7f7f7;">답변 이메일</th>
              <td style="padding: 10px; border: 1px solid #dddddd;">${escapeHtml(safeEmail)}</td>
            </tr>
            <tr>
              <th style="padding: 10px; border: 1px solid #dddddd; text-align: left; background: #f7f7f7;">접수 시각</th>
              <td style="padding: 10px; border: 1px solid #dddddd;">${escapeHtml(receivedAt)}</td>
            </tr>
            <tr>
              <th style="padding: 10px; border: 1px solid #dddddd; text-align: left; background: #f7f7f7;">접수 경로</th>
              <td style="padding: 10px; border: 1px solid #dddddd;">${escapeHtml(safeSource)}</td>
            </tr>
            ${
              safePageUrl
                ? `
            <tr>
              <th style="padding: 10px; border: 1px solid #dddddd; text-align: left; background: #f7f7f7;">페이지</th>
              <td style="padding: 10px; border: 1px solid #dddddd; word-break: break-all;">${escapeHtml(safePageUrl)}</td>
            </tr>`
                : ""
            }
          </tbody>
        </table>

        <h3 style="margin: 0 0 10px;">${escapeHtml(safeSubject)}</h3>
        <div style="padding: 16px; border-radius: 10px; background: #f7f7f7; line-height: 1.7; white-space: pre-wrap;">${escapeHtml(safeMessage)}</div>

        <p style="margin-top: 24px; color: #777777; font-size: 13px;">
          이 메일에 답장하면 사용자가 입력한 이메일로 전송됩니다.
        </p>
      </div>
    `,
  };

  try {
    const info = await transporter.sendMail(mailOptions);

    console.log("[inquiry email sent]", {
      recipient: INQUIRY_RECIPIENT,
      messageId: info.messageId || "",
      acceptedCount: Array.isArray(info.accepted) ? info.accepted.length : 0,
      rejectedCount: Array.isArray(info.rejected) ? info.rejected.length : 0,
    });

    return {
      messageId: info.messageId || "",
      accepted: Array.isArray(info.accepted) ? info.accepted : [],
      rejected: Array.isArray(info.rejected) ? info.rejected : [],
      recipient: INQUIRY_RECIPIENT,
    };
  } catch (error) {
    console.error("[inquiry email failed]", {
      code: error?.code || "",
      responseCode: error?.responseCode || 0,
      message: error?.message || String(error),
    });

    const wrapped = new Error("문의 메일 발송에 실패했습니다.");
    wrapped.code = "INQUIRY_EMAIL_SEND_FAILED";
    wrapped.status = 502;
    wrapped.cause = error;
    throw wrapped;
  }
};

module.exports = {
  sendVerificationEmail,
  sendPasswordResetEmail,
  sendInquiryEmail,
};
