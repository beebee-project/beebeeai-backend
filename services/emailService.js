const nodemailer = require("nodemailer");

const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587, // 465 쓰면 secure: true
  secure: false,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

const sendVerificationEmail = async (to, token) => {
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

  <p style="font-size: 14px; margin-top: 40px; color: #777; text-align:center;">
    인증 버튼이 작동하지 않으면 고객센터(help@beebeeai.kr)로 문의해 주세요.
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

module.exports = { sendVerificationEmail, sendPasswordResetEmail };
