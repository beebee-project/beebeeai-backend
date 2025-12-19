const User = require("../models/User");
const jwt = require("jsonwebtoken");
const crypto = require("crypto");
const bcrypt = require("bcryptjs");
const {
  sendVerificationEmail,
  sendPasswordResetEmail,
} = require("../services/emailService");

// JWT 생성 함수
const signToken = (id) => {
  return jwt.sign({ id }, process.env.JWT_SECRET, {
    expiresIn: process.env.JWT_EXPIRES_IN,
  });
};

// 회원가입 로직
exports.signup = async (req, res, next) => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      return res
        .status(400)
        .json({ message: "이메일과 비밀번호를 모두 입력해주세요." });
    }

    const existingUser = await User.findOne({ email });
    if (existingUser && existingUser.isVerified) {
      return res.status(400).json({ message: "이미 사용 중인 이메일입니다." });
    }
    if (existingUser && !existingUser.isVerified) {
      await User.deleteOne({ email });
    }

    const name = req.body.name || email.split("@")[0];

    const newUser = new User({ name, email, password });
    const verificationToken = newUser.createEmailVerificationToken();
    await newUser.save();

    await sendVerificationEmail(newUser.email, verificationToken);

    res.status(201).json({
      message:
        "회원가입 신청이 완료되었습니다. 이메일을 확인하여 인증을 완료해주세요.",
    });
  } catch (error) {
    next(error);
  }
};

// 이메일 인증 처리 로직
exports.verifyEmail = async (req, res, next) => {
  try {
    const hashedToken = crypto
      .createHash("sha256")
      .update(req.params.token)
      .digest("hex");

    const user = await User.findOne({
      emailVerificationToken: hashedToken,
      emailVerificationExpires: { $gt: Date.now() },
    });

    if (!user) {
      return res
        .status(400)
        .json({ message: "인증 토큰이 유효하지 않거나 만료되었습니다." });
    }

    user.isVerified = true;
    user.emailVerificationToken = undefined;
    user.emailVerificationExpires = undefined;
    await user.save();

    res.status(200).json({
      message:
        "이메일 인증이 성공적으로 완료되었습니다. 이제 로그인할 수 있습니다.",
    });
  } catch (error) {
    next(error);
  }
};

// 로그인 로직
exports.login = async (req, res, next) => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      return res
        .status(400)
        .json({ message: "이메일과 비밀번호를 입력해주세요." });
    }

    const user = await User.findOne({ email }).select("+password");

    if (!user || !(await user.comparePassword(password))) {
      return res
        .status(401)
        .json({ message: "이메일 또는 비밀번호가 올바르지 않습니다." });
    }

    if (!user.isVerified) {
      return res.status(401).json({
        message: "이메일 인증이 완료되지 않았습니다. 메일함을 확인해주세요.",
      });
    }

    const token = signToken(user._id);
    res.status(200).json({ token, message: "로그인 성공" });
  } catch (error) {
    next(error);
  }
};

// 비밀번호 재설정 요청 로직
exports.forgotPassword = async (req, res, next) => {
  try {
    const user = await User.findOne({ email: req.body.email });
    if (!user) {
      return res.status(200).json({
        message: "해당 이메일로 비밀번호 재설정 링크를 전송했습니다.",
      });
    }

    const resetToken = user.createPasswordResetToken();
    await user.save({ validateBeforeSave: false });

    await sendPasswordResetEmail(user.email, resetToken);

    res
      .status(200)
      .json({ message: "해당 이메일로 비밀번호 재설정 링크를 전송했습니다." });
  } catch (error) {
    next(error);
  }
};

// 비밀번호 재설정 실행 로직
exports.resetPassword = async (req, res, next) => {
  try {
    const hashedToken = crypto
      .createHash("sha256")
      .update(req.params.token)
      .digest("hex");

    const user = await User.findOne({
      passwordResetToken: hashedToken,
      passwordResetExpires: { $gt: Date.now() },
    });

    if (!user) {
      return res
        .status(400)
        .json({ message: "토큰이 유효하지 않거나 만료되었습니다." });
    }

    user.password = req.body.password;
    user.passwordResetToken = undefined;
    user.passwordResetExpires = undefined;
    await user.save();

    const token = signToken(user._id);
    res
      .status(200)
      .json({ token, message: "비밀번호가 성공적으로 재설정되었습니다." });
  } catch (error) {
    next(error);
  }
};

// 구글 로그인 성공 후 실행될 콜백 로직
exports.googleCallback = (req, res) => {
  const token = signToken(req.user._id);
  const frontendURL = process.env.FRONTEND_URL;
  res.redirect(`${frontendURL}/?token=${token}`);
};

exports.withdraw = async (req, res, next) => {
  try {
    const { password } = req.body || {};
    if (!password) {
      return res.status(400).json({ message: "비밀번호를 입력해주세요." });
    }

    // password가 select:false일 수 있으니 +password로 가져오기
    const user = await User.findById(req.user.id).select(
      "+password email name plan subscription"
    );
    if (!user) return res.status(404).json({ message: "사용자 없음" });

    // 소셜 로그인 등 password 없는 계정 처리
    if (!user.password) {
      return res.status(400).json({
        message:
          "비밀번호가 설정되지 않은 계정입니다. 비밀번호 설정 후 탈퇴할 수 있습니다.",
      });
    }

    const ok = await bcrypt.compare(password, user.password);
    if (!ok) {
      return res.status(401).json({ message: "비밀번호가 올바르지 않습니다." });
    }

    // ✅ 구독 중이면 탈퇴 불가(먼저 해지 유도)
    const sub = user.subscription || {};
    const status = String(sub.status || "").toUpperCase();

    const isSubscribed =
      status === "TRIAL" ||
      status === "ACTIVE" ||
      status === "PAST_DUE" ||
      status === "CANCELED_PENDING"; // 유료 해지 예약 중(만료일까지 사용)

    if (isSubscribed) {
      return res.status(409).json({
        message:
          "구독 중인 계정은 탈퇴할 수 없습니다. 먼저 구독 해지를 진행해주세요.",
        code: "SUBSCRIPTION_ACTIVE",
        status,
      });
    }

    const now = new Date();

    // ✅ 탈퇴 = 즉시 이용 종료 + 즉시 청구 중단
    user.plan = "FREE";
    user.subscription = {
      ...(user.subscription || {}),
      status: "CANCELED",
      canceledAt: now,
      endedAt: now,
      nextChargeAt: null, // ✅ cron 청구 트리거 제거
      cancelAtPeriodEnd: false,
    };

    // ✅ soft delete + 익명화(권장)
    user.isDeleted = true;
    user.deletedAt = now;

    // email unique 충돌 방지
    if (user.email) user.email = `deleted_${user._id}@deleted.local`;
    if (user.name) user.name = "deleted";

    // password 제거(선택) - 이후 로그인 불가
    user.password = undefined;

    await user.save();

    // JWT는 서버가 강제 폐기하기 어려움(무상태) -> 프론트에서 토큰 삭제/로그아웃 처리 권장
    return res.json({ ok: true, message: "회원 탈퇴가 완료되었습니다." });
  } catch (error) {
    next(error);
  }
};
