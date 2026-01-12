const express = require("express");
const router = express.Router();
const passport = require("passport");
const authController = require("../controllers/authController");
const { protect } = require("../middleware/authMiddleware");

// 이메일 관련 라우트
router.post("/signup", authController.signup);
router.get("/verify-email/:token", authController.verifyEmail);
router.post("/login", authController.login);
router.post("/forgot-password", authController.forgotPassword);
router.patch("/reset-password/:token", authController.resetPassword);

// 회원 탈퇴 (로그인 필요)
router.post("/withdraw", protect, authController.withdraw);

// 1. 사용자가 '구글 로그인' 버튼을 눌렀을 때 호출될 경로
router.get(
  "/google",
  passport.authenticate("google", { scope: ["profile", "email"] })
);

// 2. 구글 로그인 성공 후, 구글이 리디렉션할 경로 (callback)
router.get("/google/callback", (req, res, next) => {
  passport.authenticate("google", { session: false }, (err, user, info) => {
    const frontendURL = process.env.FRONTEND_URL;

    if (err) return next(err);

    // ✅ 재가입 차단 UX
    if (!user && info?.code === "REJOIN_BLOCKED") {
      const until = info?.until ? encodeURIComponent(info.until) : "";
      return res.redirect(
        `${frontendURL}/?authError=REJOIN_BLOCKED&until=${until}`
      );
    }

    if (!user) {
      return res.redirect(`${frontendURL}/?authError=GOOGLE_LOGIN_FAILED`);
    }

    req.user = user;
    return authController.googleCallback(req, res);
  })(req, res, next);
});

router.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

module.exports = router;
