const GoogleStrategy = require("passport-google-oauth20").Strategy;
const crypto = require("crypto");
const User = require("../models/User");

function normalizeEmail(raw) {
  return String(raw || "")
    .trim()
    .toLowerCase();
}

// 역추적 어려운 HMAC 추천 (REJOIN_PEPPER 없으면 JWT_SECRET fallback)
function emailHash(email) {
  const key = process.env.REJOIN_PEPPER || process.env.JWT_SECRET || "fallback";
  return crypto.createHmac("sha256", key).update(email).digest("hex");
}

module.exports = function (passport) {
  passport.use(
    new GoogleStrategy(
      {
        clientID: process.env.GOOGLE_CLIENT_ID,
        clientSecret: process.env.GOOGLE_CLIENT_SECRET,
        callbackURL: process.env.GOOGLE_REDIRECT_URI,
      },
      async (accessToken, refreshToken, profile, done) => {
        try {
          const rawEmail = profile?.emails?.[0]?.value || "";
          const email = normalizeEmail(rawEmail);
          const googleId = String(profile?.id || "");
          const name = profile?.displayName || "사용자";

          if (!email) {
            // 이메일이 없으면 계정 연동/생성이 불가
            return done(new Error("Google account has no email"), null);
          }

          // ✅ [추가] 30일 재가입 금지(탈퇴 후 purgeAt 이전)
          const now = new Date();
          const blocked = await User.findOne({
            isDeleted: true,
            purgeAt: { $ne: null, $gt: now },
            "authIdentity.emailHash": emailHash(email),
          }).select("purgeAt");

          if (blocked) {
            // passport에서는 에러를 던지기보다 "인증 실패"로 처리하는 게 안전
            // (라우터에서 failureRedirect로 보내짐)
            return done(null, false, {
              code: "REJOIN_BLOCKED",
              until: blocked.purgeAt
                ? new Date(blocked.purgeAt).toISOString()
                : null,
            });
          }

          // 1) 이메일로 기존 계정 찾기
          let user = await User.findOne({ email });

          if (user) {
            // 2) googleId가 비어있거나 다르면 업데이트(연동)
            if (!user.googleId || user.googleId !== googleId) {
              user.googleId = googleId;
              if (!user.name) user.name = name;
              user.isVerified = true;
              await user.save();
            }
            return done(null, user);
          }

          // 3) 신규 생성
          try {
            user = await User.create({
              googleId,
              name,
              email,
              isVerified: true,
            });
            return done(null, user);
          } catch (e) {
            // 레이스 컨디션: 동시에 생성된 경우
            if (e?.code === 11000) {
              const existing = await User.findOne({ email });
              if (existing) return done(null, existing);
            }
            throw e;
          }
        } catch (err) {
          console.error(err);
          return done(err, null);
        }
      }
    )
  );

  passport.serializeUser((user, done) => done(null, user.id));
  passport.deserializeUser((id, done) => {
    User.findById(id, (err, user) => done(err, user));
  });
};
