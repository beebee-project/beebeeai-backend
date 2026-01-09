const GoogleStrategy = require("passport-google-oauth20").Strategy;
const User = require("../models/User");

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
          const email = String(rawEmail).trim().toLowerCase();
          const googleId = String(profile?.id || "");
          const name = profile?.displayName || "사용자";

          if (!email) {
            // 이메일이 없으면 계정 연동/생성이 불가(정책적으로 실패 처리 추천)
            return done(new Error("Google account has no email"), null);
          }

          // 1) 이메일로 기존 계정 찾기
          let user = await User.findOne({ email });

          if (user) {
            // 2) googleId가 비어있거나 다르면 업데이트(연동)
            if (!user.googleId || user.googleId !== googleId) {
              user.googleId = googleId;
              if (!user.name) user.name = name;
              // 구글 로그인 성공 시 검증 처리 정책(원하면 true)
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
