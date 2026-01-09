const GoogleStrategy = require("passport-google-oauth20").Strategy;
const crypto = require("crypto");
const User = require("../models/User");

function sha256(input) {
  return crypto.createHash("sha256").update(input).digest("hex");
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
          const email = String(profile?.emails?.[0]?.value || "")
            .toLowerCase()
            .trim();
          const googleId = String(profile?.id || "");
          const now = new Date();

          const emailHash = email ? sha256(email) : null;

          // 1) (최우선) 30일 이내 "삭제된 계정" 복구: googleId 또는 emailHash로 탐색
          let user = await User.findOne({
            isDeleted: true,
            purgeAt: { $ne: null, $gt: now },
            $or: [
              googleId ? { "authIdentity.googleId": googleId } : null,
              emailHash ? { "authIdentity.emailHash": emailHash } : null,
            ].filter(Boolean),
          });

          if (user) {
            // 복구
            user.isDeleted = false;
            user.deletedAt = null;
            user.purgeAt = null;

            // email/구글ID 최신값 반영
            if (email) user.email = email;
            user.googleId = googleId;
            user.authIdentity = user.authIdentity || {};
            user.authIdentity.googleId = googleId;
            if (emailHash) user.authIdentity.emailHash = emailHash;

            user.isVerified = true;
            await user.save();

            return done(null, user);
          }

          // 2) (일반) 활성 계정 찾기: googleId 우선, 없으면 email
          user = await User.findOne({
            isDeleted: false,
            $or: [
              googleId ? { googleId } : null,
              email ? { email } : null,
            ].filter(Boolean),
          });

          if (user) {
            // googleId 누락 케이스 보정
            if (!user.googleId && googleId) {
              user.googleId = googleId;
              user.authIdentity = user.authIdentity || {};
              user.authIdentity.googleId = googleId;
              if (emailHash) user.authIdentity.emailHash = emailHash;
              await user.save();
            }
            return done(null, user);
          }

          // 3) 신규 생성
          const newUser = {
            googleId,
            name: profile.displayName,
            email,
            isVerified: true,
            authIdentity: {
              googleId,
              emailHash,
            },
          };

          user = await User.create(newUser);
          return done(null, user);
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
