const express = require("express");
const cors = require("cors");
const passport = require("passport");

const connectDB = require("./config/db");
const errorHandler = require("./middleware/errorHandler");

// 라우터 모듈
const authRoutes = require("./routes/authRoutes");
const fileRoutes = require("./routes/fileRoutes");
const convertRoutes = require("./routes/convertRoutes");
const paymentRoutes = require("./routes/paymentRoutes");
const macroRoutes = require("./routes/macroRoutes");
const adminRoutes = require("./routes/adminRoutes");
const cronRoutes = require("./routes/cronRoutes");
const { startDailySummaryCron } = require("./cron/dailySummaryCron");

// 앱 초기화
const app = express();
app.set("trust proxy", 1);

// CORS (프론트/백엔드 도메인 허용)
const ALLOWED_ORIGINS = new Set([
  "https://beebeeai.kr",
  "https://www.beebeeai.kr",
  "http://localhost:3000",
]);

// ✅ preflight(OPTIONS)는 무조건 빠르게 204로 종료 (502/timeout 방지)
app.use((req, res, next) => {
  const origin = req.headers.origin;
  if (origin && ALLOWED_ORIGINS.has(origin)) {
    res.header("Access-Control-Allow-Origin", origin);
    res.header("Vary", "Origin");
  }
  res.header("Access-Control-Allow-Credentials", "true");
  res.header(
    "Access-Control-Allow-Methods",
    "GET,POST,PUT,PATCH,DELETE,OPTIONS",
  );
  res.header(
    "Access-Control-Allow-Headers",
    "Content-Type, Authorization, x-cron-secret",
  );
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

// cors 패키지는 “헤더 셋”은 위에서 처리하므로 간단히만
app.use(
  cors({
    origin: (origin, cb) => {
      if (!origin || ALLOWED_ORIGINS.has(origin)) return cb(null, true);
      return cb(null, false);
    },
    credentials: true,
  }),
);

// 바디 파서
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: false, limit: "50mb" }));

// Passport
app.use(passport.initialize());
try {
  require("./config/passport")(passport);
} catch (_) {}

// DB 연결
connectDB();

// 헬스 체크
app.get("/api/health", (req, res) => {
  res.json({ ok: true, t: Date.now() });
});

// 라우트
app.use("/api/auth", authRoutes);
app.use("/api/files", fileRoutes);
app.use("/api/convert", convertRoutes);
app.use("/api/payments", paymentRoutes);
app.use("/api/macro", macroRoutes);
app.use("/admin", adminRoutes);
app.use("/cron", cronRoutes);

startDailySummaryCron();

// 에러 핸들러
app.use(errorHandler);

// 서버 시작
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 BeeBeeAI API running on port ${PORT}`);
});
