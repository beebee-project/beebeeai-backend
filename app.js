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
  "https://beebeeai-frontend-production.up.railway.app",
]);

const corsMiddleware = cors({
  origin: (origin, cb) => {
    // Postman/서버-서버(Origin 없음) 허용
    if (!origin) return cb(null, true);

    // 정확 매칭
    if (ALLOWED_ORIGINS.has(origin)) return cb(null, true);

    // ✅ 운영 편의: https://*.beebeeai.kr 허용
    try {
      const { protocol, hostname } = new URL(origin);
      if (protocol === "https:" && hostname.endsWith(".beebeeai.kr")) {
        return cb(null, true);
      }
    } catch (_) {}

    // ❗ 절대 Error 던지지 말기 (브라우저에서만 CORS 차단)
    console.warn("[CORS BLOCKED]", origin);
    return cb(null, false);
  },
  methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization", "x-cron-secret"],
  credentials: true,
  optionsSuccessStatus: 204,
});
app.use(corsMiddleware);
// ✅ 모든 OPTIONS 요청은 cors가 204로 응답하도록 명시
app.options("*", corsMiddleware);
// ✅ 웹 서버에서는 내부 cron을 기본 OFF
if (process.env.RUN_INTERNAL_CRON === "1") {
  startDailySummaryCron();
}

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

// 에러 핸들러
app.use(errorHandler);

// 서버 시작
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 BeeBeeAI API running on port ${PORT}`);
});
