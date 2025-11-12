require("dotenv").config();
const express = require("express");
const cors = require("cors");
const passport = require("passport");
const connectDB = require("./config/db");
const errorHandler = require("./middleware/errorHandler");

// 라우터
const authRoutes = require("./routes/authRoutes");
const fileRoutes = require("./routes/fileRoutes");
const convertRoutes = require("./routes/convertRoutes");
const paymentRoutes = require("./routes/paymentRoutes");

const app = express();

// 프록시 신뢰 (Cloudflare/Railway 뒤에 있을 때 HTTPS 스킴 등 믿도록)
app.set("trust proxy", 1);

// ==== CORS (운영 도메인만 허용) ====
const ALLOWED_ORIGINS = new Set([
  "https://beebeeai.kr",
  "https://www.beebeeai.kr",
  "https://api.beebeeai.kr",
  "http://localhost:3000",
]);

app.use(
  cors({
    origin: (origin, cb) => {
      // origin이 없을 수도 있음(서버-서버 호출/헬스체크 등) → 허용
      if (!origin || ALLOWED_ORIGINS.has(origin)) return cb(null, true);
      return cb(new Error("Not allowed by CORS"));
    },
    methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: true,
  })
);

// JSON/폼 파서
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: false, limit: "50mb" }));

// Passport
app.use(passport.initialize());
try {
  require("./config/passport")(passport);
} catch (_) {
  // 선택 모듈: 없으면 무시
}

// ==== MongoDB 연결 ====
connectDB();

// ==== Health ====
app.get("/api/health", (req, res) => {
  res.json({ ok: true, t: Date.now() });
});

// ==== 라우터 ====
app.use("/api/auth", authRoutes);
app.use("/api/files", fileRoutes);
app.use("/api/convert", convertRoutes);
app.use("/api/payments", paymentRoutes);

// ==== 에러 핸들러 ====
app.use(errorHandler);

// ==== 서버 리슨 (Railway는 PORT를 환경변수로 제공) ====
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`🚀 BeeBeeAI API is running on port ${PORT}`);
});
