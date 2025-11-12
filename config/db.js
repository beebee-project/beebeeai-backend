const mongoose = require("mongoose");

let cached = global.__mongooseConn;

/**
 * 최신 Mongoose(v8) 권장 방식:
 * - useNewUrlParser / useUnifiedTopology 옵션 불필요
 * - serverSelectionTimeoutMS 로 초기 연결 타임아웃 제어
 * - Lambda/Hot-reload 대비: 연결 캐시(cached) 사용
 */
async function connectDB() {
  if (cached && mongoose.connection.readyState === 1) {
    console.log("⚡ Mongo already connected");
    return mongoose.connection;
  }

  const uri = process.env.MONGO_URI;
  if (!uri) {
    console.error("❌ MONGO_URI is not set");
    throw new Error("MONGO_URI is not set");
  }

  try {
    console.log(
      "🔗 Connecting MongoDB:",
      uri.replace(/\/\/([^:]+):([^@]+)@/, "//<user>:<pass>@").slice(0, 80) +
        "..."
    );

    const conn = await mongoose.connect(uri, {
      serverSelectionTimeoutMS: 10000, // 10s
      // tls: true, // 필요 시 강제 TLS
      // dbName: "beebee", // URI에 /beebee가 없다면 여기서 지정 가능
    });

    global.__mongooseConn = conn;
    console.log("✅ MongoDB connected:", conn.connection.host);
    return conn.connection;
  } catch (err) {
    console.error("❌ MongoDB connection error:", err?.message || err);
    throw err;
  }
}

module.exports = connectDB;
