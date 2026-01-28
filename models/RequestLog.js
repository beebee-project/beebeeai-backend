const mongoose = require("mongoose");

const RequestLogSchema = new mongoose.Schema(
  {
    traceId: { type: String, index: true },
    userId: { type: mongoose.Schema.Types.ObjectId, ref: "User", index: true },

    // 어떤 기능 요청이었는지
    route: { type: String }, // e.g. "/convert", "/macro/generate"
    engine: { type: String, index: true }, // "formula" | "officescripts" | "appscript" | "sql" ...

    // 결과
    status: { type: String, enum: ["success", "fail"], index: true },
    reason: { type: String, index: true },
    isFallback: { type: Boolean, default: false, index: true },

    // 내용(필요 최소만)
    prompt: { type: String },

    // 성능/분석
    latencyMs: { type: Number },
    debugMeta: { type: Object }, // 6-2에서 표준화 예정
  },
  { timestamps: true }
);

module.exports =
  mongoose.models.RequestLog || mongoose.model("RequestLog", RequestLogSchema);
