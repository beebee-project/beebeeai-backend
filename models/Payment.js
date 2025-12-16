const mongoose = require("mongoose");

const paymentSchema = new mongoose.Schema(
  {
    provider: { type: String, default: "toss", index: true },
    userId: { type: mongoose.Schema.Types.ObjectId, ref: "User", index: true },

    orderId: { type: String, required: true, index: true, unique: true },
    paymentKey: { type: String, index: true, sparse: true },

    amount: { type: Number, required: true },
    currency: { type: String, default: "KRW" },

    orderName: { type: String },
    status: {
      type: String,
      enum: ["READY", "DONE", "FAILED", "CANCELED"],
      default: "READY",
      index: true,
    },

    raw: { type: Object }, // toss 원본 응답 저장(필요할 때만)
    approvedAt: { type: Date },
  },
  { timestamps: true }
);

module.exports = mongoose.model("Payment", paymentSchema);
