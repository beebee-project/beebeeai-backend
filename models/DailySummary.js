const mongoose = require("mongoose");

// YYYY-MM-DD 문자열을 day로 저장(정렬/범위 필터 쉬움)
const TotalsSchema = new mongoose.Schema(
  {
    all: { type: Number, default: 0 },
    success: { type: Number, default: 0 },
    fail: { type: Number, default: 0 },
    fallback: { type: Number, default: 0 },
  },
  { _id: false },
);

const RangeSchema = new mongoose.Schema(
  {
    from: { type: Date },
    to: { type: Date },
  },
  { _id: false },
);

const DailySummarySchema = new mongoose.Schema(
  {
    day: { type: String, required: true, unique: true, index: true }, // "2026-02-01"

    range: { type: RangeSchema, default: {} },
    totals: { type: TotalsSchema, default: {} },

    // distributions/status/engine/validatorKind 등은 구조가 바뀔 수 있으니 Mixed로 유연하게
    distributions: { type: mongoose.Schema.Types.Mixed, default: {} },

    // [{ reason, count }] 형태를 권장 (현재 스냅샷은 empty도 OK)
    reasonTop: { type: Array, default: [] },

    // validator: { failPointsTop: [...], kindTop: [...], ... }
    validator: { type: mongoose.Schema.Types.Mixed, default: {} },
  },
  { timestamps: true },
);

module.exports =
  mongoose.models.DailySummary ||
  mongoose.model("DailySummary", DailySummarySchema);
