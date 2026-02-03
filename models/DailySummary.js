const mongoose = require("mongoose");

const DailySummarySchema = new mongoose.Schema(
  {
    day: { type: String, required: true, unique: true, index: true }, // "YYYY-MM-DD" (KST 기준 추천)
    range: {
      from: { type: Date, required: true },
      to: { type: Date, required: true },
    },

    totals: {
      all: Number,
      success: Number,
      fail: Number,
      fallback: Number,
    },

    distributions: {
      engine: Object,
      status: Object,
      validatorKind: Object,
    },

    reasonTop: [
      {
        reason: String,
        count: Number,
      },
    ],

    validator: {
      failPointsTop: [
        {
          code: String,
          count: Number,
        },
      ],
    },
  },
  { timestamps: true },
);

module.exports =
  mongoose.models.DailySummary ||
  mongoose.model("DailySummary", DailySummarySchema);
