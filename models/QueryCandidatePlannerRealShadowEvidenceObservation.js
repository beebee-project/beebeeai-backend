const mongoose = require("mongoose");

const schema = new mongoose.Schema(
  {
    recordId: { type: String, required: true, unique: true, index: true },
    kind: {
      type: String,
      enum: ["EXECUTION", "LIFECYCLE"],
      required: true,
      index: true,
    },
    source: { type: String, enum: ["REAL_SHADOW_TRAFFIC"], required: true },
    actualTraffic: { type: Boolean, required: true },
    synthetic: { type: Boolean, required: true },
    observedAt: { type: Date, required: true, index: true },
    expiresAt: { type: Date, required: true, index: { expires: 0 } },
    subjectTagSha256: { type: String, required: true, index: true },
    requestFingerprintSha256: { type: String, default: "", index: true },
    uploadFingerprintSha256: { type: String, default: "", index: true },
    caseId: { type: String, default: "", index: true },
    scenarioId: { type: String, default: "", index: true },
    encryptionVersion: { type: String, required: true },
    iv: { type: String, required: true },
    authTag: { type: String, required: true },
    ciphertext: { type: String, required: true },
  },
  {
    collection: "query_candidate_planner_real_shadow_evidence",
    timestamps: true,
    strict: true,
    minimize: false,
  },
);

schema.index({ observedAt: 1, kind: 1 });

module.exports =
  mongoose.models.QueryCandidatePlannerRealShadowEvidenceObservation ||
  mongoose.model("QueryCandidatePlannerRealShadowEvidenceObservation", schema);
