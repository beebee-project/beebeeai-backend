const {
  encryptEvidencePayload,
  decryptEvidencePayload,
} = require("./queryCandidatePlannerRealShadowEvidenceCrypto");

const STORE_VERSION = "query_candidate_planner_real_shadow_evidence_store_v1";

function modelProvider() {
  return require("../models/QueryCandidatePlannerRealShadowEvidenceObservation");
}

function createMongoRealShadowEvidenceStore({ model = null, secret } = {}) {
  const EvidenceModel = model || modelProvider();
  async function record(record) {
    const encrypted = encryptEvidencePayload(record.payload, secret);
    const document = {
      recordId: record.recordId,
      kind: record.kind,
      source: "REAL_SHADOW_TRAFFIC",
      actualTraffic: true,
      synthetic: false,
      observedAt: new Date(record.observedAt),
      expiresAt: new Date(record.expiresAt),
      subjectTagSha256: record.subjectTagSha256,
      requestFingerprintSha256: record.requestFingerprintSha256 || "",
      uploadFingerprintSha256: record.uploadFingerprintSha256 || "",
      caseId: record.caseId || "",
      scenarioId: record.scenarioId || "",
      ...encrypted,
    };
    try {
      await EvidenceModel.updateOne(
        { recordId: record.recordId },
        { $setOnInsert: document },
        { upsert: true },
      );
      return Object.freeze({ stored: true, recordId: record.recordId });
    } catch (error) {
      return Object.freeze({
        stored: false,
        recordId: record.recordId,
        reason: String(error?.code || "REAL_SHADOW_EVIDENCE_STORE_FAILED"),
      });
    }
  }

  async function list({ from, to, limit = 5000 } = {}) {
    const query = {};
    if (from || to) {
      query.observedAt = {};
      if (from) query.observedAt.$gte = new Date(from);
      if (to) query.observedAt.$lte = new Date(to);
    }
    const docs = await EvidenceModel.find(query)
      .sort({ observedAt: 1, recordId: 1 })
      .limit(Math.max(1, Math.min(50000, Number(limit) || 5000)))
      .lean();
    return Object.freeze(
      docs.map((doc) =>
        Object.freeze({
          recordId: doc.recordId,
          kind: doc.kind,
          source: doc.source,
          actualTraffic: doc.actualTraffic === true,
          synthetic: doc.synthetic === true,
          observedAt: new Date(doc.observedAt).toISOString(),
          expiresAt: new Date(doc.expiresAt).toISOString(),
          subjectTagSha256: doc.subjectTagSha256,
          requestFingerprintSha256: doc.requestFingerprintSha256 || "",
          uploadFingerprintSha256: doc.uploadFingerprintSha256 || "",
          caseId: doc.caseId || "",
          scenarioId: doc.scenarioId || "",
          payload: decryptEvidencePayload(doc, secret),
        }),
      ),
    );
  }

  return Object.freeze({ version: STORE_VERSION, record, list });
}

function createMemoryRealShadowEvidenceStore() {
  const records = new Map();
  return Object.freeze({
    version: `${STORE_VERSION}_memory_test`,
    async record(record) {
      if (!records.has(record.recordId)) records.set(record.recordId, record);
      return Object.freeze({ stored: true, recordId: record.recordId });
    },
    async list() {
      return Object.freeze(
        [...records.values()].sort(
          (a, b) =>
            String(a.observedAt).localeCompare(String(b.observedAt)) ||
            String(a.recordId).localeCompare(String(b.recordId)),
        ),
      );
    },
    clear() {
      records.clear();
    },
  });
}

module.exports = Object.freeze({
  STORE_VERSION,
  createMongoRealShadowEvidenceStore,
  createMemoryRealShadowEvidenceStore,
});
