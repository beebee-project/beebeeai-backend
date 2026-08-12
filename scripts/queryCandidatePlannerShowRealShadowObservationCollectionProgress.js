const fs = require("fs");
const mongoose = require("mongoose");
const {
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceConfig");
const {
  createMongoRealShadowEvidenceStore,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceStore");
const {
  evaluateRealShadowObservationCollection,
} = require("../automation/queryCandidatePlannerRealShadowObservationCollection");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

async function main() {
  const mongoUri = process.env.MONGO_URI || process.env.MONGODB_URI;
  if (!mongoUri) throw new Error("MONGO_URI or MONGODB_URI is required");
  const config = parseQueryCandidatePlannerRealShadowEvidenceConfig(
    process.env,
  );
  if (!config.secret || config.secret.length < 32) {
    throw new Error(
      "QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET is required",
    );
  }
  const windowPath = arg(
    "--window",
    "queryCandidatePlannerRealShadowObservationCollectionWindow.private.json",
  );
  if (!fs.existsSync(windowPath))
    throw new Error("OBSERVATION_COLLECTION_WINDOW_REQUIRED");
  const window = JSON.parse(fs.readFileSync(windowPath, "utf8"));
  const from = window.startedAt;
  const to = arg("--to", "");
  await mongoose.connect(mongoUri);
  try {
    const store = createMongoRealShadowEvidenceStore({ secret: config.secret });
    const records = await store.list({ from, to, limit: config.maxRecords });
    const result = evaluateRealShadowObservationCollection({
      records,
      env: process.env,
      from,
      to,
    });
    for (const item of result.caseSummaries) {
      console.log(
        `${item.protocolReady ? "READY" : "PENDING"} ${item.caseId} ` +
          `EXEC=${item.executionCount}/5 DOWNLOAD=${item.downloadCount}/1 ` +
          `DELETE=${item.deleteCount}/1 IDENTITIES=${item.distinctUploadIdentityCount}/2`,
      );
    }
    console.log(`EXECUTIONS ${result.executionCount}/50`);
    console.log(`LIFECYCLE ${result.lifecycleCount}/20`);
    console.log(`TOTAL_RECORDS ${result.totalRecordCount}`);
    console.log(`BUILDER_MINIMUM_READY ${result.builderMinimumReady}`);
    console.log(`COLLECTION_PROTOCOL_COMPLETE ${result.protocolReady}`);
    console.log(`PRIVACY_VIOLATIONS ${result.privacyViolationCount}`);
    console.log(`GUARDRAIL_VIOLATIONS ${result.guardrailViolationCount}`);
    console.log(`READY_FOR_PATCH_15_3_2_F ${result.readyForPatch15_3_2_F}`);
    console.log("RAW_RECORDS_LOGGED false");
    console.log("FINGERPRINTS_LOGGED false");
  } finally {
    await mongoose.disconnect();
  }
}

main().catch((error) => {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
});
