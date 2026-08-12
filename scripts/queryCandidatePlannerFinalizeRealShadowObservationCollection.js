const fs = require("fs");
const path = require("path");
const mongoose = require("mongoose");
const {
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceConfig");
const {
  createMongoRealShadowEvidenceStore,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceStore");
const {
  evaluateRealShadowObservationCollection,
  buildObservationCollectionSummary,
  sha256,
} = require("../automation/queryCandidatePlannerRealShadowObservationCollection");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

function atomicWrite(filePath, value) {
  const target = path.resolve(filePath);
  const temp = `${target}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(temp, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  fs.renameSync(temp, target);
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
  const to = arg("--to", new Date().toISOString());
  const output = path.resolve(
    arg(
      "--output",
      "queryCandidatePlannerRealShadowObservationCollection.summary.private.json",
    ),
  );

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
    if (!result.ready) {
      result.errors.forEach((error) => console.error(`BLOCKED ${error}`));
      process.exitCode = 2;
      return;
    }
    const summary = buildObservationCollectionSummary(result, {
      finalizedAt: new Date().toISOString(),
    });
    atomicWrite(output, summary);
    console.log(
      "PASS patch 15.3.2-E real shadow observation collection finalized",
    );
    console.log(`EXECUTIONS ${summary.executionCount}`);
    console.log(`LIFECYCLE ${summary.lifecycleCount}`);
    console.log(`TOTAL_RECORDS ${summary.totalRecordCount}`);
    console.log(`COLLECTION_SHA256 ${sha256(summary)}`);
    console.log("READY_FOR_PATCH_15_3_2_F true");
    console.log("COLLECTOR_ENABLED_BY_THIS_OPERATION false");
    console.log("INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
    console.log("RAW_RECORDS_LOGGED false");
    console.log("FINGERPRINTS_LOGGED false");
    console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
  } finally {
    await mongoose.disconnect();
  }
}

main().catch((error) => {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
});
