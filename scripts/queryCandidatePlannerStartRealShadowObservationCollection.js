const fs = require("fs");
const path = require("path");
const {
  evaluateRealShadowLimitedActivationRuntime,
} = require("../automation/queryCandidatePlannerRealShadowLimitedActivation");

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

try {
  const runtime = evaluateRealShadowLimitedActivationRuntime({
    env: process.env,
  });
  if (!runtime.ready) {
    runtime.errors.forEach((error) => console.error(`BLOCKED ${error}`));
    process.exitCode = 2;
  } else {
    const output = path.resolve(
      arg(
        "--output",
        "queryCandidatePlannerRealShadowObservationCollectionWindow.private.json",
      ),
    );
    if (fs.existsSync(output)) {
      throw new Error("OBSERVATION_COLLECTION_WINDOW_ALREADY_EXISTS");
    }
    const startedAt = new Date().toISOString();
    atomicWrite(output, {
      version:
        "query_candidate_planner_real_shadow_observation_collection_window_v1",
      phase: "15.3-B",
      patch: "15.3.2-E",
      startedAt,
      registryCaseCount: runtime.registryCaseCount,
      allowlistEntryCount: runtime.allowlistEntryCount,
      collectorEnabled: true,
      collectorKillSwitch: false,
      privateOutputDoNotCommit: true,
    });
    console.log("PASS patch 15.3.2-E observation collection window started");
    console.log(`STARTED_AT ${startedAt}`);
    console.log(`OUTPUT ${output}`);
    console.log("COLLECTOR_ENABLED true");
    console.log("COLLECTOR_KILL_SWITCH false");
    console.log("INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
    console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
