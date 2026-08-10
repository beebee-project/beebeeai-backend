const fs = require("fs");
const path = require("path");

const root = process.cwd();
const mustNotExist = [
  "automation/queryCandidatePlannerRealShadowObservationCollection.js",
  "scripts/queryCandidatePlannerFinalizeRealShadowObservationCollection.js",
  "scripts/queryCandidatePlannerShowRealShadowObservationCollectionProgress.js",
  "scripts/queryCandidatePlannerStartRealShadowObservationCollection.js",
  "scripts/queryCandidatePlannerAssertRealShadowObservationCollectionPrivateOutputsUntracked.js",
  "README_QUERY_CANDIDATE_PATCH15_3_2_E_REAL_SHADOW_OBSERVATION_COLLECTION.md",
  "PATCH_MANIFEST_PATCH15_3_2_E.json",
  "PATCH_VALIDATION_PATCH15_3_2_E.json",
];

for (const rel of mustNotExist) {
  if (fs.existsSync(path.join(root, rel))) {
    console.error(`FAIL PATCH_E_FILE_REMAINS ${rel}`);
    process.exit(1);
  }
}

const eTests = fs.existsSync(path.join(root, "tests"))
  ? fs
      .readdirSync(path.join(root, "tests"))
      .filter((name) => /^queryCandidatePatch15_3_2_E(?!_X)/.test(name))
  : [];
if (eTests.length > 0) {
  console.error(`FAIL PATCH_E_TESTS_REMAIN count=${eTests.length}`);
  process.exit(1);
}

const mustExist = [
  "automation/queryCandidatePlannerRealShadowEvidenceConfig.js",
  "automation/queryCandidatePlannerRealShadowLimitedActivation.js",
  "scripts/queryCandidatePlannerVerifyRealShadowLimitedActivationRuntime.js",
  "PATCH_ROADMAP_OVERRIDE_PATCH15_3_2_E_X.json",
];
for (const rel of mustExist) {
  if (!fs.existsSync(path.join(root, rel))) {
    console.error(`FAIL REQUIRED_PRE_E_OR_OVERRIDE_FILE_MISSING ${rel}`);
    process.exit(1);
  }
}

for (const rel of [
  "routes/automationRoutes.js",
  "routes/fileRoutes.js",
  "automation/queryCandidatePlannerRealShadowEvidenceCollector.js",
]) {
  const abs = path.join(root, rel);
  if (!fs.existsSync(abs)) continue;
  const source = fs.readFileSync(abs, "utf8");
  if (
    source.includes("[real-shadow-evidence]") ||
    source.includes("[real-shadow-subject]")
  ) {
    console.error(`FAIL TEMP_DIAGNOSTIC_MARKER_REMAINS ${rel}`);
    process.exit(1);
  }
}

const privateWindows = fs
  .readdirSync(root)
  .filter((name) =>
    /^queryCandidatePlannerRealShadowObservationCollection.*\.private\.json$/i.test(
      name,
    ),
  );
if (privateWindows.length > 0) {
  console.error(
    `FAIL PATCH_E_PRIVATE_WINDOW_REMAINS count=${privateWindows.length}`,
  );
  process.exit(1);
}

const override = JSON.parse(
  fs.readFileSync(
    path.join(root, "PATCH_ROADMAP_OVERRIDE_PATCH15_3_2_E_X.json"),
    "utf8",
  ),
);
if (
  override?.decision?.status !== "EXCLUDED_NA" ||
  override?.decision?.prerequisiteForSubsequentPatches !== false ||
  override?.activeSequence?.[2]?.startsWith("15.3.2-F") !== true
) {
  console.error("FAIL ROADMAP_OVERRIDE_INVALID");
  process.exit(1);
}

console.log("PASS patch 15.3.2-E-X pre-E restore verification");
console.log("PATCH_15_3_2_D_PRESERVED true");
console.log("PATCH_15_3_2_E_STATUS EXCLUDED_NA");
console.log("PATCH_15_3_2_E_FILES_PRESENT false");
console.log("PATCH_15_3_2_E_PRIVATE_OUTPUTS_PRESENT false");
console.log("TEMP_DIAGNOSTIC_MARKERS_PRESENT false");
console.log("NEXT_PATCH 15.3.2-F");
console.log("INTERNAL_ALLOWLIST_INDEPENDENT_GATE_REQUIRED_BEFORE_15_3_3 true");
console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
