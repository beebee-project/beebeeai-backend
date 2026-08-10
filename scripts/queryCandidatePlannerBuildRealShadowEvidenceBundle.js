const fs = require("fs");
const path = require("path");
const {
  buildQueryCandidatePlannerRealShadowEvidenceBundle,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceBundleBuilder");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

function requiredArg(name) {
  const value = arg(name);
  if (!value) throw new Error(`${name} is required`);
  return path.resolve(value);
}

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function writeJson(filePath, value) {
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function main() {
  const root = path.resolve(__dirname, "..");
  const recordsFile = requiredArg("--records");
  const readinessFile = requiredArg("--readiness");
  const pricingFile = requiredArg("--pricing");
  const outputDir = path.resolve(
    arg("--output-dir", "real-shadow-evidence-output"),
  );
  const exportValue = readJson(recordsFile);
  const result = buildQueryCandidatePlannerRealShadowEvidenceBundle({
    records: exportValue.records || exportValue,
    readiness: readJson(readinessFile),
    approvedActualPricingPolicy: readJson(pricingFile),
    accuracyDataset: readJson(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
    ),
    accuracyThresholdPolicy: readJson(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyThresholdPolicy.v1.json",
      ),
    ),
    operationalThresholdPolicy: readJson(
      path.join(
        root,
        "evaluation/queryCandidatePlannerOperationalThresholdPolicy.v1.json",
      ),
    ),
    shadowThresholdPolicy: readJson(
      path.join(
        root,
        "evaluation/queryCandidatePlannerShadowAccuracyThresholdPolicy.v1.json",
      ),
    ),
    evaluatedAt: new Date().toISOString(),
    expiresInHours: Number(arg("--expires-hours", "24")),
  });
  fs.mkdirSync(outputDir, { recursive: true });
  writeJson(
    path.join(
      outputDir,
      "queryCandidatePlannerRealShadowEvidenceBuildResult.json",
    ),
    result,
  );
  if (result.decision !== "EVALUATION_PASS") {
    console.error(`BLOCKED ${result.reason}`);
    process.exitCode = 2;
    return;
  }
  writeJson(
    path.join(
      outputDir,
      "queryCandidatePlannerInternalCanaryEvidenceBundle.json",
    ),
    result.evidenceBundle,
  );
  writeJson(
    path.join(outputDir, "queryCandidatePlannerRealShadowAccuracyReport.json"),
    result.reports.accuracy,
  );
  writeJson(
    path.join(
      outputDir,
      "queryCandidatePlannerRealShadowOperationalReport.json",
    ),
    result.reports.operational,
  );
  writeJson(
    path.join(
      outputDir,
      "queryCandidatePlannerRealShadowEvaluationReport.json",
    ),
    result.reports.shadow,
  );
  writeJson(
    path.join(
      outputDir,
      "queryCandidatePlannerRealShadowOperationalDataset.json",
    ),
    result.datasets.operational,
  );
  writeJson(
    path.join(
      outputDir,
      "queryCandidatePlannerRealShadowObservationDataset.json",
    ),
    result.datasets.shadow,
  );
  fs.writeFileSync(
    path.join(outputDir, "railway-evidence-variable.txt"),
    `QUERY_CANDIDATE_PLANNER_INTERNAL_CANARY_EVIDENCE_JSON=${JSON.stringify(result.evidenceBundle)}\n`,
    "utf8",
  );
  console.log(
    `PASS real shadow evidence bundle sha256=${result.evidenceSha256} output=${outputDir}`,
  );
}

try {
  main();
} catch (error) {
  console.error(`FAIL ${error.message}`);
  process.exitCode = 1;
}
