const fs = require("fs");
const path = require("path");
const {
  bindUploadableSource,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}
function required(name) {
  const value = arg(name);
  if (!value) throw new Error(`${name} is required`);
  return value;
}
function atomicWrite(filePath, value) {
  const temporary = `${filePath}.${process.pid}.tmp`;
  fs.writeFileSync(temporary, value, { encoding: "utf8", mode: 0o600 });
  fs.renameSync(temporary, filePath);
}

try {
  const root = path.resolve(__dirname, "..");
  const catalogPath = path.resolve(required("--catalog"));
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
  const catalog = JSON.parse(fs.readFileSync(catalogPath, "utf8"));
  const result = bindUploadableSource({
    accuracyDataset,
    catalog,
    caseId: required("--case-id"),
    sourcePath: required("--source-file"),
    sourceKind: required("--source-kind"),
    semanticCompatibilityConfirmed:
      arg("--confirm-semantic-compatibility", "false").toLowerCase() === "true",
    verifiedAt: arg("--verified-at", new Date().toISOString()),
    operatorNote: arg("--operator-note", "real uploadable source verified"),
  });
  atomicWrite(catalogPath, `${JSON.stringify(result.catalog, null, 2)}\n`);
  console.log(`PASS bound source case=${result.recordedCaseId}`);
  console.log(`PROGRESS ${result.completedCount}/10`);
  console.log(`REMAINING ${result.remainingCount}`);
  console.log(`SOURCE_ARTIFACT_SHA256 ${result.sourceArtifactSha256}`);
  console.log(`SOURCE_SIZE_BYTES ${result.sourceSizeBytes}`);
  console.log("RAW_FILE_CONTENT_LOGGED false");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
