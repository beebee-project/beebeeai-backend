const fs = require("fs");
const path = require("path");
const {
  finalizeUploadableSourceCatalog,
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
  return path.resolve(value);
}
function privateOutput(name, fallback) {
  const output = path.resolve(arg(name, fallback));
  if (!/\.private\./i.test(path.basename(output))) {
    const error = new Error(`${name} must use a .private. filename`);
    error.code = "REAL_SHADOW_PRIVATE_OUTPUT_NAME_REQUIRED";
    throw error;
  }
  return output;
}

try {
  const root = path.resolve(__dirname, "..");
  const draftPath = required("--catalog");
  const privateOutputPath = privateOutput(
    "--output",
    "queryCandidatePlannerRealShadowUploadableSourceCatalog.private.json",
  );
  const summaryOutputPath = privateOutput(
    "--summary-output",
    "queryCandidatePlannerRealShadowUploadableSourceCatalog.summary.private.json",
  );
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
  const catalog = JSON.parse(fs.readFileSync(draftPath, "utf8"));
  const result = finalizeUploadableSourceCatalog({ accuracyDataset, catalog });
  if (!result.valid) {
    result.errors.forEach((error) => console.error(`BLOCKED ${error}`));
    process.exitCode = 2;
  } else {
    fs.writeFileSync(
      privateOutputPath,
      `${JSON.stringify(result.privateCatalog, null, 2)}\n`,
      { encoding: "utf8", mode: 0o600 },
    );
    fs.writeFileSync(
      summaryOutputPath,
      `${JSON.stringify(
        {
          version: result.publicCatalog.version,
          decision: "REAL_SHADOW_UPLOADABLE_SOURCE_CATALOG_PASS",
          catalogId: result.publicCatalog.catalogId,
          sourceCatalogSha256: result.sourceCatalogSha256,
          caseCount: result.completedCount,
          sourceKinds: [
            ...new Set(
              result.publicCatalog.cases.map((item) => item.sourceKind),
            ),
          ],
          synthetic: false,
          rawWorkbookDataIncluded: false,
          privatePathsIncluded: false,
          collectorEnabledByThisOperation: false,
          internalCanaryEnabledByThisOperation: false,
          productionPromotionAuthorized: false,
        },
        null,
        2,
      )}\n`,
      { encoding: "utf8", mode: 0o600 },
    );
    console.log(
      `PASS uploadable source catalog finalized sha256=${result.sourceCatalogSha256} cases=${result.completedCount}`,
    );
    console.log(`OUTPUT ${privateOutputPath}`);
    console.log(`SUMMARY ${summaryOutputPath}`);
    console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
    console.log("LEGACY_LEDGER_ACCEPTED false");
    console.log("COLLECTOR_ENABLED_BY_THIS_OPERATION false");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
