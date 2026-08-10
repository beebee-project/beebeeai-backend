const fs = require("fs");
const path = require("path");
const {
  buildRealShadowCaseRegistry,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

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

try {
  const root = path.resolve(__dirname, "..");
  const draftPath = requiredArg("--draft");
  const outputPath = path.resolve(
    arg("--output", "queryCandidatePlannerRealShadowCaseRegistry.json"),
  );
  const railwayOutputPath = path.resolve(
    arg(
      "--railway-output",
      "queryCandidatePlannerRealShadowCaseRegistry.railway.txt",
    ),
  );
  const summaryOutputPath = path.resolve(
    arg(
      "--summary-output",
      "queryCandidatePlannerRealShadowCaseRegistry.summary.json",
    ),
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
  const draft = JSON.parse(fs.readFileSync(draftPath, "utf8"));
  const result = buildRealShadowCaseRegistry({
    accuracyDataset,
    draft,
    requireUploadFingerprint: true,
  });
  if (!result.valid) {
    for (const error of result.errors) console.error(`BLOCKED ${error}`);
    process.exitCode = 2;
  } else {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(
      outputPath,
      `${JSON.stringify(result.registry, null, 2)}\n`,
      "utf8",
    );
    fs.writeFileSync(
      railwayOutputPath,
      `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_CASE_REGISTRY_JSON=${JSON.stringify(result.registry)}\n`,
      "utf8",
    );
    fs.writeFileSync(
      summaryOutputPath,
      `${JSON.stringify(
        {
          version: result.version,
          decision: "PREPARATION_PASS",
          registryId: result.registry.registryId,
          registrySha256: result.registrySha256,
          caseCount: result.caseCount,
          requestFingerprintCount: result.requestFingerprintCount,
          uploadFingerprintCount: result.uploadFingerprintCount,
          actualTrafficOnly: true,
          syntheticFingerprintAllowed: false,
          rawIdentityIncluded: false,
        },
        null,
        2,
      )}\n`,
      "utf8",
    );
    console.log(
      `PASS real shadow registry sha256=${result.registrySha256} cases=${result.caseCount}`,
    );
    console.log(`OUTPUT ${outputPath}`);
    console.log(`RAILWAY ${railwayOutputPath}`);
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
