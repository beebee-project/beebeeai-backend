const fs = require("fs");
const path = require("path");
const {
  buildRealShadowCaseRegistryScaffold,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

try {
  const root = path.resolve(__dirname, "..");
  const dataset = JSON.parse(
    fs.readFileSync(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
  const output = path.resolve(
    arg(
      "--output",
      "evaluation/queryCandidatePlannerRealShadowCaseRegistry.draft.json",
    ),
  );
  const registryId = arg("--registry-id", "internal_real_shadow_2026_08_v1");
  const scaffold = buildRealShadowCaseRegistryScaffold(dataset, { registryId });
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, `${JSON.stringify(scaffold, null, 2)}\n`, "utf8");
  console.log(
    `PASS real shadow registry scaffold cases=${scaffold.cases.length} output=${output}`,
  );
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
