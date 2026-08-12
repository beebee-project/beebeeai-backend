const { execFileSync } = require("child_process");

function main() {
  let staged = [];
  try {
    staged = execFileSync(
      "git",
      ["diff", "--cached", "--name-only", "--diff-filter=ACMR"],
      { encoding: "utf8" },
    )
      .split(/\r?\n/)
      .map((value) => value.trim())
      .filter(Boolean);
  } catch {
    staged = [];
  }

  const pattern =
    /queryCandidatePlannerActualPricing(RecalibratedThresholdPolicy|AbsoluteCostRecalibrationEvidence|RecalibratedOperationalReport|AbsoluteCostAssessment|RecalibratedEvaluationBaseline)\.private\.json$/i;

  const offenders = staged.filter((name) => pattern.test(name));

  if (offenders.length) {
    console.error("BLOCKED Patch 15.3.2-F.1.3 private outputs staged");
    offenders.forEach((name) => console.error(name));
    process.exitCode = 1;
    return;
  }

  console.log("PASS no Patch 15.3.2-F.1.3 private outputs staged");
}

if (require.main === module) main();
