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
      .map((line) => line.trim())
      .filter(Boolean);
  } catch {
    staged = [];
  }

  const offenders = staged.filter((name) =>
    /queryCandidatePlannerPatch13_3HistoricalReadinessEvidence\.private\.json$/i.test(
      name,
    ),
  );

  if (offenders.length) {
    console.error("BLOCKED historical readiness private output is staged");
    offenders.forEach((name) => console.error(name));
    process.exitCode = 1;
    return;
  }

  console.log(
    "PASS historical Patch 13.3 readiness private output is not staged",
  );
}

if (require.main === module) main();
