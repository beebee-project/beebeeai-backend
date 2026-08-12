const { execFileSync } = require("child_process");

const forbidden = [
  /queryCandidatePlannerApprovedActualPricingPolicy.*\.private\.json$/i,
  /real-shadow-evidence-records.*\.private\.json$/i,
  /queryCandidatePlannerRealShadowEvaluationBaseline.*\.private\.json$/i,
  /real-shadow-evidence-output.*private/i,
  /queryCandidatePlannerInternalCanaryEvidenceBundle.*\.private\.json$/i,
];

function stagedNames() {
  try {
    return execFileSync(
      "git",
      ["diff", "--cached", "--name-only", "--diff-filter=ACMR"],
      {
        encoding: "utf8",
      },
    )
      .split(/\r?\n/)
      .map((item) => item.trim())
      .filter(Boolean);
  } catch {
    return [];
  }
}

function main() {
  const offenders = stagedNames().filter((name) =>
    forbidden.some((re) => re.test(name)),
  );
  if (offenders.length) {
    console.error("BLOCKED private evaluation outputs are staged");
    for (const offender of offenders) console.error(offender);
    process.exitCode = 1;
    return;
  }
  console.log("PASS no patch 15.3.2-F private evaluation outputs staged");
}

if (require.main === module) main();
