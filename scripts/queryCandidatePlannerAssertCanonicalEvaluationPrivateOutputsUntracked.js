const { execFileSync } = require("child_process");

const patterns = [
  /queryCandidatePlannerApprovedActualPricingPolicy\.private\.json$/i,
  /queryCandidatePlannerCanonicalEvaluationInput\.private\.json$/i,
  /queryCandidatePlannerCanonicalOperationalReport\.private\.json$/i,
  /queryCandidatePlannerCanonicalEvaluationBaseline\.private\.json$/i,
  /queryCandidatePlannerPatch15_3_2_F\.private/i,
];

function main() {
  let names = [];
  try {
    names = execFileSync(
      "git",
      ["diff", "--cached", "--name-only", "--diff-filter=ACMR"],
      { encoding: "utf8" },
    )
      .split(/\r?\n/)
      .map((item) => item.trim())
      .filter(Boolean);
  } catch {
    names = [];
  }

  const offenders = names.filter((name) =>
    patterns.some((pattern) => pattern.test(name)),
  );
  if (offenders.length) {
    console.error("BLOCKED canonical evaluation private outputs are staged");
    offenders.forEach((name) => console.error(name));
    process.exitCode = 1;
    return;
  }
  console.log(
    "PASS no patch 15.3.2-F.1 canonical evaluation private outputs staged",
  );
}

if (require.main === module) main();
