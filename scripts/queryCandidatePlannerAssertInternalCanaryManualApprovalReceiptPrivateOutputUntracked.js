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

  const offenders = staged.filter((name) =>
    /queryCandidatePlannerInternalCanaryManualApprovalReceipt\.private\.json$/i.test(
      name,
    ),
  );

  if (offenders.length) {
    console.error(
      "BLOCKED Patch 15.3.2-F.1.5 private approval receipt is staged",
    );
    offenders.forEach((name) => console.error(name));
    process.exitCode = 1;
    return;
  }

  console.log("PASS Patch 15.3.2-F.1.5 private approval receipt is not staged");
}

if (require.main === module) main();
