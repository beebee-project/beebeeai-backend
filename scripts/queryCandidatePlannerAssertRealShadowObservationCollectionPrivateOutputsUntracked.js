const { execFileSync } = require("child_process");

const PRIVATE_PATTERNS = [
  /queryCandidatePlannerRealShadowObservationCollectionWindow\.private\.json$/i,
  /queryCandidatePlannerRealShadowObservationCollection\.summary\.private\.json$/i,
  /real-shadow-evidence-records.*\.json$/i,
];

try {
  const staged = execFileSync("git", ["diff", "--cached", "--name-only"], {
    encoding: "utf8",
  })
    .split(/\r?\n/)
    .map((value) => value.trim())
    .filter(Boolean);
  const offenders = staged.filter((name) =>
    PRIVATE_PATTERNS.some((pattern) => pattern.test(name)),
  );
  if (offenders.length > 0) {
    offenders.forEach((name) =>
      console.error(`BLOCKED_PRIVATE_OUTPUT ${name}`),
    );
    process.exitCode = 2;
  } else {
    console.log("PASS no observation-collection private outputs staged");
  }
} catch (error) {
  console.error(`FAIL ${error.message}`);
  process.exitCode = 1;
}
