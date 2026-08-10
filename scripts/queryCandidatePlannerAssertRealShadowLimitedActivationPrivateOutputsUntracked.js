const { execFileSync } = require("child_process");

const forbidden = [
  /queryCandidatePlannerRealShadowEvidenceSecret\.private\.env$/,
  /queryCandidatePlannerRealShadowCaseRegistry.*\.private\.(json|txt)$/,
  /queryCandidatePlannerRealShadowEvidenceFoundation.*\.private\.json$/,
  /queryCandidatePlannerRealShadowFingerprintLedger.*\.private\.json$/,
  /queryCandidatePlannerRealShadowUploadableSourceCatalog.*\.private\.json$/,
  /queryCandidatePlannerRealShadowExpectedRejection.*\.private\.json$/,
];

try {
  const staged = execFileSync("git", ["diff", "--cached", "--name-only"], {
    encoding: "utf8",
  })
    .split(/\r?\n/)
    .map((v) => v.trim())
    .filter(Boolean);
  const bad = staged.filter((name) => forbidden.some((re) => re.test(name)));
  if (bad.length) {
    bad.forEach((name) =>
      console.error(`BLOCKED private staged output ${name}`),
    );
    process.exitCode = 2;
  } else {
    console.log("PASS no limited-activation private outputs staged");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
