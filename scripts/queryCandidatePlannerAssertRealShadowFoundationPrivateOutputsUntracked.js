#!/usr/bin/env node
"use strict";

const { execFileSync } = require("child_process");

const patterns = [
  /queryCandidatePlannerRealShadowFingerprintLedger.*\.private\.json$/i,
  /queryCandidatePlannerRealShadowCaseRegistry.*\.private\.(json|txt)$/i,
  /queryCandidatePlannerRealShadowUploadableSourceCatalog.*\.private\.json$/i,
  /queryCandidatePlannerRealShadowExpectedRejectionAttestation.*\.private\.json$/i,
  /queryCandidatePlannerRealShadowEvidenceFoundation.*\.private\.json$/i,
  /queryCandidatePlannerRealShadowEvidenceSecret\.private\.txt$/i,
];

try {
  const staged = execFileSync(
    "git",
    ["diff", "--cached", "--name-only", "--diff-filter=ACMR"],
    { encoding: "utf8" },
  )
    .split(/\r?\n/)
    .map((item) => item.trim())
    .filter(Boolean);

  const violations = staged.filter((file) =>
    patterns.some((pattern) => pattern.test(file.replace(/\\/g, "/"))),
  );
  if (violations.length > 0) {
    violations.forEach((file) => console.error(`BLOCKED staged private file: ${file}`));
    process.exitCode = 2;
  } else {
    console.log("PASS no phase 15.3-A private outputs staged");
  }
} catch (error) {
  console.error(`FAIL ${error.message}`);
  process.exitCode = 1;
}
