#!/usr/bin/env node
"use strict";

const {
  generateRealShadowEvidenceSecret,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

try {
  const result = generateRealShadowEvidenceSecret();
  console.log(
    `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET=${result.secret}`,
  );
  console.log(`SECRET_SHA256=${result.secretSha256}`);
  console.log(
    "Keep the secret local, do not commit it, and do not reuse JWT or file-encryption secrets.",
  );
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
