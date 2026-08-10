#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  generateRealShadowEvidenceSecret,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}

function privateEnvPath(value) {
  const output = path.resolve(value || "queryCandidatePlannerRealShadowEvidenceSecret.private.env");
  if (!/\.private\.env$/i.test(path.basename(output))) {
    const error = new Error("secret output must use a .private.env filename");
    error.code = "REAL_SHADOW_SECRET_PRIVATE_OUTPUT_REQUIRED";
    throw error;
  }
  return output;
}

try {
  const output = privateEnvPath(arg("--output"));
  const force = process.argv.includes("--force");
  if (fs.existsSync(output) && !force) {
    const error = new Error("secret output already exists; explicit --force required for rotation");
    error.code = "REAL_SHADOW_SECRET_OUTPUT_ALREADY_EXISTS";
    throw error;
  }
  const generated = generateRealShadowEvidenceSecret();
  const content = [
    `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET=${generated.secret}`,
    `QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET_SHA256=${generated.secretSha256}`,
    "QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_ENABLED=0",
    "",
  ].join("\n");
  fs.writeFileSync(output, content, { encoding: "utf8", mode: 0o600 });
  try { fs.chmodSync(output, 0o600); } catch (_error) {}
  console.log("PASS real shadow evidence secret private file created");
  console.log(`OUTPUT ${output}`);
  console.log(`SECRET_SHA256 ${generated.secretSha256}`);
  console.log(`ENTROPY_BYTES ${generated.entropyBytes}`);
  console.log(`SECRET_FORMAT ${generated.format}`);
  console.log("RAW_SECRET_LOGGED false");
  console.log("COLLECTOR_ENABLED_BY_THIS_OPERATION false");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
