#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  verifyRealShadowPreparation,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

try {
  const root = path.resolve(__dirname, "..");
  const registryPath = path.resolve(arg("--registry", ""));
  if (!arg("--registry")) throw new Error("--registry is required");
  const secret = process.env.QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET;
  const allowlistSha256 =
    process.env.QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256;
  const accuracyDataset = JSON.parse(
    fs.readFileSync(
      path.join(
        root,
        "evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
  const registry = JSON.parse(fs.readFileSync(registryPath, "utf8"));
  const result = verifyRealShadowPreparation({
    accuracyDataset,
    registry,
    secret,
    allowlistSha256,
    env: process.env,
  });
  if (!result.ready) {
    for (const error of result.errors) console.error(`BLOCKED ${error}`);
    process.exitCode = 2;
  } else {
    console.log(
      `PASS real shadow preparation registrySha256=${result.registrySha256}`,
    );
    console.log(`SECRET_SHA256 ${result.secretSha256}`);
    console.log(`ALLOWLIST_ENTRIES ${result.allowlistEntryCount}`);
    console.log(`CASE_COUNT ${result.caseCount}`);
    console.log("COLLECTOR_ENABLED_BY_THIS_OPERATION false");
    console.log("INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
