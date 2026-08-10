#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  evaluateRealShadowSecureDeployment,
} = require("../automation/queryCandidatePlannerRealShadowSecureDeployment");

function arg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : "";
}

function requiredPath(name) {
  const value = arg(name);
  if (!value) throw new Error(`${name} is required`);
  return path.resolve(value);
}

try {
  const foundationSummaryPath = requiredPath("--foundation-summary");
  const registryPath = requiredPath("--registry");
  const foundationSummary = JSON.parse(fs.readFileSync(foundationSummaryPath, "utf8"));
  const registry = JSON.parse(fs.readFileSync(registryPath, "utf8"));
  const result = evaluateRealShadowSecureDeployment({
    foundationSummary,
    registry,
    env: process.env,
  });
  if (!result.ready) {
    for (const error of result.errors) console.error(`BLOCKED ${error}`);
    process.exitCode = 2;
  } else {
    console.log("PASS patch 15.3.2-C secure deployment preflight");
    console.log(`SECRET_SHA256 ${result.secretSha256}`);
    console.log(`REGISTRY_SHA256 ${result.registrySha256}`);
    console.log(`ALLOWLIST_ENTRIES ${result.allowlistEntryCount}`);
    console.log(`ENCRYPTION_VERSION ${result.encryptionSelfTest.encryptionVersion}`);
    console.log("ENCRYPTION_ROUND_TRIP true");
    console.log("WRONG_SECRET_REJECTED true");
    console.log("COLLECTOR_ENABLED false");
    console.log("READY_FOR_PATCH_15_3_2_D true");
    console.log("RAW_SECRET_LOGGED false");
    console.log("INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
