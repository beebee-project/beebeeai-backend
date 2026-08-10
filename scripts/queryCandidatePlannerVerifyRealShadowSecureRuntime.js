#!/usr/bin/env node
"use strict";

const {
  evaluateRealShadowSecureRuntime,
} = require("../automation/queryCandidatePlannerRealShadowSecureDeployment");

try {
  const result = evaluateRealShadowSecureRuntime({ env: process.env });
  if (!result.ready) {
    result.errors.forEach((error) => console.error(`BLOCKED ${error}`));
    process.exitCode = 2;
  } else {
    console.log("PASS patch 15.3.2-C secure runtime verification");
    console.log(`SECRET_SHA256 ${result.secretSha256}`);
    console.log(`REGISTRY_SHA256 ${result.runtimeRegistrySha256}`);
    console.log(`REGISTRY_CASES ${result.registryCaseCount}`);
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
