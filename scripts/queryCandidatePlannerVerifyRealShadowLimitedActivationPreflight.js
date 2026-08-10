#!/usr/bin/env node
"use strict";

const {
  evaluateRealShadowLimitedActivationPreflight,
} = require("../automation/queryCandidatePlannerRealShadowLimitedActivation");

try {
  const result = evaluateRealShadowLimitedActivationPreflight({ env: process.env });
  if (!result.ready) {
    result.errors.forEach((error) => console.error(`BLOCKED ${error}`));
    process.exitCode = 2;
  } else {
    console.log("PASS patch 15.3.2-D limited collector activation preflight");
    console.log(`REGISTRY_CASES ${result.registryCaseCount}`);
    console.log(`ALLOWLIST_ENTRIES ${result.allowlistEntryCount}`);
    console.log(`TTL_DAYS ${result.ttlDays}`);
    console.log(`MAX_RECORDS ${result.maxRecords}`);
    console.log("COLLECTOR_ENABLED false");
    console.log("COLLECTOR_KILL_SWITCH true");
    console.log("READY_FOR_LIMITED_ACTIVATION true");
    console.log("INTERNAL_CANARY_ENABLED_BY_THIS_OPERATION false");
    console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  }
} catch (error) {
  console.error(`FAIL ${error.code || error.message}`);
  process.exitCode = 1;
}
