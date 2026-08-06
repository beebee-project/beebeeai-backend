"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const expected = Object.freeze({
  "automation/queryCandidatePlannerRealShadowRegistryFinalization.js":
    "6b5489473e09ba0ffbb2208f0d5dc5c639b88d89a45a4f6f60bb005a2cb20568",
  "scripts/queryCandidatePlannerScaffoldRealShadowFingerprintLedger.js":
    "a761c1a031b8be8abe580dbddc18245efaef721bcd39a6a44056c15d193b0f49",
  "scripts/queryCandidatePlannerRecordRealShadowFingerprint.js":
    "8f4b2085c8dbbc0183c682b42ed74b30b60fa7c50fa03b7d937d491a461fce4a",
  "scripts/queryCandidatePlannerShowRealShadowRegistryProgress.js":
    "285fa82de00753194eea2d80f976e8b22801d10367855cdf9578f506901d0681",
  "scripts/queryCandidatePlannerFinalizeRealShadowCaseRegistry.js":
    "1e335f8bdb7f21b9a0b036939055f76c3c98e78a545904ff93d535a7377362b4",
  "scripts/queryCandidatePlannerAssertRealShadowPrivateOutputsUntracked.js":
    "d694899c71b5f7c0dac4f2c86914988bc46b18f5bfb514e13f7e9c8eaf9825f8",
  "evaluation/queryCandidatePlannerRealShadowFingerprintLedger.template.json":
    "d37c996ffcb5031b08ca6387cb1868aee11415d356c2f5cce6bfd608390b4b1a",
});

for (const [relative, expectedHash] of Object.entries(expected)) {
  const content = fs.readFileSync(path.join(__dirname, "..", relative));
  const actual = crypto.createHash("sha256").update(content).digest("hex");
  assert.strictEqual(actual, expectedHash, `SHA-256 mismatch: ${relative}`);
}
console.log("PASS query candidate patch15.3.2-B source integrity smoke");
