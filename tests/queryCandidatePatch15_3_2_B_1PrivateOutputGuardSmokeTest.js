"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const source = fs.readFileSync(
  path.join(
    __dirname,
    "../scripts/queryCandidatePlannerAssertRealShadowPrivateOutputsUntracked.js",
  ),
  "utf8",
);
assert(source.includes("RealShadowUploadableSourceCatalog"));
assert(source.includes("FingerprintLedger.*\\.private"));
console.log("PASS query candidate patch15.3.2-B.1 private output guard smoke");
