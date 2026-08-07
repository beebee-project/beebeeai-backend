"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const expected = Object.freeze({
  "automation/queryCandidatePlannerRealShadowUploadableSourceCatalog.js": "375833e0619a83fc0b3fbfce333dd954a5ac0b82e2f39b53e19612f81e2c9784",
  "automation/queryCandidatePlannerRealShadowRegistryFinalization.js": "cb4dc4a5e8a11aeb64f4ef0d6e00bfa565610a9825a26eab5772ebf66136304f",
  "scripts/queryCandidatePlannerScaffoldRealShadowUploadableSourceCatalog.js": "4777e10be1593510499770782affc16bd8c88e12f967bf2bdf3ffb56ebebe55f",
  "scripts/queryCandidatePlannerBindRealShadowUploadableSource.js": "2a4953ca696450c285d347eeee522ebd7050f69930b2344ff3a787bd700d5c64",
  "scripts/queryCandidatePlannerShowRealShadowUploadableSourceProgress.js": "21ed321a93803980e5dfb0ee548d835d3c7fcc78e2c4d3ac1e151da577593447",
  "scripts/queryCandidatePlannerFinalizeRealShadowUploadableSourceCatalog.js": "39b8d1093de17e8defc99f97bf82274c88b100cac8bf7864a98acbb0d6b2b822",
  "scripts/queryCandidatePlannerInvalidateLegacyRealShadowLedger.js": "9348a24331b5228b9d361fdb133df2393249fe94706b0b865c957da2f1b150f7",
  "scripts/queryCandidatePlannerScaffoldRealShadowFingerprintLedger.js": "5e222ccd4def79eaca37020d8e84e4ab542d4de350707a9b744fd05dc73b9de9",
  "scripts/queryCandidatePlannerRecordRealShadowFingerprint.js": "1b2938ea9ece0fdfe2c41fa45b42e25a7b6aca710e9c212e0d0d95e5028d41e2",
  "scripts/queryCandidatePlannerShowRealShadowRegistryProgress.js": "1cf1465e9470ee5607f724b9cc1ef6ddb1eff5005152f8c98dce99e0b8d06efc",
  "scripts/queryCandidatePlannerFinalizeRealShadowCaseRegistry.js": "12c8091729e670e7ab933c162d3ed0161a0095f19932e2f8943b9c9db4445cea",
  "scripts/queryCandidatePlannerAssertRealShadowPrivateOutputsUntracked.js": "b933d4b41b7627ffd8356940109e3c6abd31da4fc61db59264f27d8d5fe380fa",
  "evaluation/queryCandidatePlannerRealShadowUploadableSourceCatalog.template.json": "6be68fd8d6c983e3e3fdf331010b4889459ec77ef2909693fc0ed39522da2fe6",
  "evaluation/queryCandidatePlannerRealShadowFingerprintLedger.v2.template.json": "09dd6220e9103a23f9196a26bc8f5ba847fb03b668e55f2bfc7f88eb06f76b6e",
});

for (const [relative, expectedHash] of Object.entries(expected)) {
  const content = fs.readFileSync(path.join(__dirname, "..", relative));
  const actual = crypto.createHash("sha256").update(content).digest("hex");
  assert.strictEqual(actual, expectedHash, `SHA-256 mismatch: ${relative}`);
}
console.log("PASS query candidate patch15.3.2-B.1 source integrity smoke");
