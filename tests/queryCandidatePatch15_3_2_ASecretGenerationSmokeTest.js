"use strict";

const assert = require("assert");
const crypto = require("crypto");
const {
  generateRealShadowEvidenceSecret,
} = require("../automation/queryCandidatePlannerRealShadowPreparation");

const result = generateRealShadowEvidenceSecret({
  randomBytes: (size) => Buffer.alloc(size, 0x5a),
});
assert.strictEqual(result.secret.length, 64);
assert.match(result.secret, /^[A-Za-z0-9_-]{64}$/);
assert.strictEqual(
  result.secretSha256,
  crypto.createHash("sha256").update(result.secret).digest("hex"),
);
assert.strictEqual(result.entropyBytes, 48);
assert.strictEqual(result.reusableWithOtherSecrets, false);
console.log("PASS query candidate patch15.3.2-A secret generation smoke");
