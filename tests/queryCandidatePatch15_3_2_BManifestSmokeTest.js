"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const root = path.resolve(__dirname, "..");
const manifest = JSON.parse(
  fs.readFileSync(path.join(root, "PATCH_MANIFEST_PATCH15_3_2_B.json"), "utf8"),
);
assert.strictEqual(
  manifest.version,
  "query_candidate_patch15_3_2_B_manifest_v1",
);
assert.strictEqual(manifest.fileCount, manifest.files.length);
for (const entry of manifest.files) {
  const content = fs.readFileSync(path.join(root, entry.path));
  const actual = crypto.createHash("sha256").update(content).digest("hex");
  assert.strictEqual(content.length, entry.bytes, `byte mismatch: ${entry.path}`);
  assert.strictEqual(actual, entry.sha256, `SHA-256 mismatch: ${entry.path}`);
}
console.log(
  `PASS query candidate patch15.3.2-B manifest smoke files=${manifest.fileCount}`,
);
