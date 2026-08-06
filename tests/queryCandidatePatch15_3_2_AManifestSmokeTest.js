"use strict";

const assert = require("assert");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const root = path.resolve(__dirname, "..");
const manifest = JSON.parse(
  fs.readFileSync(path.join(root, "PATCH_MANIFEST_PATCH15_3_2_A.json"), "utf8"),
);
assert.strictEqual(
  manifest.version,
  "query_candidate_patch15_3_2_A_manifest_v1",
);
assert.strictEqual(manifest.fileCount, manifest.files.length);
for (const entry of manifest.files) {
  const filePath = path.join(root, entry.path);
  assert(fs.existsSync(filePath), `manifest file missing: ${entry.path}`);
  const data = fs.readFileSync(filePath);
  assert.strictEqual(data.length, entry.bytes, `byte mismatch: ${entry.path}`);
  const actual = crypto.createHash("sha256").update(data).digest("hex");
  assert.strictEqual(actual, entry.sha256, `SHA-256 mismatch: ${entry.path}`);
}
console.log(
  `PASS query candidate patch15.3.2-A manifest smoke files=${manifest.fileCount}`,
);
