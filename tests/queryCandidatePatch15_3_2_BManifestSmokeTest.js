"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const root = path.resolve(__dirname, "..");
const predecessor = JSON.parse(
  fs.readFileSync(path.join(root, "PATCH_MANIFEST_PATCH15_3_2_B.json"), "utf8"),
);
const successor = JSON.parse(
  fs.readFileSync(path.join(root, "PATCH_MANIFEST_PATCH15_3_2_B_1.json"), "utf8"),
);
assert.strictEqual(predecessor.version, "query_candidate_patch15_3_2_B_manifest_v1");
assert.strictEqual(predecessor.fileCount, predecessor.files.length);
const successorByPath = new Map(successor.files.map((entry) => [entry.path, entry]));
let superseded = 0;
for (const entry of predecessor.files) {
  const content = fs.readFileSync(path.join(root, entry.path));
  const actual = crypto.createHash("sha256").update(content).digest("hex");
  if (actual === entry.sha256 && content.length === entry.bytes) continue;
  const replacement = successorByPath.get(entry.path);
  assert(replacement, `unexpected Patch 15.3.2-B manifest drift: ${entry.path}`);
  assert.strictEqual(content.length, replacement.bytes, `byte mismatch: ${entry.path}`);
  assert.strictEqual(actual, replacement.sha256, `SHA-256 mismatch: ${entry.path}`);
  superseded += 1;
}
console.log(`PASS query candidate patch15.3.2-B manifest smoke superseded=${superseded}`);
