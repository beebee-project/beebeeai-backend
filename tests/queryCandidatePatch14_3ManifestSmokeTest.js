"use strict";

const assert = require("assert");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const manifestPath = path.join(ROOT, "PATCH_MANIFEST_PATCH14_3.json");
assert(fs.existsSync(manifestPath), "PATCH_MANIFEST_PATCH14_3.json is required");
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
assert.strictEqual(manifest.version, "query_candidate_patch14_3_manifest_v1");
assert.strictEqual(manifest.patch, "Internal UI Preview");
assert(Array.isArray(manifest.files));
assert.strictEqual(manifest.fileCount, manifest.files.length);
assert(manifest.fileCount >= 20);

function loadSuccessor(fileName, expectedVersion, supersedeKey) {
  const fullPath = path.join(ROOT, fileName);
  if (!fs.existsSync(fullPath)) return { files: new Map(), paths: new Set() };
  const value = JSON.parse(fs.readFileSync(fullPath, "utf8"));
  assert.strictEqual(value.version, expectedVersion);
  return {
    files: new Map((value.files || []).map((entry) => [entry.path, entry])),
    paths: new Set(value.supersedes?.[supersedeKey] || []),
  };
}
const patch15_3_2 = loadSuccessor("PATCH_MANIFEST_PATCH15_3_2.json", "query_candidate_patch15_3_2_manifest_v1", "patch14_3Files");
const patch15_3 = loadSuccessor("PATCH_MANIFEST_PATCH15_3.json", "query_candidate_patch15_3_manifest_v1", "patch14_3Files");

let superseded = 0;
for (const entry of manifest.files) {
  const fullPath = path.join(ROOT, entry.path);
  assert(fs.existsSync(fullPath), `manifest file missing: ${entry.path}`);
  const bytes = fs.readFileSync(fullPath);
  const hash = crypto.createHash("sha256").update(bytes).digest("hex");
  if (hash === entry.sha256 && bytes.length === entry.bytes) continue;
  const matched = [patch15_3_2, patch15_3].some((successor) => {
    if (!successor.paths.has(entry.path)) return false;
    const candidate = successor.files.get(entry.path);
    return Boolean(candidate && candidate.sha256 === hash && candidate.bytes === bytes.length);
  });
  assert(matched, `unexpected Patch 14.3 manifest drift: ${entry.path}`);
  superseded += 1;
}

assert.deepStrictEqual(
  manifest.supersedes.patch14_2Files.slice().sort(),
  ["routes/automationRoutes.js", "tests/queryCandidatePatch14_2ManifestSmokeTest.js"],
);
assert.strictEqual(manifest.access.defaultEnabled, false);
assert.strictEqual(manifest.access.authenticatedRouteRequired, true);
assert.strictEqual(manifest.access.internalTokenRequired, true);
assert.strictEqual(manifest.access.queryTokenAccepted, false);
assert.strictEqual(manifest.storage.persistence, "MEMORY_ONLY");
assert.strictEqual(manifest.guardrails.readOnly, true);
assert.strictEqual(manifest.guardrails.observationOnly, true);
assert.strictEqual(manifest.guardrails.productionCandidateMerge, false);
assert.strictEqual(manifest.guardrails.productionReadyAssignment, false);
assert.strictEqual(manifest.guardrails.productionRouteChanged, false);
assert.strictEqual(manifest.guardrails.candidateExecutionAvailable, false);
assert.strictEqual(manifest.guardrails.candidateSelectionAvailable, false);

console.log(`PASS query candidate patch14.3 manifest smoke superseded=${superseded}`);
