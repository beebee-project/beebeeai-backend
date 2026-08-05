"use strict";

const assert = require("assert");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const manifestPath = path.join(ROOT, "PATCH_MANIFEST_PATCH15_3.json");
assert(fs.existsSync(manifestPath), "PATCH_MANIFEST_PATCH15_3.json is required");
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
assert.strictEqual(manifest.version, "query_candidate_patch15_3_manifest_v1");
assert.strictEqual(manifest.patch, "Internal Allowlist Canary");
assert(Array.isArray(manifest.files));
assert.strictEqual(manifest.fileCount, manifest.files.length);
assert(manifest.fileCount >= 45);

const successorPath = path.join(ROOT, "PATCH_MANIFEST_PATCH15_3_2.json");
let successorFiles = new Map();
let supersededPaths = new Set();
if (fs.existsSync(successorPath)) {
  const successor = JSON.parse(fs.readFileSync(successorPath, "utf8"));
  assert.strictEqual(successor.version, "query_candidate_patch15_3_2_manifest_v1");
  successorFiles = new Map((successor.files || []).map((entry) => [entry.path, entry]));
  supersededPaths = new Set(successor.supersedes?.patch15_3Files || []);
}

let superseded = 0;
for (const entry of manifest.files) {
  const fullPath = path.join(ROOT, entry.path);
  assert(fs.existsSync(fullPath), `manifest file missing: ${entry.path}`);
  const bytes = fs.readFileSync(fullPath);
  const hash = crypto.createHash("sha256").update(bytes).digest("hex");
  if (hash === entry.sha256 && bytes.length === entry.bytes) continue;
  assert(supersededPaths.has(entry.path), `unexpected Patch 15.3 manifest drift: ${entry.path}`);
  const successorEntry = successorFiles.get(entry.path);
  assert(successorEntry, `Patch 15.3.2 manifest entry missing: ${entry.path}`);
  assert.strictEqual(hash, successorEntry.sha256, `Patch 15.3.2 SHA-256 mismatch: ${entry.path}`);
  assert.strictEqual(bytes.length, successorEntry.bytes, `Patch 15.3.2 byte size mismatch: ${entry.path}`);
  superseded += 1;
}

assert.deepStrictEqual(
  manifest.supersedes.patch14_2Files.slice().sort(),
  ["routes/automationRoutes.js", "tests/queryCandidatePatch14_2ManifestSmokeTest.js"],
);
assert.strictEqual(manifest.activation.defaultEnabled, false);
assert.strictEqual(manifest.activation.defaultKillSwitch, true);
assert.strictEqual(manifest.activation.audienceMode, "ALLOWLIST_ONLY");
assert.strictEqual(manifest.activation.rolloutPercent, 0);
assert.strictEqual(manifest.evidence.realShadowTrafficRequired, true);
assert.strictEqual(manifest.evidence.syntheticEvidenceAccepted, false);
assert.strictEqual(manifest.llmPolicy.mode, "SEMANTIC_PROFILER_ONLY");
assert.strictEqual(manifest.llmPolicy.plannerEscalationAllowed, false);
assert.strictEqual(manifest.llmPolicy.maxProviderCalls, 1);
assert.strictEqual(manifest.guardrails.routeWired, true);
assert.strictEqual(manifest.guardrails.controllerWired, false);
assert.strictEqual(manifest.guardrails.generalUsersBlocked, true);
assert.strictEqual(manifest.guardrails.primaryFallback, true);
assert.strictEqual(manifest.guardrails.productionReadyAssignment, false);
assert.strictEqual(manifest.guardrails.productionRouteChanged, false);
assert.strictEqual(manifest.guardrails.failClosed, true);

console.log(`PASS query candidate patch15.3 manifest smoke superseded=${superseded}`);
