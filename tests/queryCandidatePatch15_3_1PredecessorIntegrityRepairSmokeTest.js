'use strict';

const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const repoRoot = path.resolve(__dirname, '..');
const expected = [
  {
    relativePath: 'automation/queryCandidatePlannerControlledProductionPromotionGate.js',
    bytes: 15746,
    sha256: '803e745ab95681b24b25ebefe216adcf9710fd89d237a91f768f38e4e59b7ef6',
  },
  {
    relativePath: 'automation/queryCandidatePlannerApiUiRollbackQualityGate.js',
    bytes: 23875,
    sha256: '1826697f646a535a9db5bc6d76bfd699b5869d496478aae109fa41abf6a8580e',
  },
  {
    relativePath: 'automation/queryCandidatePlannerCostCacheLatencyEvaluator.js',
    bytes: 32317,
    sha256: '67a0ff4d5ac83103d78c9172aa4cc072c008d195a22e99b3366f8440b9d8658c',
  },
  {
    relativePath: 'automation/queryCandidatePlannerShadowAccuracyEvaluator.js',
    bytes: 33926,
    sha256: '0c70432d4ddf838eb4b8d407821d1e44c3e4f89b83564aa554970becb6890f1e',
  },
];

for (const item of expected) {
  const absolutePath = path.join(repoRoot, item.relativePath);
  assert.ok(fs.existsSync(absolutePath), `missing protected source: ${item.relativePath}`);
  const content = fs.readFileSync(absolutePath);
  const hash = crypto.createHash('sha256').update(content).digest('hex');
  assert.strictEqual(content.length, item.bytes, `byte mismatch: ${item.relativePath}`);
  assert.strictEqual(hash, item.sha256, `SHA-256 mismatch: ${item.relativePath}`);
}

console.log('PASS query candidate patch15.3.1 predecessor integrity repair smoke files=4');
