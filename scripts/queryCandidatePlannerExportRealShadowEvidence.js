#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const mongoose = require("mongoose");
const {
  parseQueryCandidatePlannerRealShadowEvidenceConfig,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceConfig");
const {
  createMongoRealShadowEvidenceStore,
} = require("../automation/queryCandidatePlannerRealShadowEvidenceStore");

function arg(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}

async function main() {
  const mongoUri = process.env.MONGO_URI || process.env.MONGODB_URI;
  if (!mongoUri) throw new Error("MONGO_URI or MONGODB_URI is required");
  const config = parseQueryCandidatePlannerRealShadowEvidenceConfig(process.env);
  if (!config.secret || config.secret.length < 32) {
    throw new Error("QUERY_CANDIDATE_PLANNER_REAL_SHADOW_EVIDENCE_SECRET is required");
  }
  const output = path.resolve(arg("--output", "real-shadow-evidence-records.json"));
  const from = arg("--from", "");
  const to = arg("--to", "");
  const limit = Number(arg("--limit", String(config.maxRecords || 5000)));
  await mongoose.connect(mongoUri);
  try {
    const store = createMongoRealShadowEvidenceStore({ secret: config.secret });
    const records = await store.list({ from, to, limit });
    const exportValue = {
      version: "query_candidate_planner_real_shadow_evidence_export_v1",
      source: "REAL_SHADOW_TRAFFIC",
      actualTraffic: true,
      synthetic: false,
      exportedAt: new Date().toISOString(),
      recordCount: records.length,
      records,
      privacy: {
        rawRowsIncluded: false,
        fileNamesIncluded: false,
        userIdentityIncluded: false,
        rawProviderResponseIncluded: false,
      },
    };
    fs.writeFileSync(output, `${JSON.stringify(exportValue, null, 2)}\n`, "utf8");
    console.log(`PASS exported real shadow evidence records=${records.length} output=${output}`);
  } finally {
    await mongoose.disconnect();
  }
}

main().catch((error) => {
  console.error(`FAIL ${error.message}`);
  process.exitCode = 1;
});
