"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const {
  buildUploadableSourceCatalogScaffold,
  bindUploadableSource,
  finalizeUploadableSourceCatalog,
} = require("../automation/queryCandidatePlannerRealShadowUploadableSourceCatalog");
const {
  buildRealShadowFingerprintLedgerScaffold,
  upsertRealShadowFingerprintCapture,
} = require("../automation/queryCandidatePlannerRealShadowRegistryFinalization");

function accuracyDataset() {
  return JSON.parse(
    fs.readFileSync(
      path.join(
        __dirname,
        "../evaluation/queryCandidatePlannerAccuracyEvaluationDataset.v1.json",
      ),
      "utf8",
    ),
  );
}

function hash(character) {
  return String(character).repeat(64).slice(0, 64);
}

function createSourceFiles(dataset = accuracyDataset()) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "beebee-b1-sources-"));
  const byCaseId = new Map();
  dataset.cases.forEach((item, index) => {
    const filePath = path.join(dir, `${index + 1}.csv`);
    fs.writeFileSync(
      filePath,
      `case_id,metric,value\n${item.caseId},row_${index + 1},${index + 10}\n`,
      "utf8",
    );
    byCaseId.set(item.caseId, filePath);
  });
  return { dir, byCaseId };
}

function completeSourceCatalog() {
  const dataset = accuracyDataset();
  const sources = createSourceFiles(dataset);
  let catalog = buildUploadableSourceCatalogScaffold(dataset, {
    catalogId: "test_real_uploadable_sources_v1",
  });
  dataset.cases.forEach((item, index) => {
    catalog = bindUploadableSource({
      accuracyDataset: dataset,
      catalog,
      caseId: item.caseId,
      sourcePath: sources.byCaseId.get(item.caseId),
      sourceKind: index % 2 === 0 ? "PUBLIC_DATASET" : "REAL_ANONYMIZED",
      semanticCompatibilityConfirmed: true,
      verifiedAt: "2026-08-06T08:00:00.000Z",
      now: Date.parse("2026-08-06T08:30:00.000Z"),
    }).catalog;
  });
  const finalized = finalizeUploadableSourceCatalog({
    accuracyDataset: dataset,
    catalog,
    now: Date.parse("2026-08-06T08:30:00.000Z"),
  });
  if (!finalized.valid) throw new Error(finalized.errors.join("; "));
  return { dataset, sources, catalog: finalized.privateCatalog, finalized };
}

function completeLedger() {
  const support = completeSourceCatalog();
  let ledger = buildRealShadowFingerprintLedgerScaffold(
    support.dataset,
    support.catalog,
    {
      registryId: "test_real_shadow_v2",
      now: Date.parse("2026-08-06T08:30:00.000Z"),
    },
  );
  support.dataset.cases.forEach((item, index) => {
    ledger = upsertRealShadowFingerprintCapture({
      accuracyDataset: support.dataset,
      sourceCatalog: support.catalog,
      ledger,
      sourceFilePath: support.sources.byCaseId.get(item.caseId),
      caseId: item.caseId,
      requestFingerprintSha256: hash((index + 1).toString(16)),
      uploadFingerprintSha256: hash((index + 11).toString(16)),
      captureSource:
        index % 2 === 0 ? "API_SHADOW_OBSERVATION" : "INTERNAL_PREVIEW",
      capturedAt: "2026-08-06T08:10:00.000Z",
      now: Date.parse("2026-08-06T08:30:00.000Z"),
    }).ledger;
  });
  return { ...support, ledger };
}

module.exports = Object.freeze({
  accuracyDataset,
  hash,
  createSourceFiles,
  completeSourceCatalog,
  completeLedger,
});
