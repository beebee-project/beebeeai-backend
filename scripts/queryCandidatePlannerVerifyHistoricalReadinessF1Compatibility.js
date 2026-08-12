const fs = require("fs");
const path = require("path");
const {
  validateHistoricalReadinessCapsule,
} = require("../automation/queryCandidatePlannerHistoricalReadinessEvidenceRecovery");

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) out[argv[i].slice(2)] = argv[++i] || "";
  }
  return out;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.readiness) throw new Error("--readiness is required");

  const readiness = JSON.parse(
    fs.readFileSync(path.resolve(args.readiness), "utf8"),
  );

  validateHistoricalReadinessCapsule(readiness);

  let f1Validator = null;
  try {
    f1Validator = require(
      path.resolve(
        "automation/queryCandidatePlannerCanonicalEvaluationInput.js",
      ),
    );
  } catch {
    throw new Error("Patch 15.3.2-F.1 canonical evaluation module is missing.");
  }

  if (typeof f1Validator.validateLiveParityReadiness !== "function") {
    throw new Error(
      "Patch 15.3.2-F.1 validateLiveParityReadiness export is missing.",
    );
  }

  const facts = f1Validator.validateLiveParityReadiness(readiness);

  console.log("PASS historical readiness is Patch 15.3.2-F.1 compatible");
  console.log(`ORIGIN_PROVIDER_CALLS ${facts.originProviderCalls}`);
  console.log(`REPLAY_PROVIDER_CALLS ${facts.replayProviderCalls}`);
  console.log(`PLANNER_CACHE_SOURCE ${facts.plannerCacheSource}`);
  console.log(`REENTRY_CACHE_SOURCE ${facts.reentryCacheSource}`);
  console.log(`PARITY_VALID ${facts.parityValid}`);
  console.log(
    `ENCRYPTED_PERSISTENT_FILE_COUNT ${facts.encryptedPersistentFileCount}`,
  );
  console.log(
    `PLAINTEXT_PERSISTENT_FILE_COUNT ${facts.plaintextPersistentFileCount}`,
  );
  console.log(`READINESS_ELIGIBLE ${facts.readinessEligible}`);
  console.log("PROVIDER_CALLS_EXECUTED_BY_VERIFICATION 0");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.code || error.message}`);
    process.exitCode = 1;
  }
}
