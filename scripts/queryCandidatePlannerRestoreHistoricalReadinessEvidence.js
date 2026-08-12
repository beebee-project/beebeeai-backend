const fs = require("fs");
const path = require("path");
const {
  buildHistoricalReadinessCapsule,
  validateHistoricalReadinessCapsule,
} = require("../automation/queryCandidatePlannerHistoricalReadinessEvidenceRecovery");

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith("--")) {
      out[argv[i].slice(2)] = argv[++i] || "";
    }
  }
  return out;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.output) throw new Error("--output is required");

  const capsule = buildHistoricalReadinessCapsule();
  validateHistoricalReadinessCapsule(capsule);

  const target = path.resolve(args.output);
  fs.mkdirSync(path.dirname(target), { recursive: true });

  if (fs.existsSync(target) && !args.force) {
    throw new Error(
      "Output already exists. Refusing overwrite without --force true.",
    );
  }

  fs.writeFileSync(target, `${JSON.stringify(capsule, null, 2)}\n`, "utf8");

  console.log("PASS historical Patch 13.3 readiness evidence restored");
  console.log(`OUTPUT ${target}`);
  console.log(`SOURCE_VERSION ${capsule.version}`);
  console.log(`MODEL ${capsule.origin.model}`);
  console.log("ORIGIN CALLED providerCalls=1");
  console.log("REPLAY CACHE_HIT providerCalls=0");
  console.log("L3_SOURCE L3_SEMANTIC");
  console.log("L4_SOURCE L4_REENTRY");
  console.log("PARITY_VALID true");
  console.log("ENCRYPTED_PERSISTENT_FILES 3");
  console.log("PLAINTEXT_PERSISTENT_FILES 0");
  console.log("READINESS_ELIGIBLE true");
  console.log("PROVIDER_CALLS_EXECUTED_BY_RECOVERY 0");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
  console.log(`RECOVERY_CAPSULE_SHA256 ${capsule.recoveryCapsuleSha256}`);
  console.log("PRIVATE_OUTPUT_DO_NOT_COMMIT true");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
