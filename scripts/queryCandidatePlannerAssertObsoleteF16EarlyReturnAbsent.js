const fs = require("fs");

const target =
  "automation/queryCandidatePlannerInternalAllowlistCanaryService.js";

const obsoleteMarker =
  "// PATCH 15.3.2-F.1.6 APPROVAL BINDING GATE INTEGRATION";

const obsoleteComment = "prior evidence path remains";

function main() {
  if (!fs.existsSync(target)) {
    throw new Error(`Target missing: ${target}`);
  }

  const source = fs.readFileSync(target, "utf8");

  if (source.includes(obsoleteMarker) || source.includes(obsoleteComment)) {
    throw new Error("Obsolete F.1.6 early-return integration detected");
  }

  console.log(
    "PASS obsolete Patch 15.3.2-F.1.6 early-return integration absent",
  );
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
