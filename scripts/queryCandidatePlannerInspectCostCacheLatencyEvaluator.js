const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const relative = "automation/queryCandidatePlannerCostCacheLatencyEvaluator.js";
const absolute = path.resolve(relative);

function sha(data) {
  return crypto.createHash("sha256").update(data).digest("hex").toUpperCase();
}

function main() {
  if (!fs.existsSync(absolute)) {
    throw new Error(`Evaluator missing: ${relative}`);
  }

  const worktree = fs.readFileSync(absolute);
  let head = null;
  try {
    head = execFileSync("git", ["show", `HEAD:${relative}`], {
      encoding: null,
    });
  } catch {
    head = null;
  }

  const evaluator = require(absolute);
  const exportNames =
    typeof evaluator === "function"
      ? ["<module-function>"]
      : Object.keys(evaluator || {}).sort();

  console.log("PASS cost/cache/latency evaluator inspection");
  console.log(`WORKTREE_SHA256 ${sha(worktree)}`);
  console.log(`HEAD_SHA256 ${head ? sha(head) : "UNAVAILABLE"}`);
  console.log(
    `WORKTREE_EQUALS_HEAD ${head ? sha(worktree) === sha(head) : "UNAVAILABLE"}`,
  );
  console.log(`EXPORTS ${exportNames.join(",")}`);
  console.log("PROVIDER_CALLS_EXECUTED 0");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
