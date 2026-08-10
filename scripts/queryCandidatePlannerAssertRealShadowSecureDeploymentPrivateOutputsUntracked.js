const childProcess = require("child_process");

const PRIVATE_PATTERNS = Object.freeze([
  /queryCandidatePlannerRealShadow.*\.private\.(json|txt|env)$/i,
  /queryCandidatePlannerRealShadowEvidenceSecret.*\.private\.env$/i,
  /BeeBeeAI-Private-Evidence/i,
]);

function isForbiddenPrivatePath(filePath) {
  const value = String(filePath == null ? "" : filePath).replace(/\\/g, "/");
  return PRIVATE_PATTERNS.some((pattern) => pattern.test(value));
}

function stagedPrivatePaths(cwd = process.cwd()) {
  const output = childProcess.execFileSync(
    "git",
    ["diff", "--cached", "--name-only", "--diff-filter=ACMR"],
    { cwd, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] },
  );
  return output
    .split(/\r?\n/)
    .map((item) => item.trim())
    .filter(Boolean)
    .filter(isForbiddenPrivatePath);
}

function main() {
  try {
    const staged = stagedPrivatePaths();
    if (staged.length > 0) {
      staged.forEach((item) => console.error(`BLOCKED_PRIVATE_STAGED ${item}`));
      process.exitCode = 2;
      return;
    }
    console.log("PASS no secure-deployment private outputs staged");
  } catch (error) {
    console.error(`FAIL ${error.code || error.message}`);
    process.exitCode = 1;
  }
}

if (require.main === module) main();

module.exports = Object.freeze({
  PRIVATE_PATTERNS,
  isForbiddenPrivatePath,
  stagedPrivatePaths,
});
