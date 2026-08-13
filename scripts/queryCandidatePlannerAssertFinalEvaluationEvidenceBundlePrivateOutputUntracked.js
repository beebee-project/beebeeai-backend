"use strict";
const { execFileSync } = require("child_process");
const path = require("path");

const target = path
  .relative(process.cwd(), path.resolve(
    "queryCandidatePlannerPatch15_3_2_G.private/queryCandidatePlannerFinalEvaluationEvidenceBundle.private.json",
  ))
  .replace(/\\/g, "/");

function run(args) {
  try {
    return execFileSync("git", args, { encoding: "utf8" }).trim();
  } catch (error) {
    return String(error.stdout || "").trim();
  }
}

const staged = run(["diff", "--cached", "--name-only", "--", target]);
const tracked = run(["ls-files", "--", target]);
if (staged || tracked) {
  console.error("BLOCKED G_PRIVATE_EVIDENCE_OUTPUT_TRACKED_OR_STAGED");
  process.exitCode = 1;
} else {
  console.log("PASS Patch 15.3.2-G private evidence bundle is not tracked or staged");
}
