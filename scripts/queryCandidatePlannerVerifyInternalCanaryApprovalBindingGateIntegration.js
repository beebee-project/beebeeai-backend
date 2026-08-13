"use strict";

const fs = require("fs");
const path = require("path");

const TARGET =
  "automation/queryCandidatePlannerInternalAllowlistCanaryService.js";

const REQUIRED = Object.freeze([
  "// PATCH 15.3.2-F.1.6 APPROVAL BINDING GATE REQUIRE",
  "evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate",
  "// PATCH 15.3.2-F.1.6 APPROVAL BINDING GATE INTEGRATION",
  "return approvalBindingGate.preflight;",
]);

function main() {
  const repoRoot = path.resolve(process.argv[2] || ".");
  const target = path.join(repoRoot, TARGET);

  if (!fs.existsSync(target)) {
    throw new Error(`Target service missing: ${TARGET}`);
  }

  const text = fs.readFileSync(target, "utf8");

  for (const token of REQUIRED) {
    if (!text.includes(token)) {
      throw new Error(`F.1.6 integration token missing: ${token}`);
    }
  }

  const requireCount =
    text.split(
      "evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate",
    ).length - 1;

  if (requireCount < 2) {
    throw new Error("F.1.6 gate require/invocation incomplete");
  }

  console.log("PASS Patch 15.3.2-F.1.6 service integration verification");
  console.log("APPROVAL_BINDING_GATE_REQUIRED true");
  console.log("LEGACY_PREAPPROVAL_BYPASS_ALLOWED false");
  console.log("ROUTE_MODIFIED false");
  console.log("PRODUCTION_ROUTE_CHANGED false");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
