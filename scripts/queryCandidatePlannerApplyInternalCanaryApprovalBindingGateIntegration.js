"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const TARGET =
  "automation/queryCandidatePlannerInternalAllowlistCanaryService.js";

const REQUIRE_MARKER =
  "// PATCH 15.3.2-F.1.6 APPROVAL BINDING GATE REQUIRE";

const INTEGRATION_MARKER =
  "// PATCH 15.3.2-F.1.6 APPROVAL BINDING GATE INTEGRATION";

function sha256(data) {
  return crypto.createHash("sha256").update(data).digest("hex").toUpperCase();
}

function fail(message) {
  throw new Error(message);
}

function main() {
  const repoRoot = path.resolve(process.argv[2] || ".");
  const target = path.join(repoRoot, TARGET);

  if (!fs.existsSync(target)) {
    fail(`Target service missing: ${TARGET}`);
  }

  let source = fs.readFileSync(target, "utf8");
  const beforeSha = sha256(Buffer.from(source, "utf8"));

  if (
    source.includes(REQUIRE_MARKER) &&
    source.includes(INTEGRATION_MARKER)
  ) {
    console.log("PASS Patch 15.3.2-F.1.6 integration already applied");
    console.log(`SERVICE_SHA256 ${beforeSha}`);
    console.log("IDEMPOTENT true");
    console.log("ROUTE_MODIFIED false");
    console.log("PROVIDER_CALLS_EXECUTED_BY_PATCH 0");
    return;
  }

  const subjectAnchor =
    "  const subject = deriveQueryCandidatePlannerInternalCanarySubject(request);";

  const subjectCount = source.split(subjectAnchor).length - 1;
  if (subjectCount !== 1) {
    fail(
      `Integration anchor count invalid for subject derivation: ${subjectCount}`,
    );
  }

  const requireBlock = [
    REQUIRE_MARKER,
    'const {',
    '  evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate,',
    '} = require("./queryCandidatePlannerInternalCanaryApprovalBindingGate");',
    "",
  ].join("\n");

  if (source.startsWith('"use strict";')) {
    const strictLine = '"use strict";';
    source =
      strictLine +
      "\n\n" +
      requireBlock +
      source.slice(strictLine.length).replace(/^\s*/, "");
  } else {
    source = requireBlock + source;
  }

  const integrationBlock = [
    subjectAnchor,
    "",
    `  ${INTEGRATION_MARKER}`,
    "  const approvalBindingGate =",
    "    evaluateQueryCandidatePlannerInternalCanaryApprovalBindingGate({",
    "      env,",
    "      featureControl,",
    "      subject,",
    "    });",
    "",
    "  // F.1.6 is mandatory after installation. The prior evidence path remains",
    "  // in source for rollback/reference but is intentionally superseded here.",
    "  return approvalBindingGate.preflight;",
  ].join("\n");

  source = source.replace(subjectAnchor, integrationBlock);

  const backupRoot = path.join(repoRoot, ".patch_backups");
  fs.mkdirSync(backupRoot, { recursive: true });

  const timestamp = new Date()
    .toISOString()
    .replace(/[:.]/g, "-");

  const backupDir = path.join(
    backupRoot,
    `query_candidate_patch15_3_2_F_1_6_${timestamp}`,
  );
  fs.mkdirSync(backupDir, { recursive: true });

  const backupFile = path.join(
    backupDir,
    path.basename(target),
  );
  fs.copyFileSync(target, backupFile);

  fs.writeFileSync(target, source, "utf8");

  const after = fs.readFileSync(target);
  const afterSha = sha256(after);

  if (!source.includes(REQUIRE_MARKER)) {
    fail("Require marker missing after integration");
  }
  if (!source.includes(INTEGRATION_MARKER)) {
    fail("Integration marker missing after integration");
  }

  console.log("PASS Patch 15.3.2-F.1.6 service integration applied");
  console.log(`TARGET ${TARGET}`);
  console.log(`BEFORE_SHA256 ${beforeSha}`);
  console.log(`AFTER_SHA256 ${afterSha}`);
  console.log(`BACKUP ${path.relative(repoRoot, backupFile)}`);
  console.log("ROUTE_MODIFIED false");
  console.log("ENVIRONMENT_MODIFIED false");
  console.log("PROVIDER_CALLS_EXECUTED_BY_PATCH 0");
  console.log("PRODUCTION_PROMOTION_AUTHORIZED false");
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`BLOCKED ${error.message}`);
    process.exitCode = 1;
  }
}
