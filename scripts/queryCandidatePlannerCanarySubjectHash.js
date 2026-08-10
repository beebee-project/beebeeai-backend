const {
  hashQueryCandidatePlannerInternalCanarySubject,
} = require("../automation/queryCandidatePlannerInternalCanarySubject");

const accountId = String(process.argv[2] || "").trim();
const tenantId = String(process.argv[3] || "").trim();

if (!accountId) {
  console.error(
    "Usage: node scripts/queryCandidatePlannerCanarySubjectHash.js <immutableAccountId> [tenantId]",
  );
  process.exitCode = 1;
} else {
  process.stdout.write(
    `${hashQueryCandidatePlannerInternalCanarySubject({ accountId, tenantId })}\n`,
  );
}
