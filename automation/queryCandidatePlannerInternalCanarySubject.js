const crypto = require("crypto");

const SUBJECT_VERSION = "query_candidate_planner_internal_canary_subject_v1";
const SUBJECT_HASH_PREFIX = "beebee-query-candidate-internal-canary-subject-v1";

function text(value, maxLength = 240) {
  return String(value == null ? "" : value)
    .trim()
    .replace(/[\r\n\t]/g, " ")
    .slice(0, maxLength);
}

function sha256(value) {
  return crypto.createHash("sha256").update(String(value)).digest("hex");
}

function firstText(...values) {
  for (const value of values) {
    const normalized = text(value);
    if (normalized) return normalized;
  }
  return "";
}

function normalizeIdentityPart(value) {
  return text(value, 240).toLowerCase();
}

function deriveQueryCandidatePlannerInternalCanarySubject(request = {}) {
  const user =
    request.user && typeof request.user === "object" ? request.user : {};
  const accountId = firstText(
    user.accountId,
    user.userId,
    user._id,
    user.id,
    user.sub,
  );
  const tenantId = firstText(
    user.tenantId,
    user.organizationId,
    user.orgId,
    request.tenantId,
  );

  if (!accountId) {
    return Object.freeze({
      version: SUBJECT_VERSION,
      complete: false,
      reason: "IMMUTABLE_ACCOUNT_ID_REQUIRED",
      subjectSha256: "",
      subjectTagSha256: "",
      source: "NONE",
      privacy: Object.freeze({
        rawAccountIdIncluded: false,
        rawTenantIdIncluded: false,
        emailIncluded: false,
        nameIncluded: false,
      }),
    });
  }

  const canonical = [
    SUBJECT_HASH_PREFIX,
    normalizeIdentityPart(tenantId) || "no-tenant",
    normalizeIdentityPart(accountId),
  ].join(":");
  const subjectSha256 = sha256(canonical);

  return Object.freeze({
    version: SUBJECT_VERSION,
    complete: true,
    reason: "IMMUTABLE_ACCOUNT_SUBJECT_DERIVED",
    subjectSha256,
    subjectTagSha256: sha256(`safe-tag:${subjectSha256}`),
    source: tenantId ? "TENANT_AND_ACCOUNT_ID" : "ACCOUNT_ID",
    privacy: Object.freeze({
      rawAccountIdIncluded: false,
      rawTenantIdIncluded: false,
      emailIncluded: false,
      nameIncluded: false,
    }),
  });
}

function hashQueryCandidatePlannerInternalCanarySubject({
  accountId,
  tenantId = "",
} = {}) {
  return deriveQueryCandidatePlannerInternalCanarySubject({
    user: { accountId, tenantId },
  }).subjectSha256;
}

module.exports = Object.freeze({
  SUBJECT_VERSION,
  SUBJECT_HASH_PREFIX,
  deriveQueryCandidatePlannerInternalCanarySubject,
  hashQueryCandidatePlannerInternalCanarySubject,
});
