const fs = require("fs");
const path = require("path");
const mongoose = require("mongoose");

try {
  require("dotenv").config({
    path: path.resolve(".env"),
  });
} catch {}

const User = require("../models/User");

const {
  deriveQueryCandidatePlannerInternalCanarySubject,
} = require("../automation/queryCandidatePlannerInternalCanarySubject");

function text(value) {
  return String(value == null ? "" : value).trim();
}

function first(...values) {
  for (const value of values) {
    const v = text(value);
    if (v) return v;
  }
  return "";
}

async function main() {
  const allowlistSha = text(
    process.env.QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256,
  ).toUpperCase();

  const output = text(process.env.BEEBEE_E2E_CANARY_MATCH_FILE);

  if (!/^[A-F0-9]{64}$/.test(allowlistSha)) {
    throw new Error("ALLOWLIST_SHA_INVALID");
  }

  if (!output) {
    throw new Error("MATCH_OUTPUT_PATH_REQUIRED");
  }

  const uri = text(process.env.MONGO_URI);

  if (!uri) {
    throw new Error("MONGO_URI_MISSING");
  }

  await mongoose.connect(uri, {
    serverSelectionTimeoutMS: 10000,
  });

  const users = await User.find({})
    .select("_id accountId userId tenantId organizationId orgId")
    .lean();

  const matches = [];

  for (const user of users) {
    const accountId = first(user.accountId, user.userId, user._id);

    const tenantId = first(user.tenantId, user.organizationId, user.orgId);

    if (!accountId) continue;

    const request = {
      user: {
        accountId,
        tenantId,
      },
    };

    const subject = deriveQueryCandidatePlannerInternalCanarySubject(request);

    if (text(subject.subjectSha256).toUpperCase() === allowlistSha) {
      matches.push({
        accountId,
        tenantId,
        source: subject.source,
      });
    }
  }

  console.log(`USER_COUNT_SCANNED ${users.length}`);
  console.log(`MATCH_COUNT ${matches.length}`);

  if (matches.length === 0) {
    console.log("MATCH_FOUND false");
    return;
  }

  if (matches.length !== 1) {
    throw new Error(`AMBIGUOUS_ALLOWLIST_SUBJECT_MATCH_${matches.length}`);
  }

  fs.writeFileSync(output, JSON.stringify(matches[0]), {
    encoding: "utf8",
    mode: 0o600,
  });

  console.log("MATCH_FOUND true");
  console.log(`MATCH_SOURCE ${matches[0].source}`);
  console.log(`MATCH_HAS_TENANT ${Boolean(matches[0].tenantId)}`);
  console.log("RAW_ACCOUNT_ID_LOGGED false");
  console.log("RAW_TENANT_ID_LOGGED false");
}

main()
  .catch((error) => {
    console.error(`BLOCKED ${error.code || error.message}`);
    process.exitCode = 1;
  })
  .finally(async () => {
    await mongoose.disconnect().catch(() => {});
  });
