"use strict";

const crypto = require("crypto");

const LIFECYCLE_VERSION =
  "query_candidate_planner_upload_lifecycle_v1";
const IDENTITY_VERSION =
  "query_candidate_planner_upload_identity_v1";
const INVALIDATION_VERSION =
  "query_candidate_planner_upload_invalidation_wiring_v1";

function text(value) {
  return String(value == null ? "" : value).trim();
}

function sha256(value) {
  const serialized =
    typeof value === "string" ? value : JSON.stringify(value);
  return crypto.createHash("sha256").update(serialized).digest("hex");
}

function userIdFromRequest(request = {}) {
  return text(request.user?.id || request.user?._id);
}

function toPlainFileInfo(fileInfo = null) {
  if (!fileInfo) return null;
  if (typeof fileInfo.toObject === "function") {
    return fileInfo.toObject();
  }
  return fileInfo;
}

function uploadedFiles(request = {}) {
  const files = request.user?.uploadedFiles;
  return Array.isArray(files) ? files.map(toPlainFileInfo) : [];
}

function decodeOriginalName(rawName = "") {
  const raw = text(rawName);
  if (!raw) return "";
  try {
    return Buffer.from(raw, "latin1").toString("utf8");
  } catch (_error) {
    return raw;
  }
}

function requestFileName(request = {}, primaryPayload = {}) {
  return text(
    request.params?.originalName ||
      request.body?.fileName ||
      primaryPayload.fileName ||
      decodeOriginalName(request.file?.originalname),
  );
}

function findUploadFileInfo({
  request = {},
  primaryPayload = {},
  explicitFileInfo = null,
} = {}) {
  if (explicitFileInfo) return toPlainFileInfo(explicitFileInfo);
  const files = uploadedFiles(request);
  const queryTablesKey = text(
    request.body?.queryTablesKey || primaryPayload.queryTablesKey,
  );
  if (queryTablesKey) {
    const byQueryKey = files.find(
      (file) => text(file?.queryJsonKey) === queryTablesKey,
    );
    if (byQueryKey) return byQueryKey;
  }
  const fileName = requestFileName(request, primaryPayload);
  if (fileName) {
    const byName = files.find(
      (file) => text(file?.originalName) === fileName,
    );
    if (byName) return byName;
  }
  return null;
}

function deriveQueryObjectIdentity({
  request = {},
  primaryPayload = {},
  fileInfo = null,
} = {}) {
  const resolvedFile = findUploadFileInfo({
    request,
    primaryPayload,
    explicitFileInfo: fileInfo,
  });
  return text(
    request.body?.queryTablesKey ||
      primaryPayload.queryTablesKey ||
      resolvedFile?.queryJsonKey,
  );
}

function deriveStorageObjectIdentity(fileInfo = null) {
  const resolved = toPlainFileInfo(fileInfo) || {};
  return text(
    resolved.localName ||
      resolved.gcsName ||
      resolved.storageKey ||
      resolved.queryJsonKey,
  );
}

function deriveQueryCandidatePlannerUploadIdentity({
  request = {},
  primaryPayload = {},
  fileInfo = null,
} = {}) {
  const tenantId = userIdFromRequest(request);
  if (!tenantId) {
    return Object.freeze({
      version: IDENTITY_VERSION,
      complete: false,
      reason: "TENANT_ID_UNAVAILABLE",
      tenantId: "",
      uploadFingerprintSha256: "",
      queryJsonSha256: "",
      source: "NONE",
      privacy: Object.freeze({
        originalFileNameIncluded: false,
        queryTablesKeyIncluded: false,
        storageObjectKeyIncluded: false,
      }),
    });
  }

  const resolvedFile = findUploadFileInfo({
    request,
    primaryPayload,
    explicitFileInfo: fileInfo,
  });
  const queryObjectIdentity = deriveQueryObjectIdentity({
    request,
    primaryPayload,
    fileInfo: resolvedFile,
  });
  const storageObjectIdentity = deriveStorageObjectIdentity(resolvedFile);
  const fileHash = text(primaryPayload.fileHash || resolvedFile?.fileHash);
  const sheetStateSig = text(
    primaryPayload.sheetStateSig || resolvedFile?.sheetStateSig,
  );
  const stableUploadObject =
    queryObjectIdentity || storageObjectIdentity || `${fileHash}:${sheetStateSig}`;

  if (!stableUploadObject) {
    return Object.freeze({
      version: IDENTITY_VERSION,
      complete: false,
      reason: "UPLOAD_OBJECT_IDENTITY_UNAVAILABLE",
      tenantId,
      uploadFingerprintSha256: "",
      queryJsonSha256: "",
      source: "NONE",
      privacy: Object.freeze({
        originalFileNameIncluded: false,
        queryTablesKeyIncluded: false,
        storageObjectKeyIncluded: false,
      }),
    });
  }

  const uploadFingerprintSha256 = sha256({
    version: IDENTITY_VERSION,
    kind: "UPLOAD_OBJECT",
    tenantId,
    stableUploadObject,
  });
  const queryJsonSha256 = sha256({
    version: IDENTITY_VERSION,
    kind: "QUERY_JSON_OBJECT",
    tenantId,
    queryObjectIdentity: queryObjectIdentity || stableUploadObject,
  });

  return Object.freeze({
    version: IDENTITY_VERSION,
    complete: true,
    reason: "UPLOAD_IDENTITY_READY",
    tenantId,
    uploadFingerprintSha256,
    queryJsonSha256,
    source: queryObjectIdentity
      ? "QUERY_JSON_OBJECT"
      : storageObjectIdentity
        ? "STORAGE_OBJECT"
        : "CONTENT_FALLBACK",
    privacy: Object.freeze({
      originalFileNameIncluded: false,
      queryTablesKeyIncluded: false,
      storageObjectKeyIncluded: false,
    }),
  });
}

function publicUploadIdentity(identity = {}) {
  return Object.freeze({
    version: text(identity.version || IDENTITY_VERSION),
    complete: identity.complete === true,
    reason: text(identity.reason),
    source: text(identity.source),
    uploadFingerprintSha256: text(identity.uploadFingerprintSha256),
    queryJsonSha256: text(identity.queryJsonSha256),
    privacy: Object.freeze({
      tenantIdIncluded: false,
      originalFileNameIncluded: false,
      queryTablesKeyIncluded: false,
      storageObjectKeyIncluded: false,
    }),
  });
}

function resolveInvalidationFunction(moduleValue = {}) {
  if (typeof moduleValue.invalidateCandidatePlannerUploadCache === "function") {
    return moduleValue.invalidateCandidatePlannerUploadCache;
  }
  if (typeof moduleValue.invalidateQueryCandidatePlannerUploadCache === "function") {
    return moduleValue.invalidateQueryCandidatePlannerUploadCache;
  }
  return null;
}

async function invalidateQueryCandidatePlannerUploadCache({
  identity = {},
  runtime = null,
  reason = "UPLOAD_DELETED",
  invalidator = null,
} = {}) {
  if (!identity.complete) {
    return Object.freeze({
      version: INVALIDATION_VERSION,
      status: "SKIPPED",
      reason: text(identity.reason || "UPLOAD_IDENTITY_INCOMPLETE"),
      invalidated: false,
      privacy: Object.freeze({ tenantIdIncluded: false }),
    });
  }
  if (!runtime?.enabled || !runtime.hierarchicalCache || !runtime.cacheSecret) {
    return Object.freeze({
      version: INVALIDATION_VERSION,
      status: "SKIPPED",
      reason: text(runtime?.reason || "CACHE_RUNTIME_UNAVAILABLE"),
      invalidated: false,
      privacy: Object.freeze({ tenantIdIncluded: false }),
    });
  }

  let invalidate = invalidator;
  if (typeof invalidate !== "function") {
    let controls;
    try {
      controls = require("./queryCandidatePlannerCacheOperationalControls");
    } catch (_error) {
      controls = {};
    }
    invalidate = resolveInvalidationFunction(controls);
  }
  if (typeof invalidate !== "function") {
    return Object.freeze({
      version: INVALIDATION_VERSION,
      status: "FAILED_SAFE",
      reason: "UPLOAD_INVALIDATOR_UNAVAILABLE",
      invalidated: false,
      privacy: Object.freeze({ tenantIdIncluded: false }),
    });
  }

  const result = await invalidate({
    hierarchicalCache: runtime.hierarchicalCache,
    tenantId: identity.tenantId,
    cacheSecret: runtime.cacheSecret,
    uploadFingerprintSha256: identity.uploadFingerprintSha256,
    queryJsonSha256: identity.queryJsonSha256,
    reason,
  });
  return Object.freeze({
    version: INVALIDATION_VERSION,
    status: "INVALIDATED",
    reason,
    invalidated: true,
    result: result && typeof result === "object"
      ? Object.freeze({
          version: text(result.version),
          status: text(result.status),
          invalidatedCount: Number(result.invalidatedCount || result.count || 0),
        })
      : null,
    identity: publicUploadIdentity(identity),
    privacy: Object.freeze({
      tenantIdIncluded: false,
      cacheSecretIncluded: false,
      rawInvalidationResultIncluded: false,
    }),
  });
}

module.exports = Object.freeze({
  LIFECYCLE_VERSION,
  IDENTITY_VERSION,
  INVALIDATION_VERSION,
  sha256,
  userIdFromRequest,
  decodeOriginalName,
  findUploadFileInfo,
  deriveQueryCandidatePlannerUploadIdentity,
  publicUploadIdentity,
  invalidateQueryCandidatePlannerUploadCache,
});
