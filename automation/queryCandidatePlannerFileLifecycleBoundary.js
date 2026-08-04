"use strict";

const {
  getQueryCandidatePlannerCacheRuntime,
} = require("./queryCandidatePlannerCacheRuntime");
const {
  deriveQueryCandidatePlannerUploadIdentity,
  publicUploadIdentity,
  invalidateQueryCandidatePlannerUploadCache,
} = require("./queryCandidatePlannerUploadLifecycle");

const BOUNDARY_VERSION =
  "query_candidate_planner_file_lifecycle_boundary_v1";

function text(value) {
  return String(value == null ? "" : value).trim();
}

function requestedOriginalName(request = {}) {
  const raw = text(request.file?.originalname);
  if (!raw) return "";
  try {
    return Buffer.from(raw, "latin1").toString("utf8");
  } catch (_error) {
    return raw;
  }
}

function existingUploadForRequest(request = {}, action = "") {
  const files = Array.isArray(request.user?.uploadedFiles)
    ? request.user.uploadedFiles
    : [];
  const targetName =
    action === "UPLOAD_REPLACEMENT"
      ? requestedOriginalName(request)
      : text(request.params?.originalName);
  if (!targetName) return null;
  return files.find(
    (file) => text(file?.originalName) === targetName,
  ) || null;
}

function lifecycleObservation({
  action,
  cacheDisposition,
  identity,
  invalidation,
  reason,
} = {}) {
  return Object.freeze({
    version: BOUNDARY_VERSION,
    action,
    cacheDisposition,
    reason: text(reason),
    identity: publicUploadIdentity(identity),
    invalidation: invalidation
      ? Object.freeze({
          version: text(invalidation.version),
          status: text(invalidation.status),
          reason: text(invalidation.reason),
          invalidated: invalidation.invalidated === true,
        })
      : null,
    privacy: Object.freeze({
      tenantIdIncluded: false,
      originalFileNameIncluded: false,
      queryTablesKeyIncluded: false,
      storageObjectKeyIncluded: false,
      cacheSecretIncluded: false,
    }),
  });
}

function setObservation(res, observation, task = null) {
  res.locals = res.locals || {};
  res.locals.queryCandidatePlannerCacheLifecycleObservation = observation;
  if (task) {
    res.locals.queryCandidatePlannerCacheLifecycleTask = task;
  }
}

function defaultLogger(observation = {}) {
  if ([
    "RETAINED",
    "NO_ACTIVE_CACHE",
    "NO_PREVIOUS_UPLOAD",
  ].includes(observation.cacheDisposition)) return;
  console.info("[query-candidate-cache-lifecycle]", {
    version: observation.version,
    action: observation.action,
    cacheDisposition: observation.cacheDisposition,
    reason: observation.reason,
    identityComplete: observation.identity?.complete === true,
    invalidationStatus: observation.invalidation?.status || "",
  });
}

function createQueryCandidatePlannerMutationBoundary({
  handler,
  action,
  runtimeProvider = getQueryCandidatePlannerCacheRuntime,
  invalidate = invalidateQueryCandidatePlannerUploadCache,
  onObservation = defaultLogger,
} = {}) {
  if (typeof handler !== "function") {
    throw new TypeError("file lifecycle boundary handler must be a function");
  }
  if (!["UPLOAD_REPLACEMENT", "DELETE"].includes(action)) {
    throw new TypeError("unsupported file lifecycle mutation action");
  }

  return async function queryCandidatePlannerMutationBoundary(req, res, next) {
    const existingFile = existingUploadForRequest(req, action);
    const identity = deriveQueryCandidatePlannerUploadIdentity({
      request: req,
      fileInfo: existingFile,
    });

    let invalidation = null;
    if (existingFile && identity.complete) {
      const task = Promise.resolve()
        .then(() =>
          invalidate({
            identity,
            runtime: runtimeProvider(),
            reason: "UPLOAD_DELETED",
          }),
        )
        .catch((error) =>
          Object.freeze({
            version: "query_candidate_planner_upload_invalidation_wiring_v1",
            status: "FAILED_SAFE",
            reason: text(error?.code || "UPLOAD_INVALIDATION_FAILED"),
            invalidated: false,
          }),
        );
      setObservation(
        res,
        lifecycleObservation({
          action,
          cacheDisposition: "INVALIDATION_PENDING",
          identity,
          reason: "PRE_MUTATION_INVALIDATION",
        }),
        task,
      );
      invalidation = await task;
    }

    let cacheDisposition = "NO_PREVIOUS_UPLOAD";
    let reason = "NO_PREVIOUS_UPLOAD_CACHE";
    if (existingFile && !identity.complete) {
      cacheDisposition = "IDENTITY_UNAVAILABLE";
      reason = text(identity.reason || "UPLOAD_IDENTITY_INCOMPLETE");
    } else if (existingFile && invalidation?.invalidated) {
      cacheDisposition = "INVALIDATED";
      reason = text(invalidation.reason || "UPLOAD_DELETED");
    } else if (existingFile && invalidation?.status === "SKIPPED") {
      cacheDisposition = "NO_ACTIVE_CACHE";
      reason = text(invalidation.reason || "CACHE_RUNTIME_UNAVAILABLE");
    } else if (existingFile) {
      cacheDisposition = "INVALIDATION_FAILED_SAFE";
      reason = text(
        invalidation?.reason || "UPLOAD_INVALIDATION_NOT_COMPLETED",
      );
    }

    const observation = lifecycleObservation({
      action,
      cacheDisposition,
      identity,
      invalidation,
      reason,
    });
    setObservation(res, observation);
    if (typeof onObservation === "function") {
      onObservation(observation, { req, res });
    }

    return handler(req, res, next);
  };
}

function createQueryCandidatePlannerDownloadRetentionBoundary({
  handler,
  action = "DOWNLOAD",
  onObservation = defaultLogger,
} = {}) {
  if (typeof handler !== "function") {
    throw new TypeError("download retention boundary handler must be a function");
  }
  return async function queryCandidatePlannerDownloadRetentionBoundary(
    req,
    res,
    next,
  ) {
    const identity = deriveQueryCandidatePlannerUploadIdentity({
      request: req,
    });
    const observation = lifecycleObservation({
      action,
      cacheDisposition: "RETAINED",
      identity,
      invalidation: null,
      reason: "DOWNLOAD_DOES_NOT_INVALIDATE_CACHE",
    });
    setObservation(res, observation);
    if (typeof onObservation === "function") {
      onObservation(observation, { req, res });
    }
    return handler(req, res, next);
  };
}

module.exports = Object.freeze({
  BOUNDARY_VERSION,
  requestedOriginalName,
  existingUploadForRequest,
  lifecycleObservation,
  defaultLogger,
  createQueryCandidatePlannerMutationBoundary,
  createQueryCandidatePlannerDownloadRetentionBoundary,
});
