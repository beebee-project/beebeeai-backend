"use strict";

const crypto = require("crypto");
const path = require("path");

const RUNTIME_VERSION =
  "query_candidate_planner_cache_runtime_v1";

let runtimeOverride = null;
let runtimeInstance = null;

function text(value) {
  return String(value == null ? "" : value).trim();
}

function decodeKey(value) {
  const raw = text(value);
  if (!raw) return null;
  if (/^[0-9a-f]{64}$/i.test(raw)) {
    return Buffer.from(raw, "hex");
  }
  try {
    const decoded = Buffer.from(raw, "base64");
    if (decoded.length === 32) return decoded;
  } catch (_error) {
    // Fall through to deterministic derivation.
  }
  return crypto.createHash("sha256").update(raw, "utf8").digest();
}

function disabledRuntime(reason) {
  return Object.freeze({
    version: RUNTIME_VERSION,
    enabled: false,
    reason,
    hierarchicalCache: null,
    cacheSecret: "",
    privacy: Object.freeze({
      cacheSecretIncluded: false,
      encryptionKeyIncluded: false,
      tenantIdIncluded: false,
    }),
  });
}

function resolveCacheFactory(moduleValue = {}) {
  if (typeof moduleValue.createEncryptedHierarchicalCandidatePlannerCache === "function") {
    return moduleValue.createEncryptedHierarchicalCandidatePlannerCache;
  }
  if (typeof moduleValue.createQueryCandidatePlannerHierarchicalEncryptedCache === "function") {
    return moduleValue.createQueryCandidatePlannerHierarchicalEncryptedCache;
  }
  if (typeof moduleValue.default === "function") return moduleValue.default;
  if (typeof moduleValue === "function") return moduleValue;
  return null;
}

function buildRuntimeFromEnvironment(env = process.env) {
  const rawKey = text(env.QUERY_CANDIDATE_PLANNER_CACHE_KEY);
  const rawSecret = text(
    env.QUERY_CANDIDATE_PLANNER_CACHE_SECRET || rawKey,
  );
  if (!rawKey || !rawSecret) {
    return disabledRuntime("CACHE_RUNTIME_ENV_NOT_CONFIGURED");
  }

  const key = decodeKey(rawKey);
  if (!key || key.length !== 32) {
    return disabledRuntime("CACHE_RUNTIME_KEY_INVALID");
  }

  let cacheModule;
  try {
    cacheModule = require("./queryCandidatePlannerHierarchicalEncryptedCache");
  } catch (_error) {
    return disabledRuntime("CACHE_RUNTIME_MODULE_UNAVAILABLE");
  }

  const factory = resolveCacheFactory(cacheModule);
  if (!factory) {
    return disabledRuntime("CACHE_RUNTIME_FACTORY_MISSING");
  }

  const rootDir = path.resolve(
    text(env.QUERY_CANDIDATE_PLANNER_CACHE_ROOT) ||
      path.join(process.cwd(), ".local_uploads", "query-candidate-planner-cache"),
  );
  const keyId =
    text(env.QUERY_CANDIDATE_PLANNER_CACHE_KEY_ID) ||
    "candidate-planner-cache-primary";

  let codecOptions = {
    key,
    keyId,
    primaryKey: key,
    primaryKeyId: keyId,
  };
  if (typeof cacheModule.createRotatingAes256GcmCacheCodec === "function") {
    codecOptions = cacheModule.createRotatingAes256GcmCacheCodec({
      primary: { key, keyId },
      legacy: [],
    });
  }

  try {
    const hierarchicalCache = factory({
      rootDir,
      ...codecOptions,
    });
    if (!hierarchicalCache || typeof hierarchicalCache !== "object") {
      return disabledRuntime("CACHE_RUNTIME_FACTORY_INVALID_RESULT");
    }
    return Object.freeze({
      version: RUNTIME_VERSION,
      enabled: true,
      reason: "CACHE_RUNTIME_READY",
      hierarchicalCache,
      cacheSecret: rawSecret,
      privacy: Object.freeze({
        cacheSecretIncluded: false,
        encryptionKeyIncluded: false,
        tenantIdIncluded: false,
      }),
    });
  } catch (_error) {
    return disabledRuntime("CACHE_RUNTIME_INITIALIZATION_FAILED");
  }
}

function getQueryCandidatePlannerCacheRuntime() {
  if (runtimeOverride) return runtimeOverride;
  if (!runtimeInstance) {
    runtimeInstance = buildRuntimeFromEnvironment(process.env);
  }
  return runtimeInstance;
}

function getQueryCandidatePlannerCacheRuntimeSnapshot() {
  const runtime = getQueryCandidatePlannerCacheRuntime();
  return Object.freeze({
    version: RUNTIME_VERSION,
    enabled: runtime.enabled === true,
    reason: text(runtime.reason),
    cachePresent: Boolean(runtime.hierarchicalCache),
    cacheSecretPresent: Boolean(runtime.cacheSecret),
    privacy: Object.freeze({
      cacheSecretIncluded: false,
      encryptionKeyIncluded: false,
      tenantIdIncluded: false,
    }),
  });
}

function resetQueryCandidatePlannerCacheRuntimeForTests({
  runtime = null,
} = {}) {
  runtimeOverride = runtime;
  runtimeInstance = null;
  return runtimeOverride;
}

module.exports = Object.freeze({
  RUNTIME_VERSION,
  decodeKey,
  buildRuntimeFromEnvironment,
  getQueryCandidatePlannerCacheRuntime,
  getQueryCandidatePlannerCacheRuntimeSnapshot,
  resetQueryCandidatePlannerCacheRuntimeForTests,
});
