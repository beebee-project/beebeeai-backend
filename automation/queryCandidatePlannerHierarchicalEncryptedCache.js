"use strict";

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const {
  normalizeText,
  sha256,
  stableStringify,
} = require("./queryCandidateObservation");

const QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_VERSION =
  "query_candidate_planner_hierarchical_cache_v1";
const QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_KEY_VERSION =
  "query_candidate_planner_hierarchical_cache_key_v1";
const QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_ENTRY_VERSION =
  "query_candidate_planner_hierarchical_cache_entry_v1";
const QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_POLICY_VERSION =
  "encrypted_hierarchical_candidate_planner_cache_policy_v1";
const QUERY_CANDIDATE_PLANNER_CACHE_OPERATIONAL_CONTROL_VERSION =
  "query_candidate_planner_cache_operational_control_v1";
const QUERY_CANDIDATE_PLANNER_CACHE_AUDIT_EVENT_VERSION =
  "query_candidate_planner_cache_audit_event_v1";

const CACHE_INVALIDATION_REASONS = Object.freeze({
  TTL_EXPIRED: "TTL_EXPIRED",
  TENANT_DELETED: "TENANT_DELETED",
  UPLOAD_DELETED: "UPLOAD_DELETED",
  LAYER_INVALIDATED: "LAYER_INVALIDATED",
  MANUAL_DELETE: "MANUAL_DELETE",
  CORRUPT_ENTRY: "CORRUPT_ENTRY",
  KEY_ROTATED: "KEY_ROTATED",
});

const CACHE_LAYERS = Object.freeze({
  L2_UPLOAD: "L2_UPLOAD",
  L3_SEMANTIC: "L3_SEMANTIC",
  L4_REENTRY: "L4_REENTRY",
});

const CACHE_ARTIFACT_TYPES = Object.freeze({
  UPLOAD_QUERY: "UPLOAD_QUERY",
  PLANNER_PROVIDER_RESULT: "PLANNER_PROVIDER_RESULT",
  PLANNER_RESOLUTION: "PLANNER_RESOLUTION",
  SHADOW_REENTRY: "SHADOW_REENTRY",
});

const CACHE_READ_SOURCE = Object.freeze({
  L1_MEMORY: "L1_MEMORY",
  L2_UPLOAD: "L2_UPLOAD",
  L3_SEMANTIC: "L3_SEMANTIC",
  L4_REENTRY: "L4_REENTRY",
  MISS: "MISS",
});

const DEFAULT_TTL_MS = Object.freeze({
  L2_UPLOAD: 24 * 60 * 60 * 1000,
  L3_SEMANTIC: 7 * 24 * 60 * 60 * 1000,
  L4_REENTRY: 7 * 24 * 60 * 60 * 1000,
});

const DEFAULT_MEMORY_TTL_MS = 5 * 60 * 1000;
const AES_WRAPPER_VERSION = "beebee_aes_256_gcm_cache_wrapper_v1";
const CACHEABLE_OUTCOME_STATUSES = new Set([
  "CALLED",
  "CACHE_HIT",
  "VALIDATED",
  "READY",
  "SHADOW_COMPLETED",
]);
const NON_CACHEABLE_OUTCOME_STATUSES = new Set([
  "FAILED_SAFE",
  "REQUIRED_NOT_RUN",
  "SKIPPED",
  "ERROR",
]);

function assertFunction(name, value) {
  if (typeof value !== "function") throw new TypeError(`${name} 함수가 필요합니다.`);
}

function asObject(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function asPositiveInteger(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? Math.floor(number) : fallback;
}

function normalizedHash(value, fieldName, { required = false } = {}) {
  const text = normalizeText(value || "").toLowerCase();
  if (!text && !required) return "";
  if (!/^[a-f0-9]{64}$/.test(text)) {
    throw new Error(`${fieldName}는 SHA-256 64자리 hex여야 합니다.`);
  }
  return text;
}

function hmacHex(secret, label, value) {
  if (!secret) throw new Error("cacheSecret이 필요합니다.");
  return crypto
    .createHmac("sha256", Buffer.from(String(secret), "utf8"))
    .update(`${label}:${value}`)
    .digest("hex");
}

function validateLayer(layer) {
  const normalized = normalizeText(layer);
  if (!Object.values(CACHE_LAYERS).includes(normalized)) {
    throw new Error(`지원하지 않는 cache layer입니다: ${normalized || "(empty)"}`);
  }
  return normalized;
}

function validateArtifactType(artifactType) {
  const normalized = normalizeText(artifactType);
  if (!Object.values(CACHE_ARTIFACT_TYPES).includes(normalized)) {
    throw new Error(`지원하지 않는 cache artifactType입니다: ${normalized || "(empty)"}`);
  }
  return normalized;
}

function buildHierarchicalCacheIdentity({
  tenantId,
  cacheSecret,
  layer,
  artifactType,
  uploadFingerprintSha256,
  queryJsonSha256,
  semanticProfileSha256,
  plannerInputSha256,
  plannerProposalSetSha256,
  upstreamCacheKeySha256,
  model = "",
  reasoningEffort = "",
  promptVersion = "",
  schemaVersion = "",
  plannerPolicyVersion = "",
  resolverPolicyVersion = "",
  familyPolicyVersion = "",
  feasibilityPolicyVersion = "",
  rankerPolicyVersion = "",
  extraIdentity = {},
} = {}) {
  const tenant = normalizeText(tenantId);
  if (!tenant) throw new Error("tenantId가 필요합니다.");
  const normalizedLayer = validateLayer(layer);
  const normalizedArtifactType = validateArtifactType(artifactType);
  const scope = {
    uploadFingerprintSha256: normalizedHash(
      uploadFingerprintSha256,
      "uploadFingerprintSha256",
    ),
    queryJsonSha256: normalizedHash(queryJsonSha256, "queryJsonSha256"),
    semanticProfileSha256: normalizedHash(
      semanticProfileSha256,
      "semanticProfileSha256",
    ),
    plannerInputSha256: normalizedHash(
      plannerInputSha256,
      "plannerInputSha256",
    ),
    plannerProposalSetSha256: normalizedHash(
      plannerProposalSetSha256,
      "plannerProposalSetSha256",
    ),
    upstreamCacheKeySha256: normalizedHash(
      upstreamCacheKeySha256,
      "upstreamCacheKeySha256",
    ),
  };

  if (
    normalizedLayer === CACHE_LAYERS.L2_UPLOAD &&
    (!scope.uploadFingerprintSha256 || !scope.queryJsonSha256)
  ) {
    throw new Error(
      "L2_UPLOAD identity에는 uploadFingerprintSha256와 queryJsonSha256가 필요합니다.",
    );
  }
  if (
    normalizedLayer === CACHE_LAYERS.L3_SEMANTIC &&
    !(
      (scope.semanticProfileSha256 && scope.plannerInputSha256) ||
      scope.upstreamCacheKeySha256
    )
  ) {
    throw new Error(
      "L3_SEMANTIC identity에는 semanticProfileSha256+plannerInputSha256 또는 upstreamCacheKeySha256가 필요합니다.",
    );
  }
  if (
    normalizedLayer === CACHE_LAYERS.L4_REENTRY &&
    !scope.plannerProposalSetSha256
  ) {
    throw new Error(
      "L4_REENTRY identity에는 plannerProposalSetSha256가 필요합니다.",
    );
  }

  const material = {
    version: QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_KEY_VERSION,
    layer: normalizedLayer,
    artifactType: normalizedArtifactType,
    scope,
    contract: {
      model: normalizeText(model),
      reasoningEffort: normalizeText(reasoningEffort),
      promptVersion: normalizeText(promptVersion),
      schemaVersion: normalizeText(schemaVersion),
      plannerPolicyVersion: normalizeText(plannerPolicyVersion),
      resolverPolicyVersion: normalizeText(resolverPolicyVersion),
      familyPolicyVersion: normalizeText(familyPolicyVersion),
      feasibilityPolicyVersion: normalizeText(feasibilityPolicyVersion),
      rankerPolicyVersion: normalizeText(rankerPolicyVersion),
    },
    extraIdentity: asObject(extraIdentity),
  };
  const tenantDigest = hmacHex(cacheSecret, "tenant", tenant);
  const materialSha256 = sha256(material);
  const keyDigest = hmacHex(
    cacheSecret,
    "hierarchical-cache-key",
    stableStringify({ tenantDigest, material }),
  );
  return Object.freeze({
    version: QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_KEY_VERSION,
    layer: normalizedLayer,
    artifactType: normalizedArtifactType,
    tenantDigest,
    keyDigest,
    materialSha256,
  });
}

function normalizeAesKey(key) {
  if (Buffer.isBuffer(key)) {
    if (key.length !== 32) throw new Error("AES-256-GCM key는 32바이트여야 합니다.");
    return Buffer.from(key);
  }
  const text = normalizeText(key || "");
  if (/^[a-f0-9]{64}$/i.test(text)) return Buffer.from(text, "hex");
  try {
    const decoded = Buffer.from(text, "base64");
    if (decoded.length === 32) return decoded;
  } catch (_) {
    // fall through
  }
  throw new Error("AES-256-GCM key는 32바이트 Buffer, 64자리 hex 또는 base64여야 합니다.");
}

function aadForContext(context = {}) {
  return Buffer.from(
    stableStringify({
      purpose: normalizeText(context.purpose || "query-candidate-planner-hierarchical-cache"),
      layer: normalizeText(context.layer || ""),
      artifactType: normalizeText(context.artifactType || ""),
      tenantDigest: normalizeText(context.tenantDigest || ""),
      keyDigest: normalizeText(context.keyDigest || ""),
    }),
    "utf8",
  );
}

function parseEncryptedWrapper(encrypted) {
  let wrapper;
  try {
    wrapper = JSON.parse(Buffer.from(encrypted).toString("utf8"));
  } catch (cause) {
    const error = new Error("암호화 cache wrapper JSON이 유효하지 않습니다.");
    error.code = "CACHE_WRAPPER_INVALID";
    error.cause = cause;
    throw error;
  }
  if (wrapper.version !== AES_WRAPPER_VERSION) {
    const error = new Error("암호화 cache wrapper version이 유효하지 않습니다.");
    error.code = "CACHE_WRAPPER_VERSION_INVALID";
    throw error;
  }
  return wrapper;
}

function createAes256GcmCacheCodec({ key, keyId = "default" } = {}) {
  const keyBuffer = normalizeAesKey(key);
  const normalizedKeyId = normalizeText(keyId) || "default";

  async function encryptBuffer(plaintext, context = {}) {
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv("aes-256-gcm", keyBuffer, iv);
    cipher.setAAD(aadForContext(context));
    const ciphertext = Buffer.concat([
      cipher.update(Buffer.from(plaintext)),
      cipher.final(),
    ]);
    const tag = cipher.getAuthTag();
    return Buffer.from(
      `${JSON.stringify({
        version: AES_WRAPPER_VERSION,
        keyId: normalizedKeyId,
        iv: iv.toString("base64"),
        tag: tag.toString("base64"),
        ciphertext: ciphertext.toString("base64"),
      })}\n`,
      "utf8",
    );
  }

  async function decryptBufferWithMetadata(encrypted, context = {}) {
    const wrapper = parseEncryptedWrapper(encrypted);
    if (wrapper.keyId !== normalizedKeyId) {
      const error = new Error("암호화 cache keyId가 일치하지 않습니다.");
      error.code = "CACHE_KEY_ID_MISMATCH";
      throw error;
    }
    try {
      const decipher = crypto.createDecipheriv(
        "aes-256-gcm",
        keyBuffer,
        Buffer.from(wrapper.iv, "base64"),
      );
      decipher.setAAD(aadForContext(context));
      decipher.setAuthTag(Buffer.from(wrapper.tag, "base64"));
      return {
        plaintext: Buffer.concat([
          decipher.update(Buffer.from(wrapper.ciphertext, "base64")),
          decipher.final(),
        ]),
        keyId: normalizedKeyId,
        needsRotation: false,
      };
    } catch (cause) {
      const error = new Error("암호화 cache 인증 또는 복호화에 실패했습니다.");
      error.code = "CACHE_DECRYPT_FAILED";
      error.cause = cause;
      throw error;
    }
  }

  return {
    encryptBuffer,
    async decryptBuffer(encrypted, context = {}) {
      const result = await decryptBufferWithMetadata(encrypted, context);
      return result.plaintext;
    },
    decryptBufferWithMetadata,
    keyId: normalizedKeyId,
  };
}

function createRotatingAes256GcmCacheCodec({ primary, legacy = [] } = {}) {
  const primaryConfig = asObject(primary);
  if (!primaryConfig.key) throw new Error("rotation primary key가 필요합니다.");
  const primaryCodec = createAes256GcmCacheCodec(primaryConfig);
  const codecs = new Map([[primaryCodec.keyId, primaryCodec]]);
  for (const entry of Array.isArray(legacy) ? legacy : []) {
    const codec = createAes256GcmCacheCodec(asObject(entry));
    if (codecs.has(codec.keyId)) {
      throw new Error(`중복 cache rotation keyId입니다: ${codec.keyId}`);
    }
    codecs.set(codec.keyId, codec);
  }
  return {
    encryptBuffer: primaryCodec.encryptBuffer,
    async decryptBufferWithMetadata(encrypted, context = {}) {
      const wrapper = parseEncryptedWrapper(encrypted);
      const codec = codecs.get(normalizeText(wrapper.keyId));
      if (!codec) {
        const error = new Error("암호화 cache rotation keyId를 찾을 수 없습니다.");
        error.code = "CACHE_ROTATION_KEY_NOT_FOUND";
        throw error;
      }
      const result = await codec.decryptBufferWithMetadata(encrypted, context);
      return {
        ...result,
        needsRotation: codec.keyId !== primaryCodec.keyId,
        primaryKeyId: primaryCodec.keyId,
      };
    },
    async decryptBuffer(encrypted, context = {}) {
      const result = await this.decryptBufferWithMetadata(encrypted, context);
      return result.plaintext;
    },
    keyId: primaryCodec.keyId,
    activeKeyId: primaryCodec.keyId,
    legacyKeyIds: [...codecs.keys()].filter((keyId) => keyId !== primaryCodec.keyId),
  };
}

function privacyBoundaryValid(privacy = {}) {
  return (
    privacy.rawRowsIncluded !== true &&
    privacy.sampleValuesIncluded !== true &&
    privacy.originalFileIncluded !== true &&
    privacy.fileNameIncluded !== true &&
    privacy.rawRowsSent !== true &&
    privacy.sampleValuesSent !== true &&
    privacy.originalFileSent !== true &&
    privacy.fileNameSent !== true
  );
}

function evaluateCacheability(metadata = {}) {
  const normalized = asObject(metadata);
  const outcomeStatus = normalizeText(normalized.outcomeStatus || "");
  const failureCode = normalizeText(normalized.failureCode || "");
  if (normalized.cacheable !== true) {
    return { cacheable: false, reason: "CACHEABLE_FLAG_REQUIRED" };
  }
  if (normalized.validationValid !== true) {
    return { cacheable: false, reason: "VALIDATION_REQUIRED" };
  }
  if (!privacyBoundaryValid(normalized.privacy)) {
    return { cacheable: false, reason: "PRIVACY_BOUNDARY_VIOLATION" };
  }
  if (failureCode) {
    return { cacheable: false, reason: "FAILURE_CODE_PRESENT" };
  }
  if (NON_CACHEABLE_OUTCOME_STATUSES.has(outcomeStatus)) {
    return { cacheable: false, reason: "NON_CACHEABLE_OUTCOME" };
  }
  if (!CACHEABLE_OUTCOME_STATUSES.has(outcomeStatus)) {
    return { cacheable: false, reason: "UNRECOGNIZED_CACHEABLE_OUTCOME" };
  }
  return { cacheable: true, reason: "CACHEABLE" };
}

function validateIdentity(identity = {}) {
  if (identity.version !== QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_KEY_VERSION) {
    throw new Error("hierarchical cache identity version이 유효하지 않습니다.");
  }
  validateLayer(identity.layer);
  validateArtifactType(identity.artifactType);
  for (const field of ["tenantDigest", "keyDigest", "materialSha256"]) {
    normalizedHash(identity[field], field, { required: true });
  }
  return identity;
}

function safePathSegment(value, fieldName) {
  const text = normalizeText(value);
  if (!/^[A-Z0-9_]+$/.test(text)) {
    throw new Error(`${fieldName} 경로 세그먼트가 안전하지 않습니다.`);
  }
  return text;
}

function createEncryptedHierarchicalCandidatePlannerCache({
  rootDir,
  encryptBuffer,
  decryptBuffer,
  decryptBufferWithMetadata,
  activeKeyId = "",
  rotateOnRead = true,
  auditSink,
  now = () => Date.now(),
  ttlByLayer = DEFAULT_TTL_MS,
  memoryTtlMs = DEFAULT_MEMORY_TTL_MS,
  deleteCorrupt = true,
} = {}) {
  if (!rootDir) throw new Error("암호화 hierarchical cache rootDir이 필요합니다.");
  assertFunction("encryptBuffer", encryptBuffer);
  assertFunction("decryptBuffer", decryptBuffer);
  assertFunction("now", now);
  if (decryptBufferWithMetadata !== undefined) {
    assertFunction("decryptBufferWithMetadata", decryptBufferWithMetadata);
  }
  if (auditSink !== undefined) assertFunction("auditSink", auditSink);
  const absoluteRoot = path.resolve(rootDir);
  const memory = new Map();
  const normalizedMemoryTtlMs = asPositiveInteger(memoryTtlMs, DEFAULT_MEMORY_TTL_MS);
  const normalizedActiveKeyId = normalizeText(activeKeyId || "");

  function emitAudit(action, identity = {}, details = {}) {
    if (typeof auditSink !== "function") return;
    const event = {
      version: QUERY_CANDIDATE_PLANNER_CACHE_AUDIT_EVENT_VERSION,
      timestamp: Number(now()),
      action: normalizeText(action),
      layer: normalizeText(identity.layer || details.layer || ""),
      artifactType: normalizeText(
        identity.artifactType || details.artifactType || "",
      ),
      tenantDigest: normalizeText(
        identity.tenantDigest || details.tenantDigest || "",
      ),
      keyDigest: normalizeText(identity.keyDigest || details.keyDigest || ""),
      source: normalizeText(details.source || ""),
      reason: normalizeText(details.reason || ""),
      payloadSha256: normalizeText(details.payloadSha256 || ""),
      previousKeyId: normalizeText(details.previousKeyId || ""),
      activeKeyId: normalizeText(details.activeKeyId || normalizedActiveKeyId),
    };
    try {
      auditSink(Object.freeze(event));
    } catch (_) {
      // Operational audit failures must never break cache behavior.
    }
  }

  function memoryKey(identity) {
    return [
      identity.tenantDigest,
      identity.layer,
      identity.artifactType,
      identity.keyDigest,
    ].join(":");
  }

  function filePath(identity) {
    validateIdentity(identity);
    return path.join(
      absoluteRoot,
      identity.tenantDigest,
      safePathSegment(identity.layer, "layer"),
      safePathSegment(identity.artifactType, "artifactType"),
      `${identity.keyDigest}.enc`,
    );
  }

  function codecContext(identity) {
    return {
      purpose: "query-candidate-planner-hierarchical-cache",
      layer: identity.layer,
      artifactType: identity.artifactType,
      tenantDigest: identity.tenantDigest,
      keyDigest: identity.keyDigest,
    };
  }

  function persistentTtl(identity, explicitTtlMs) {
    return asPositiveInteger(
      explicitTtlMs,
      asPositiveInteger(ttlByLayer?.[identity.layer], DEFAULT_TTL_MS[identity.layer]),
    );
  }

  function putMemory(identity, envelope) {
    const current = Number(now());
    memory.set(memoryKey(identity), {
      envelope,
      expiresAt: Math.min(
        Number(envelope.expiresAt || current),
        current + normalizedMemoryTtlMs,
      ),
    });
  }

  function removeFile(target) {
    if (!fs.existsSync(target)) return false;
    fs.unlinkSync(target);
    return true;
  }

  function atomicWrite(target, bytes, current = Number(now())) {
    fs.mkdirSync(path.dirname(target), { recursive: true });
    const temporary = `${target}.${process.pid}.${current}.tmp`;
    fs.writeFileSync(temporary, bytes);
    fs.renameSync(temporary, target);
  }

  async function decryptEncrypted(encrypted, context) {
    if (typeof decryptBufferWithMetadata === "function") {
      const result = await decryptBufferWithMetadata(encrypted, context);
      if (!result || !Buffer.isBuffer(result.plaintext)) {
        throw Object.assign(
          new TypeError("decryptBufferWithMetadata.plaintext는 Buffer여야 합니다."),
          { code: "CACHE_DECRYPT_METADATA_INVALID" },
        );
      }
      return {
        plaintext: result.plaintext,
        keyId: normalizeText(result.keyId || ""),
        needsRotation: result.needsRotation === true,
      };
    }
    return {
      plaintext: await decryptBuffer(encrypted, context),
      keyId: "",
      needsRotation: false,
    };
  }

  function identityFromEnvelope(envelope = {}) {
    return {
      version: QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_KEY_VERSION,
      layer: envelope.layer,
      artifactType: envelope.artifactType,
      tenantDigest: envelope.tenantDigest,
      keyDigest: envelope.keyDigest,
      materialSha256: envelope.materialSha256,
    };
  }

  async function readEnvelopeFromFile(target, pathIdentity) {
    const encrypted = fs.readFileSync(target);
    const context = codecContext(pathIdentity);
    const decrypted = await decryptEncrypted(encrypted, context);
    const envelope = JSON.parse(Buffer.from(decrypted.plaintext).toString("utf8"));
    return { encrypted, decrypted, envelope, context };
  }

  async function get({ identity } = {}) {
    validateIdentity(identity);
    const current = Number(now());
    const key = memoryKey(identity);
    const resident = memory.get(key);
    if (resident) {
      if (resident.expiresAt > current && Number(resident.envelope.expiresAt) > current) {
        emitAudit("READ_HIT", identity, {
          source: CACHE_READ_SOURCE.L1_MEMORY,
          reason: "VALID_CACHE_HIT",
          payloadSha256: resident.envelope.payloadSha256,
        });
        return {
          hit: true,
          source: CACHE_READ_SOURCE.L1_MEMORY,
          value: resident.envelope.payload,
          metadata: resident.envelope.metadata,
          payloadSha256: resident.envelope.payloadSha256,
        };
      }
      memory.delete(key);
    }

    const target = filePath(identity);
    if (!fs.existsSync(target)) {
      emitAudit("READ_MISS", identity, { source: CACHE_READ_SOURCE.MISS, reason: "NOT_FOUND" });
      return { hit: false, source: CACHE_READ_SOURCE.MISS, reason: "NOT_FOUND" };
    }
    try {
      const { decrypted, envelope, context } = await readEnvelopeFromFile(
        target,
        identity,
      );
      if (
        envelope.version !==
        QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_ENTRY_VERSION
      ) {
        throw Object.assign(new Error("cache entry version mismatch"), {
          code: "CACHE_ENTRY_VERSION_MISMATCH",
        });
      }
      for (const field of ["layer", "artifactType", "tenantDigest", "keyDigest", "materialSha256"]) {
        if (envelope[field] !== identity[field]) {
          throw Object.assign(new Error(`cache identity mismatch: ${field}`), {
            code: "CACHE_IDENTITY_MISMATCH",
          });
        }
      }
      if (Number(envelope.expiresAt || 0) <= current) {
        removeFile(target);
        emitAudit("INVALIDATE", identity, {
          source: CACHE_READ_SOURCE.MISS,
          reason: CACHE_INVALIDATION_REASONS.TTL_EXPIRED,
          payloadSha256: envelope.payloadSha256,
        });
        return { hit: false, source: CACHE_READ_SOURCE.MISS, reason: "EXPIRED" };
      }
      if (sha256(envelope.payload) !== envelope.payloadSha256) {
        throw Object.assign(new Error("cache payload SHA mismatch"), {
          code: "CACHE_PAYLOAD_SHA_MISMATCH",
        });
      }
      const cacheability = evaluateCacheability(envelope.metadata);
      if (!cacheability.cacheable) {
        throw Object.assign(new Error(`cache policy mismatch: ${cacheability.reason}`), {
          code: "CACHE_POLICY_MISMATCH",
        });
      }
      if (decrypted.needsRotation && rotateOnRead) {
        const rotated = await encryptBuffer(decrypted.plaintext, context);
        if (!Buffer.isBuffer(rotated)) {
          throw Object.assign(new TypeError("rotation encryptBuffer는 Buffer를 반환해야 합니다."), {
            code: "CACHE_ROTATION_ENCRYPT_INVALID",
          });
        }
        atomicWrite(target, rotated, current);
        emitAudit("ROTATE", identity, {
          source: identity.layer,
          reason: CACHE_INVALIDATION_REASONS.KEY_ROTATED,
          payloadSha256: envelope.payloadSha256,
          previousKeyId: decrypted.keyId,
        });
      }
      putMemory(identity, envelope);
      emitAudit("READ_HIT", identity, {
        source: identity.layer,
        reason: "VALID_CACHE_HIT",
        payloadSha256: envelope.payloadSha256,
      });
      return {
        hit: true,
        source: identity.layer,
        value: envelope.payload,
        metadata: envelope.metadata,
        payloadSha256: envelope.payloadSha256,
        keyRotated: decrypted.needsRotation && rotateOnRead,
      };
    } catch (error) {
      memory.delete(key);
      if (deleteCorrupt) removeFile(target);
      emitAudit("INVALIDATE", identity, {
        source: CACHE_READ_SOURCE.MISS,
        reason: CACHE_INVALIDATION_REASONS.CORRUPT_ENTRY,
      });
      return {
        hit: false,
        source: CACHE_READ_SOURCE.MISS,
        reason: "CORRUPT_ENTRY",
        errorCode: normalizeText(error?.code || "CACHE_READ_FAILED"),
      };
    }
  }

  async function set({ identity, value, metadata = {}, ttlMs } = {}) {
    validateIdentity(identity);
    const cacheability = evaluateCacheability(metadata);
    if (!cacheability.cacheable) {
      emitAudit("WRITE_SKIPPED", identity, { reason: cacheability.reason });
      return { stored: false, reason: cacheability.reason };
    }
    const current = Number(now());
    const effectiveTtlMs = persistentTtl(identity, ttlMs);
    const envelope = {
      version: QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_ENTRY_VERSION,
      policyVersion:
        QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_POLICY_VERSION,
      layer: identity.layer,
      artifactType: identity.artifactType,
      tenantDigest: identity.tenantDigest,
      keyDigest: identity.keyDigest,
      materialSha256: identity.materialSha256,
      createdAt: current,
      expiresAt: current + effectiveTtlMs,
      payloadSha256: sha256(value),
      metadata: asObject(metadata),
      payload: value,
    };
    const plaintext = Buffer.from(`${JSON.stringify(envelope)}\n`, "utf8");
    const encrypted = await encryptBuffer(plaintext, codecContext(identity));
    if (!Buffer.isBuffer(encrypted)) {
      throw new TypeError("encryptBuffer는 Buffer를 반환해야 합니다.");
    }
    const target = filePath(identity);
    atomicWrite(target, encrypted, current);
    putMemory(identity, envelope);
    emitAudit("WRITE", identity, {
      source: identity.layer,
      reason: "STORED",
      payloadSha256: envelope.payloadSha256,
    });
    return {
      stored: true,
      path: target,
      layer: identity.layer,
      payloadSha256: envelope.payloadSha256,
      expiresAt: envelope.expiresAt,
    };
  }

  async function deleteEntry({ identity, reason = CACHE_INVALIDATION_REASONS.MANUAL_DELETE } = {}) {
    validateIdentity(identity);
    memory.delete(memoryKey(identity));
    const removed = removeFile(filePath(identity));
    emitAudit("INVALIDATE", identity, {
      reason,
      source: identity.layer,
    });
    return removed;
  }

  function clearMemory() {
    const count = memory.size;
    memory.clear();
    return count;
  }

  function tenantDigestFor({ tenantId, cacheSecret } = {}) {
    const tenant = normalizeText(tenantId);
    if (!tenant) throw new Error("tenantId가 필요합니다.");
    return hmacHex(cacheSecret, "tenant", tenant);
  }

  function listEncryptedFiles(directory) {
    const files = [];
    function walk(current) {
      if (!fs.existsSync(current)) return;
      for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
        const fullPath = path.join(current, entry.name);
        if (entry.isDirectory()) walk(fullPath);
        else if (entry.isFile() && entry.name.endsWith(".enc")) files.push(fullPath);
      }
    }
    walk(directory);
    return files.sort();
  }

  function pathIdentityForFile(target) {
    const relative = path.relative(absoluteRoot, target).split(path.sep);
    if (relative.length !== 4) {
      throw Object.assign(new Error("cache path depth mismatch"), {
        code: "CACHE_PATH_INVALID",
      });
    }
    const [tenantDigest, layer, artifactType, fileName] = relative;
    const keyDigest = fileName.replace(/\.enc$/, "");
    return {
      version: QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_KEY_VERSION,
      tenantDigest,
      layer,
      artifactType,
      keyDigest,
      materialSha256: "0".repeat(64),
    };
  }

  function metadataTagsMatch(metadata = {}, tags = {}) {
    const expected = asObject(tags);
    const actual = asObject(metadata.invalidationTags);
    const entries = Object.entries(expected).filter(([, value]) => normalizeText(value || ""));
    return entries.length > 0 && entries.every(([key, value]) => {
      return normalizeText(actual[key] || "").toLowerCase() === normalizeText(value).toLowerCase();
    });
  }

  async function inspectAndMaybeRemove(target, predicate, reason) {
    const pathIdentity = pathIdentityForFile(target);
    try {
      const { envelope } = await readEnvelopeFromFile(target, pathIdentity);
      const identity = identityFromEnvelope(envelope);
      validateIdentity(identity);
      if (!predicate(envelope, identity)) return { removed: false, corrupt: false };
      memory.delete(memoryKey(identity));
      const removed = removeFile(target);
      if (removed) emitAudit("INVALIDATE", identity, { reason, source: identity.layer });
      return { removed, corrupt: false };
    } catch (error) {
      const removed = deleteCorrupt ? removeFile(target) : false;
      emitAudit("INVALIDATE", pathIdentity, {
        reason: CACHE_INVALIDATION_REASONS.CORRUPT_ENTRY,
        source: CACHE_READ_SOURCE.MISS,
      });
      return { removed, corrupt: true };
    }
  }

  function invalidateTenant({ tenantId, cacheSecret, reason = CACHE_INVALIDATION_REASONS.TENANT_DELETED } = {}) {
    const tenantDigest = tenantDigestFor({ tenantId, cacheSecret });
    const target = path.join(absoluteRoot, tenantDigest);
    let diskRemoved = false;
    if (fs.existsSync(target)) {
      fs.rmSync(target, { recursive: true, force: true });
      diskRemoved = true;
    }
    let memoryRemoved = 0;
    for (const key of memory.keys()) {
      if (key.startsWith(`${tenantDigest}:`)) {
        memory.delete(key);
        memoryRemoved += 1;
      }
    }
    emitAudit("INVALIDATE_TENANT", { tenantDigest }, { reason });
    return { tenantDigest, diskRemoved, memoryRemoved, reason };
  }

  async function invalidateByTags({ tenantId, cacheSecret, tags, reason = CACHE_INVALIDATION_REASONS.UPLOAD_DELETED } = {}) {
    const tenantDigest = tenantDigestFor({ tenantId, cacheSecret });
    const tenantRoot = path.join(absoluteRoot, tenantDigest);
    const files = listEncryptedFiles(tenantRoot);
    let removed = 0;
    let corruptRemoved = 0;
    for (const target of files) {
      const result = await inspectAndMaybeRemove(
        target,
        (envelope) => metadataTagsMatch(envelope.metadata, tags),
        reason,
      );
      if (result.removed) removed += 1;
      if (result.corrupt) corruptRemoved += 1;
    }
    return { tenantDigest, scanned: files.length, removed, corruptRemoved, reason };
  }

  async function invalidateLayer({ tenantId, cacheSecret, layer, artifactType = "", reason = CACHE_INVALIDATION_REASONS.LAYER_INVALIDATED } = {}) {
    const tenantDigest = tenantDigestFor({ tenantId, cacheSecret });
    const normalizedLayer = validateLayer(layer);
    const normalizedArtifact = normalizeText(artifactType || "");
    if (normalizedArtifact) validateArtifactType(normalizedArtifact);
    const base = path.join(
      absoluteRoot,
      tenantDigest,
      normalizedLayer,
      ...(normalizedArtifact ? [normalizedArtifact] : []),
    );
    const files = listEncryptedFiles(base);
    let removed = 0;
    for (const target of files) {
      const pathIdentity = pathIdentityForFile(target);
      removeFile(target);
      memory.delete(memoryKey(pathIdentity));
      emitAudit("INVALIDATE", pathIdentity, { reason, source: normalizedLayer });
      removed += 1;
    }
    return { tenantDigest, layer: normalizedLayer, artifactType: normalizedArtifact, removed, reason };
  }

  async function sweepExpired({ tenantId, cacheSecret } = {}) {
    const tenantDigest = tenantId
      ? tenantDigestFor({ tenantId, cacheSecret })
      : "";
    const base = tenantDigest ? path.join(absoluteRoot, tenantDigest) : absoluteRoot;
    const files = listEncryptedFiles(base);
    const current = Number(now());
    let expiredRemoved = 0;
    let corruptRemoved = 0;
    for (const target of files) {
      const result = await inspectAndMaybeRemove(
        target,
        (envelope) => Number(envelope.expiresAt || 0) <= current,
        CACHE_INVALIDATION_REASONS.TTL_EXPIRED,
      );
      if (result.removed && !result.corrupt) expiredRemoved += 1;
      if (result.corrupt) corruptRemoved += 1;
    }
    return { scanned: files.length, expiredRemoved, corruptRemoved, tenantDigest };
  }

  return {
    version: QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_VERSION,
    policyVersion:
      QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_POLICY_VERSION,
    operationalControlVersion:
      QUERY_CANDIDATE_PLANNER_CACHE_OPERATIONAL_CONTROL_VERSION,
    get,
    set,
    delete: deleteEntry,
    clearMemory,
    invalidateTenant,
    invalidateByTags,
    invalidateLayer,
    sweepExpired,
    pathFor(identity) {
      return filePath(identity);
    },
    stats() {
      return { memoryEntryCount: memory.size, rootDir: absoluteRoot, activeKeyId: normalizedActiveKeyId };
    },
  };
}

function createPlannerProviderHierarchicalCacheAdapter({
  hierarchicalCache,
  tenantId,
  cacheSecret,
  model = "",
  reasoningEffort = "",
  promptVersion = "query_candidate_planner_prompt_v1",
  schemaVersion = "query_candidate_planner_model_output_v1",
  plannerPolicyVersion = "conditional_llm_candidate_planner_policy_v1",
  invalidationTags = {},
  ttlMs,
} = {}) {
  if (!hierarchicalCache || typeof hierarchicalCache.get !== "function") {
    throw new TypeError("hierarchicalCache.get 함수가 필요합니다.");
  }
  if (typeof hierarchicalCache.set !== "function") {
    throw new TypeError("hierarchicalCache.set 함수가 필요합니다.");
  }

  function identityFor(cacheKey) {
    const upstreamCacheKeySha256 = normalizedHash(
      sha256(normalizeText(cacheKey || "")),
      "upstreamCacheKeySha256",
      { required: true },
    );
    return buildHierarchicalCacheIdentity({
      tenantId,
      cacheSecret,
      layer: CACHE_LAYERS.L3_SEMANTIC,
      artifactType: CACHE_ARTIFACT_TYPES.PLANNER_PROVIDER_RESULT,
      upstreamCacheKeySha256,
      model,
      reasoningEffort,
      promptVersion,
      schemaVersion,
      plannerPolicyVersion,
    });
  }

  return {
    async get(cacheKey) {
      const result = await hierarchicalCache.get({ identity: identityFor(cacheKey) });
      return result.hit ? result.value : null;
    },
    async set(cacheKey, value) {
      return hierarchicalCache.set({
        identity: identityFor(cacheKey),
        value,
        ttlMs,
        metadata: {
          cacheable: true,
          validationValid: true,
          outcomeStatus: "CALLED",
          failureCode: "",
          privacy: {
            rawRowsIncluded: false,
            sampleValuesIncluded: false,
            originalFileIncluded: false,
            fileNameIncluded: false,
          },
          invalidationTags: asObject(invalidationTags),
        },
      });
    },
    async delete(cacheKey) {
      return hierarchicalCache.delete({ identity: identityFor(cacheKey) });
    },
    identityFor,
  };
}

module.exports = {
  QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_VERSION,
  QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_KEY_VERSION,
  QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_ENTRY_VERSION,
  QUERY_CANDIDATE_PLANNER_HIERARCHICAL_CACHE_POLICY_VERSION,
  QUERY_CANDIDATE_PLANNER_CACHE_OPERATIONAL_CONTROL_VERSION,
  QUERY_CANDIDATE_PLANNER_CACHE_AUDIT_EVENT_VERSION,
  CACHE_INVALIDATION_REASONS,
  CACHE_LAYERS,
  CACHE_ARTIFACT_TYPES,
  CACHE_READ_SOURCE,
  DEFAULT_TTL_MS,
  DEFAULT_MEMORY_TTL_MS,
  AES_WRAPPER_VERSION,
  buildHierarchicalCacheIdentity,
  createAes256GcmCacheCodec,
  createRotatingAes256GcmCacheCodec,
  evaluateCacheability,
  createEncryptedHierarchicalCandidatePlannerCache,
  createPlannerProviderHierarchicalCacheAdapter,
};
