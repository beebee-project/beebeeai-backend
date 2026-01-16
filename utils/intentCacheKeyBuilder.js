const crypto = require("crypto");

function _normPrompt(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function _safeStr(s) {
  return String(s == null ? "" : s).trim();
}

/**
 * Build a stable cache key for Intent-only caching.
 * NOTE: The returned key is a SHA256 hash of the payload (so raw user strings are not stored as-is).
 *
 * @param {Object} args
 * @param {number} [args.version] - cache key version
 * @param {string} args.builderType
 * @param {string} args.model
 * @param {string} args.schemaVersion
 * @param {string} args.userKey
 * @param {string} args.prompt
 * @param {string} args.sheetStateSig
 * @param {string|null} [args.targetRangeSig]
 * @returns {{ key: string, payload: object }}
 */
function buildIntentCacheKey(args = {}) {
  const payload = {
    v: Number(args.version || 1),
    builderType: _safeStr(args.builderType),
    model: _safeStr(args.model),
    schema: _safeStr(args.schemaVersion),
    user: _safeStr(args.userKey),
    prompt: _normPrompt(args.prompt),
    sheetStateSig: _safeStr(args.sheetStateSig),
    targetRangeSig: _safeStr(args.targetRangeSig || ""),
  };

  const raw = JSON.stringify(payload);
  const key = crypto.createHash("sha256").update(raw).digest("hex");
  return { key, payload };
}

module.exports = { buildIntentCacheKey };
