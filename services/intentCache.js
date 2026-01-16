// -----------------------------
// Storage selection:
// 1) Redis (if REDIS_URL and ioredis installed)
// 2) In-memory Map (fallback)
// -----------------------------

const CACHE_PREFIX = "intent_cache:";

let _redis = null;
let _redisInitTried = false;

function _tryInitRedis() {
  if (_redisInitTried) return _redis;
  _redisInitTried = true;

  const url = process.env.REDIS_URL;
  if (!url) return null;

  try {
    // optional dependency
    // npm i ioredis
    const Redis = require("ioredis");
    _redis = new Redis(url, {
      maxRetriesPerRequest: 1,
      enableReadyCheck: true,
      lazyConnect: true,
    });
    // best-effort connect (non-blocking on first use)
    _redis.connect().catch(() => {});
    return _redis;
  } catch (e) {
    console.warn(
      "[intentCache] Redis not available (install ioredis or set REDIS_URL):",
      e?.message || e
    );
    _redis = null;
    return null;
  }
}

// In-memory L1 cache with TTL
// key -> { value: any, exp: number }
const _mem = new Map();

function _nowMs() {
  return Date.now();
}

function _memGet(k) {
  const hit = _mem.get(k);
  if (!hit) return null;
  if (hit.exp && hit.exp <= _nowMs()) {
    _mem.delete(k);
    return null;
  }
  return hit.value ?? null;
}

function _memSet(k, v, ttlSec) {
  const ttl = Number(ttlSec || 0);
  const exp = ttl > 0 ? _nowMs() + ttl * 1000 : 0;
  _mem.set(k, { value: v, exp });
}

function _k(key) {
  return `${CACHE_PREFIX}${key}`;
}

/**
 * @param {string} key
 * @returns {Promise<{intent: object, meta?: object} | null>}
 */
async function get(key) {
  if (!isEnabled()) return null;
  if (!key) return null;

  // 1) L1 memory
  const memHit = _memGet(key);
  if (memHit) return memHit;

  // 2) Redis
  const r = _tryInitRedis();
  if (!r) return null;

  try {
    const raw = await r.get(_k(key));
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    // hydrate L1 with a short TTL to reduce redis chatter
    _memSet(key, parsed, 30);
    return parsed;
  } catch (e) {
    console.warn("[intentCache.get] error:", e?.message || e);
    return null;
  }
}

/**
 * @param {string} key
 * @param {{intent: object, meta?: object}} value
 * @param {number} ttlSec
 * @returns {Promise<boolean>}
 */
async function set(key, value, ttlSec = 600) {
  if (!isEnabled()) return false;
  if (!key || !value || typeof value !== "object") return false;

  // write L1 always
  _memSet(key, value, Math.min(Number(ttlSec || 0), 60) || 30);

  // write Redis if available
  const r = _tryInitRedis();
  if (!r) return true;

  try {
    const ttl = Math.max(1, Number(ttlSec || 600));
    await r.set(_k(key), JSON.stringify(value), "EX", ttl);
    return true;
  } catch (e) {
    console.warn("[intentCache.set] error:", e?.message || e);
    return false;
  }
}

module.exports = { isEnabled, get, set };
