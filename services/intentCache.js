/**
 * Intent Cache (Redis + Memory fallback)
 * - Default: DISABLED
 * - Stores INTENT ONLY (never formula/script)
 */

const CACHE_PREFIX = "intent_cache:";

// ============================
// Enable switch (FIXED)
// ============================
function isEnabled() {
  return process.env.INTENT_CACHE_ENABLED === "1";
}

// ============================
// Redis (optional)
// ============================
let _redis = null;
let _redisInitTried = false;

function tryInitRedis() {
  if (_redisInitTried) return _redis;
  _redisInitTried = true;

  const url = process.env.REDIS_URL;
  if (!url) return null;

  try {
    const Redis = require("ioredis");
    _redis = new Redis(url, {
      maxRetriesPerRequest: 1,
      enableReadyCheck: true,
      lazyConnect: true,
    });

    _redis.connect().catch(() => {});
    return _redis;
  } catch (e) {
    console.warn("[intentCache] Redis init failed:", e?.message || e);
    _redis = null;
    return null;
  }
}

// ============================
// In-memory fallback (L1)
// ============================
const mem = new Map();

function now() {
  return Date.now();
}

function memGet(key) {
  const hit = mem.get(key);
  if (!hit) return null;
  if (hit.exp && hit.exp < now()) {
    mem.delete(key);
    return null;
  }
  return hit.value;
}

function memSet(key, value, ttlSec) {
  const ttl = Number(ttlSec || 0);
  const exp = ttl > 0 ? now() + ttl * 1000 : 0;
  mem.set(key, { value, exp });
}

function k(key) {
  return `${CACHE_PREFIX}${key}`;
}

// ============================
// Public API
// ============================
async function get(key) {
  if (!isEnabled() || !key) return null;

  // 1️⃣ memory
  const m = memGet(key);
  if (m) return m;

  // 2️⃣ redis
  const r = tryInitRedis();
  if (!r) return null;

  try {
    const raw = await r.get(k(key));
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    memSet(key, parsed, 30); // short L1 ttl
    return parsed;
  } catch (e) {
    console.warn("[intentCache.get] error:", e?.message || e);
    return null;
  }
}

async function set(key, value, ttlSec = 600) {
  if (!isEnabled() || !key || !value) return false;

  // memory always
  memSet(key, value, Math.min(ttlSec, 60));

  const r = tryInitRedis();
  if (!r) return true;

  try {
    await r.set(k(key), JSON.stringify(value), "EX", ttlSec);
    return true;
  } catch (e) {
    console.warn("[intentCache.set] error:", e?.message || e);
    return false;
  }
}

module.exports = {
  isEnabled,
  get,
  set,
};
