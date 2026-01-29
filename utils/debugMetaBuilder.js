function safeMs(x) {
  return typeof x === "number" && Number.isFinite(x) ? x : null;
}

function buildDebugMeta({
  rawReason,
  cacheHit,
  intentOp,
  intentCacheKey,
  validator,
  timing,
  extra,
}) {
  return {
    rawReason: rawReason || undefined,
    cacheHit: cacheHit ?? undefined,
    intentOp: intentOp || undefined,
    intentCacheKey: intentCacheKey || undefined,

    validatorOk: validator?.ok ?? undefined,
    validatorKind: validator?.kind ?? undefined,
    validatorFailPoints:
      Array.isArray(validator?.issues) && validator.issues.length
        ? validator.issues
        : undefined,

    timingMs: timing
      ? {
          preprocess: safeMs(timing.preprocess),
          intent: safeMs(timing.intent),
          build: safeMs(timing.build),
          total: safeMs(timing.total),
        }
      : undefined,

    ...(extra && typeof extra === "object" ? extra : {}),
  };
}

module.exports = { buildDebugMeta };
