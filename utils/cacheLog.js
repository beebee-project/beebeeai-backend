function shouldLogCache() {
  // Dev: always log
  if (process.env.NODE_ENV !== "production") return true;

  const v = (process.env.CACHE_LOG ?? "").trim();
  if (!v || v === "0") return false;
  if (v === "1") return true;

  const rate = Number(v);
  if (!(rate > 0)) return false;

  return Math.random() < rate;
}

module.exports = { shouldLogCache };
