module.exports = function requireCronSecret(req, res, next) {
  const secret = process.env.CRON_SECRET;
  if (!secret) return res.status(500).json({ error: "CRON_SECRET is not set" });

  const provided =
    req.get("x-cron-secret") ||
    (req.get("authorization") || "").replace(/^Bearer\s+/i, "");

  if (provided !== secret) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
};
