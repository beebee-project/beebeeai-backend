const DEFAULT_CRON_URL =
  "https://beebeeai-backend-production.up.railway.app/api/payments/cron/charge";

function resolveCronUrl() {
  const raw =
    process.env.SUBSCRIPTION_CRON_URL ||
    process.env.CRON_CHARGE_URL ||
    DEFAULT_CRON_URL;

  const trimmed = String(raw || "").trim();
  if (!trimmed) return DEFAULT_CRON_URL;

  if (/^https?:\/\//i.test(trimmed)) return trimmed;

  return `https://${trimmed.replace(/^\/+/, "")}`;
}

function maskUrl(url) {
  return String(url || "").replace(/(x-cron-secret=)[^&]+/i, "$1***");
}

async function readResponseBody(response) {
  const text = await response.text();
  if (!text) return { text: "", json: null };

  try {
    return { text, json: JSON.parse(text) };
  } catch (_) {
    return { text, json: null };
  }
}

async function runSubscriptionChargeCron() {
  const secret = process.env.CRON_SECRET;
  if (!secret) {
    const error = new Error("CRON_SECRET is not set");
    error.code = "CRON_SECRET_MISSING";
    throw error;
  }

  const url = resolveCronUrl();
  const timeoutMs = Number(process.env.SUBSCRIPTION_CRON_TIMEOUT_MS || 120000);
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  console.log(`[subscription-cron] POST ${maskUrl(url)}`);

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-cron-secret": secret,
      },
      signal: controller.signal,
    });

    const { text, json } = await readResponseBody(response);
    console.log(`[subscription-cron] status=${response.status}`);
    if (text) console.log(`[subscription-cron] body=${text}`);

    if (!response.ok) {
      const error = new Error(
        `Subscription cron request failed: ${response.status}`,
      );
      error.code = "SUBSCRIPTION_CRON_HTTP_ERROR";
      error.status = response.status;
      error.body = text;
      throw error;
    }

    if (json && json.ok === false) {
      const error = new Error("Subscription cron returned ok=false");
      error.code = "SUBSCRIPTION_CRON_NOT_OK";
      error.body = json;
      throw error;
    }

    if (json && Number(json.failCount || 0) > 0) {
      const error = new Error(
        `Subscription cron finished with failCount=${json.failCount}`,
      );
      error.code = "SUBSCRIPTION_CRON_CHARGE_FAILURE";
      error.body = json;
      throw error;
    }

    return json || { ok: true, rawBody: text };
  } finally {
    clearTimeout(timeout);
  }
}

if (require.main === module) {
  runSubscriptionChargeCron()
    .then((result) => {
      console.log("[subscription-cron] completed", JSON.stringify(result));
      process.exit(0);
    })
    .catch((error) => {
      console.error("[subscription-cron] failed", error?.stack || error);
      process.exit(1);
    });
}

module.exports = { runSubscriptionChargeCron, resolveCronUrl };
