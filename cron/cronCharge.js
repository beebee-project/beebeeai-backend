const fetch = require("node-fetch");

async function run() {
  try {
    const res = await fetch(
      "https://beebeeai-backend-production.up.railway.app/api/payments/cron/charge",
      {
        method: "POST",
        headers: {
          "x-cron-secret": process.env.CRON_SECRET,
        },
      },
    );

    const text = await res.text();
    console.log("[CRON STATUS]", res.status);
    console.log("[CRON BODY]", text);

    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch (e) {
      console.error("[CRON PARSE ERROR]", e.message);
    }

    if (!res.ok) {
      throw new Error(`cron charge failed: status=${res.status} body=${text}`);
    }

    console.log("[CRON RESULT]", data);
  } catch (e) {
    console.error("[CRON ERROR]", e);
    process.exitCode = 1;
  }
}

run();
