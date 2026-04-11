const fetch = require("node-fetch");

async function run() {
  try {
    const res = await fetch("https://beebeeai.kr/api/payments/cron/charge", {
      method: "POST",
      headers: {
        "x-cron-secret": process.env.CRON_SECRET,
      },
    });

    const data = await res.json();
    console.log("[CRON RESULT]", data);
  } catch (e) {
    console.error("[CRON ERROR]", e);
  }
}

run();
