const axios = require("axios");

const TOSS_SECRET_KEY_BASIC =
  process.env.TOSS_SECRET_KEY_BASIC ||
  "Basic " + Buffer.from(process.env.TOSS_SECRET_KEY + ":").toString("base64");

const toss = axios.create({
  baseURL: "https://api.tosspayments.com",
  headers: {
    Authorization: TOSS_SECRET_KEY_BASIC,
    "Content-Type": "application/json",
  },
});

module.exports = toss;
