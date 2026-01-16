const fs = require("fs");
const path = require("path");

const FEEDBACK_PATH = path.join(process.cwd(), "feedback.json"); // ✅ 루트에 저장
const MAX_ITEMS = Number(process.env.FEEDBACK_MAX_ITEMS || "5000");

function readAll() {
  try {
    if (!fs.existsSync(FEEDBACK_PATH)) return [];
    const raw = fs.readFileSync(FEEDBACK_PATH, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeAll(arr) {
  const tmp = `${FEEDBACK_PATH}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(arr, null, 2), "utf8");
  fs.renameSync(tmp, FEEDBACK_PATH);
}

function appendFeedback(event) {
  const all = readAll();
  all.push(event);
  const pruned = all.length > MAX_ITEMS ? all.slice(-MAX_ITEMS) : all;
  writeAll(pruned);
}

module.exports = { FEEDBACK_PATH, appendFeedback };
