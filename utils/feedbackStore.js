const fs = require("fs");
const path = require("path");

const DATA_DIR = path.join(process.cwd(), "data");
const FEEDBACK_PATH = path.join(DATA_DIR, "feedback.json");

function ensureDir() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
}

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
  ensureDir();
  const tmp = `${FEEDBACK_PATH}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(arr, null, 2), "utf8");
  fs.renameSync(tmp, FEEDBACK_PATH);
}

function appendFeedback(event, maxItems = 5000) {
  const all = readAll();
  all.push(event);
  const pruned = all.length > maxItems ? all.slice(-maxItems) : all;
  writeAll(pruned);
}

module.exports = {
  FEEDBACK_PATH,
  appendFeedback,
};
