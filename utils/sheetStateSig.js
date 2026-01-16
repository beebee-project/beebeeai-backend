const crypto = require("crypto");

function _sha256(s) {
  return crypto.createHash("sha256").update(String(s)).digest("hex");
}

/**
 * Create a lightweight signature of the uploaded sheet "structure" only.
 * - includes sheet names, rowCount/startRow/lastDataRow
 * - includes headers (metaData keys)
 * - includes dominantType if present (optional)
 *
 * IMPORTANT: Do not include raw cell values.
 *
 * @param {Object<string, any>} allSheetsData
 * @returns {string} a short signature string (hash)
 */
function makeSheetStateSig(allSheetsData) {
  if (!allSheetsData || typeof allSheetsData !== "object") {
    return "nosheet";
  }

  const sheetNames = Object.keys(allSheetsData).sort();
  const parts = [];

  for (const name of sheetNames) {
    const info = allSheetsData[name] || {};
    const meta = info.metaData || {};

    const headers = Object.keys(meta).sort();
    // dominantType is optional; include if present (from sheetMetaBuilder)
    const types = headers
      .map((h) => meta?.[h]?.dominantType || "")
      .filter(Boolean);

    parts.push({
      sheet: name,
      rowCount: info.rowCount || 0,
      startRow: info.startRow || 0,
      lastDataRow: info.lastDataRow || 0,
      headers,
      typesTop: types.slice(0, 50), // cap for stability
    });
  }

  // Hash the structural summary
  return `hdr:${_sha256(JSON.stringify(parts)).slice(0, 16)}|sheets:${
    sheetNames.length
  }`;
}

module.exports = { makeSheetStateSig };
