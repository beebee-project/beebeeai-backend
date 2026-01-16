const crypto = require("crypto");
const XLSX = require("xlsx");
const { readMetaCache, writeMetaCache } = require("./storage");
const { buildAllSheetsData } = require("./sheetMetaBuilder");

function md5Buffer(buf) {
  return crypto.createHash("md5").update(buf).digest("hex");
}

async function getOrBuildAllSheetsData(fileBuffer) {
  const hash = md5Buffer(fileBuffer);
  const cacheKey = `sheetsMeta_${hash}`;

  // meta-cache logging (optional)
  const logMetaCache =
    process.env.NODE_ENV !== "production" || process.env.META_CACHE_LOG === "1";

  // ✅ Enable meta-cache read (performance)
  const cached = await readMetaCache(cacheKey);
  if (cached && cached.allSheetsData) {
    if (logMetaCache) {
      console.log("[metaCache] HIT", hash.slice(0, 8));
    }
    return { fileHash: hash, allSheetsData: cached.allSheetsData };
  }

  const workbook = XLSX.read(fileBuffer, { type: "buffer" });
  const allSheetsData = buildAllSheetsData(workbook);

  await writeMetaCache(cacheKey, { allSheetsData });

  return { fileHash: hash, allSheetsData };
}

module.exports = { getOrBuildAllSheetsData };
