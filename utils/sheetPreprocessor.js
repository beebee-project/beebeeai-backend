const crypto = require("crypto");
const XLSX = require("xlsx");
const { readMetaCache, writeMetaCache } = require("./storage");
const { buildAllSheetsData } = require("./sheetMetaBuilder");
const { shouldLogCache } = require("./cacheLog");

function md5Buffer(buf) {
  return crypto.createHash("md5").update(buf).digest("hex");
}

async function getOrBuildAllSheetsData(fileBuffer) {
  const hash = md5Buffer(fileBuffer);
  const cacheKey = `sheetsMeta_${hash}`;

  // meta-cache logging (prod: sampling via META_CACHE_LOG)
  // - dev: always log
  // - prod:
  //    META_CACHE_LOG unset/0 -> off
  //    META_CACHE_LOG=1       -> always on
  //    META_CACHE_LOG=0.01    -> 1% sampling
  function shouldLogMetaCache() {
    if (process.env.NODE_ENV !== "production") return true;
    const v = (process.env.META_CACHE_LOG ?? "").trim();
    if (!v || v === "0") return false;
    if (v === "1") return true;
    const rate = Number(v);
    if (!(rate > 0)) return false;
    return Math.random() < rate;
  }

  // ✅ Enable meta-cache read (performance)
  const cached = await readMetaCache(cacheKey);
  if (cached && cached.allSheetsData) {
    if (shouldLogCache()) {
      console.log("[metaCache] HIT", hash.slice(0, 8));
    }
    return { fileHash: hash, allSheetsData: cached.allSheetsData };
  }

  if (shouldLogCache()) {
    console.log("[metaCache] MISS", hash.slice(0, 8));
  }

  const workbook = XLSX.read(fileBuffer, { type: "buffer" });
  const allSheetsData = buildAllSheetsData(workbook);

  await writeMetaCache(cacheKey, { allSheetsData });

  return { fileHash: hash, allSheetsData };
}

module.exports = { getOrBuildAllSheetsData };
