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

  // const cached = await readMetaCache(cacheKey);
  // if (cached && cached.allSheetsData) {
  //   return { fileHash: hash, allSheetsData: cached.allSheetsData };
  // }

  const workbook = XLSX.read(fileBuffer, { type: "buffer" });
  const allSheetsData = buildAllSheetsData(workbook);

  await writeMetaCache(cacheKey, { allSheetsData });

  return { fileHash: hash, allSheetsData };
}

module.exports = { getOrBuildAllSheetsData };
