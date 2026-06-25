const { FORMULA_HEURISTICS } = require("../config/formulaHeuristicsConfig");

function encodeColumn(index = 0) {
  let n = Math.max(0, Number(index) || 0);
  let result = "";

  do {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);

  return result;
}

function stripHeaderUnitSuffix(value = "") {
  return String(value || "")
    .replace(/\([^)]*\)/g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/（[^）]*）/g, "")
    .trim();
}

function normalizeHeaderForMatch(value = "") {
  return stripHeaderUnitSuffix(value)
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "")
    .trim();
}

function normalizeHeaderLoose(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "")
    .trim();
}

function getSourceColumnHeader(column = {}) {
  return (
    column.header ||
    column.originalHeader ||
    column.name ||
    column.key ||
    column.accessor ||
    ""
  );
}

function isNumericSourceColumn(column = {}) {
  const text = [
    column.type,
    column.dominantType,
    column.role,
    column.inferredRole,
    column.semanticType,
    column.header,
    column.originalHeader,
  ]
    .filter(Boolean)
    .join(" ");

  return FORMULA_HEURISTICS.numericColumnPattern.test(text);
}

function scoreHeaderMatch(queryHeader = "", sourceHeader = "") {
  const queryRaw = String(queryHeader || "").trim();
  const sourceRaw = String(sourceHeader || "").trim();

  if (!queryRaw || !sourceRaw) return 0;

  if (queryRaw === sourceRaw) return 100;
  if (queryRaw.toLowerCase() === sourceRaw.toLowerCase()) return 98;

  const queryBase = normalizeHeaderForMatch(queryRaw);
  const sourceBase = normalizeHeaderForMatch(sourceRaw);

  if (!queryBase || !sourceBase) return 0;
  if (queryBase === sourceBase) return 94;

  const queryLoose = normalizeHeaderLoose(queryRaw);
  const sourceLoose = normalizeHeaderLoose(sourceRaw);

  if (queryLoose && sourceLoose && queryLoose === sourceLoose) return 90;

  if (queryBase.length < 2 || sourceBase.length < 2) return 0;

  if (sourceBase.startsWith(queryBase)) return 78;
  if (queryBase.startsWith(sourceBase)) return 72;
  if (sourceBase.includes(queryBase)) return 66;
  if (queryBase.includes(sourceBase)) return 60;

  return 0;
}

function getSourceColumnLetter(column = {}, index = 0) {
  if (column.columnLetter) return column.columnLetter;
  if (column.letter) return column.letter;

  const rawIndex = Number(column.columnIndex);
  const zeroBasedIndex = Number.isFinite(rawIndex) ? rawIndex - 1 : index;

  return encodeColumn(zeroBasedIndex);
}

function createSourceColumnMap(table = {}) {
  const columns = Array.isArray(table.columns) ? table.columns : [];
  const map = new Map();
  const entries = [];

  columns.forEach((column, index) => {
    const header = getSourceColumnHeader(column);
    if (!header) return;

    const entry = {
      header,
      column,
      index,
      letter: getSourceColumnLetter(column, index),
      normalizedHeader: normalizeHeaderForMatch(header),
      looseHeader: normalizeHeaderLoose(header),
      isNumeric: isNumericSourceColumn(column),
    };

    map.set(header, entry);
    entries.push(entry);
  });

  map.__entries = entries;

  return map;
}

function resolveSourceColumn(columnMap, header = "", options = {}) {
  if (!columnMap || !header) return null;

  const exact = columnMap.get(header);
  if (exact) return exact;

  const entries = columnMap.__entries || Array.from(columnMap.values());
  if (!entries.length) return null;

  const scored = entries
    .map((entry) => {
      const baseScore = scoreHeaderMatch(header, entry.header);
      const numericBonus =
        options.preferNumeric && entry.isNumeric
          ? FORMULA_HEURISTICS.headerMatch.numericBonus
          : 0;

      return {
        entry,
        baseScore,
        score: baseScore > 0 ? baseScore + numericBonus : 0,
      };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score);

  if (!scored.length) return null;

  const [best, second] = scored;

  if (second && second.score === best.score) {
    return null;
  }

  if (best.baseScore < FORMULA_HEURISTICS.headerMatch.minScore) {
    return null;
  }

  return best.entry;
}

module.exports = {
  encodeColumn,
  stripHeaderUnitSuffix,
  normalizeHeaderForMatch,
  normalizeHeaderLoose,
  getSourceColumnHeader,
  getSourceColumnLetter,
  isNumericSourceColumn,
  scoreHeaderMatch,
  createSourceColumnMap,
  resolveSourceColumn,
};
