const PERCENTAGE_VALUE_CONTRACT_VERSION = "percentage_value_contract_v1";
const PERCENTAGE_SCALE_RATIO = "ratio";
const PERCENTAGE_SCALE_PERCENT_POINTS = "percent-points";

function normalizeToken(value = "") {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
}

function normalizeScale(value = "") {
  const token = normalizeToken(value);
  if (["ratio", "fraction", "decimal"].includes(token)) {
    return PERCENTAGE_SCALE_RATIO;
  }
  if (
    [
      "percent",
      "percentage",
      "percentpoint",
      "percentpoints",
      "percentagepoint",
      "percentagepoints",
    ].includes(token)
  ) {
    return PERCENTAGE_SCALE_PERCENT_POINTS;
  }
  return "";
}

function resultRows(result = {}) {
  return Array.isArray(result.rows) ? result.rows : [];
}

function resultHeaders(result = {}) {
  const headers = new Set();
  for (const row of resultRows(result)) {
    for (const key of Object.keys(row || {})) headers.add(String(key));
  }
  return Array.from(headers);
}

function hasHeader(result = {}, expected = "") {
  const normalized = normalizeToken(expected);
  return resultHeaders(result).some(
    (header) => normalizeToken(header) === normalized,
  );
}

function explicitColumnScale(result = {}, header = "") {
  const maps = [
    result.percentageColumnScales,
    result.meta?.percentageColumnScales,
    result.candidate?.meta?.percentageColumnScales,
  ].filter(Boolean);

  const normalizedHeader = normalizeToken(header);
  for (const map of maps) {
    for (const [key, value] of Object.entries(map || {})) {
      if (normalizeToken(key) !== normalizedHeader) continue;
      const scale = normalizeScale(value);
      if (scale) return scale;
    }
  }
  return "";
}

function explicitResultScale(result = {}) {
  const candidates = [
    result.percentageValueScale,
    result.rateValueScale,
    result.meta?.percentageValueScale,
    result.meta?.rateValueScale,
    result.candidate?.meta?.percentageValueScale,
    result.candidate?.meta?.rateValueScale,
  ];

  for (const value of candidates) {
    const scale = normalizeScale(value);
    if (scale) return scale;
  }
  return "";
}

function isPercentPointsHeader(header = "") {
  const token = normalizeToken(header);
  return (
    token.endsWith("percent") ||
    token.endsWith("percentage") ||
    token.includes("백분율") ||
    token.includes("퍼센트") ||
    String(header || "").includes("%")
  );
}

function isRatioHeader(header = "") {
  const token = normalizeToken(header);
  return ["ratio", "fraction", "share", "비율", "구성비"].includes(token);
}

function isRateLikeHeader(header = "") {
  const token = normalizeToken(header);
  return (
    /률|율/.test(String(header || "")) ||
    /ratio|rate|share|percent|percentage/.test(token) ||
    String(header || "").includes("%")
  );
}

function pairedPercentHeaderExists(result = {}, header = "") {
  const raw = String(header || "");
  if (!raw) return false;
  return (
    hasHeader(result, `${raw}Percent`) ||
    hasHeader(result, `${raw} Percentage`) ||
    hasHeader(result, `${raw} 백분율`)
  );
}

function inferScaleFromValues(header = "", result = {}) {
  const normalizedHeader = normalizeToken(header);
  const values = [];

  for (const row of resultRows(result)) {
    for (const [key, value] of Object.entries(row || {})) {
      if (normalizeToken(key) !== normalizedHeader) continue;
      const number = Number(value);
      if (Number.isFinite(number)) values.push(Math.abs(number));
    }
  }

  if (!values.length) return "";
  return Math.max(...values) <= 1.0000001
    ? PERCENTAGE_SCALE_RATIO
    : PERCENTAGE_SCALE_PERCENT_POINTS;
}

function isCompositionRatioResult(result = {}) {
  const recipeType = normalizeToken(result.recipeType || result.recipeId || "");
  const operation = normalizeToken(result.operation || "");
  return recipeType === "compositionratio" || operation === "compositionratio";
}

function isGrowthRateResult(result = {}) {
  const recipeType = normalizeToken(result.recipeType || result.recipeId || "");
  const operation = normalizeToken(result.operation || "");
  return recipeType === "timegrowth" || operation === "growthrate";
}

function inferPercentageScale(header = "", result = {}) {
  const explicitColumn = explicitColumnScale(result, header);
  if (explicitColumn) return explicitColumn;

  if (isPercentPointsHeader(header)) {
    return PERCENTAGE_SCALE_PERCENT_POINTS;
  }

  if (pairedPercentHeaderExists(result, header)) {
    return PERCENTAGE_SCALE_RATIO;
  }

  if (isCompositionRatioResult(result) && isRatioHeader(header)) {
    return PERCENTAGE_SCALE_RATIO;
  }

  if (isRatioHeader(header)) {
    return PERCENTAGE_SCALE_RATIO;
  }

  const explicitResult = explicitResultScale(result);
  if (explicitResult && isRateLikeHeader(header)) {
    return explicitResult;
  }

  if (isGrowthRateResult(result) && isRateLikeHeader(header)) {
    return PERCENTAGE_SCALE_PERCENT_POINTS;
  }

  if (isRateLikeHeader(header)) {
    return inferScaleFromValues(header, result);
  }

  return "";
}

function percentageNumberFormat(scale = "") {
  if (scale === PERCENTAGE_SCALE_RATIO) return "0.00%";
  if (scale === PERCENTAGE_SCALE_PERCENT_POINTS) return '0.00"%"';
  return "";
}

function inferPercentageCellFormat(header = "", result = {}) {
  const scale = inferPercentageScale(header, result);
  if (!scale) return null;
  return {
    contractVersion: PERCENTAGE_VALUE_CONTRACT_VERSION,
    scale,
    z: percentageNumberFormat(scale),
    mutateValue: false,
  };
}

function applyPercentageCellFormat(cell = null, header = "", result = {}) {
  if (!cell || typeof cell.v !== "number") return null;
  const format = inferPercentageCellFormat(header, result);
  if (!format) return null;
  cell.z = format.z;
  return format;
}

module.exports = {
  PERCENTAGE_VALUE_CONTRACT_VERSION,
  PERCENTAGE_SCALE_RATIO,
  PERCENTAGE_SCALE_PERCENT_POINTS,
  normalizeScale,
  isPercentPointsHeader,
  isRatioHeader,
  isRateLikeHeader,
  inferPercentageScale,
  inferPercentageCellFormat,
  applyPercentageCellFormat,
};
