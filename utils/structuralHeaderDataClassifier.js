"use strict";

const STRUCTURAL_HEADER_DATA_CLASSIFIER_VERSION =
  "structural_header_data_classifier_v1";

function compactText(value = "") {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function isEmpty(value) {
  return value == null || compactText(value) === "";
}

function isBooleanLike(value) {
  if (typeof value === "boolean") return true;
  const text = compactText(value).toLowerCase();
  return [
    "true",
    "false",
    "yes",
    "no",
    "y",
    "n",
    "예",
    "아니오",
  ].includes(text);
}

function isNumberLike(value) {
  if (value instanceof Date || typeof value === "boolean") return false;
  if (typeof value === "number") return Number.isFinite(value);

  const text = compactText(value).replace(/,/g, "");
  if (!text || /[^\d.+\-eE]/.test(text)) return false;
  return Number.isFinite(Number(text));
}

function isDateLike(value) {
  if (value instanceof Date) return true;
  const text = compactText(value);
  if (!text) return false;

  return /^(19|20)\d{2}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])(?:[ T].*)?$/.test(
    text,
  );
}

function isTimeLike(value) {
  const text = compactText(value);
  return /^([01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d)?$/.test(text);
}

function isPeriodLike(value) {
  const text = compactText(value).toUpperCase();
  if (!text) return false;

  return (
    /^(19|20)\d{2}$/.test(text) ||
    /^(19|20)\d{2}\s*년$/.test(text) ||
    /^(19|20)\d{2}\s*[-./년]\s*(0?[1-9]|1[0-2])\s*월?$/.test(text) ||
    /^(0?[1-9]|1[0-2])\s*월$/.test(text) ||
    /^(Q[1-4]|[1-4]\s*분기)$/.test(text)
  );
}

function classifyCellStructuralType(value) {
  if (isEmpty(value)) return "empty";
  if (isBooleanLike(value)) return "boolean";
  if (isDateLike(value)) return "date";
  if (isTimeLike(value)) return "time";
  if (isPeriodLike(value)) return "period";
  if (isNumberLike(value)) return "number";
  return "text";
}

function broadStructuralType(type = "") {
  if (type === "period" || type === "date" || type === "time") {
    return "temporal";
  }
  if (type === "number") return "numeric";
  if (type === "boolean") return "boolean";
  if (type === "text") return "text";
  return "empty";
}

function looksLikeSentenceOrNote(value = "") {
  const text = compactText(value);
  if (!text) return false;

  return Boolean(
    text.length >= 50 ||
      /[.!?。]|입니다|합니다|관련하여|참고|주의|설명/.test(text) ||
      /^\s*(※|\*|주\s*[:：.]|note\b|remark\b)/i.test(text),
  );
}

function splitConcatenatedParts(value = "") {
  return compactText(value)
    .split(/[_\n\r|]+/)
    .map((part) => compactText(part))
    .filter(Boolean);
}

function looksLikeSuspiciousConcatenatedHeader(value = "") {
  const parts = splitConcatenatedParts(value);
  if (parts.length < 3) return false;

  const distinct = new Set(parts.map((part) => part.toLowerCase()));
  if (distinct.size < 3) return false;

  const valueLike = parts.filter((part) => {
    const type = classifyCellStructuralType(part);
    return ["number", "date", "time", "period"].includes(type);
  }).length;
  const compactTokens = parts.filter(
    (part) => part.length <= 30 && !looksLikeSentenceOrNote(part),
  ).length;

  return valueLike >= 2 || compactTokens / parts.length >= 0.8;
}

function headerLexicalToken(value) {
  const text = compactText(value);
  if (!text || looksLikeSentenceOrNote(text)) return false;
  if (looksLikeSuspiciousConcatenatedHeader(text)) return false;

  const type = classifyCellStructuralType(text);
  if (type === "text") return text.length <= 48;

  // 연도·월·분기는 넓은 표의 실제 헤더일 수 있다.
  return type === "period";
}

function rowBounds(row = []) {
  const indexes = [];
  for (let index = 0; index < row.length; index += 1) {
    if (!isEmpty(row[index])) indexes.push(index);
  }

  if (!indexes.length) {
    return {
      min: -1,
      max: -1,
      span: 0,
      fillRatio: 0,
    };
  }

  const min = Math.min(...indexes);
  const max = Math.max(...indexes);
  const span = max - min + 1;

  return {
    min,
    max,
    span,
    fillRatio: span ? indexes.length / span : 0,
  };
}

function buildRowStructuralProfile(row = []) {
  const types = (row || []).map(classifyCellStructuralType);
  const nonEmptyIndexes = [];

  types.forEach((type, index) => {
    if (type !== "empty") nonEmptyIndexes.push(index);
  });

  const nonEmpty = nonEmptyIndexes.length;
  const count = (type) => types.filter((item) => item === type).length;
  const textCount = count("text");
  const periodCount = count("period");
  const numberCount = count("number");
  const dateCount = count("date");
  const timeCount = count("time");
  const booleanCount = count("boolean");
  const measureValueCount =
    numberCount + dateCount + timeCount + booleanCount;
  const dataValueCount = measureValueCount + periodCount;
  const headerLexicalCount = nonEmptyIndexes.filter((index) =>
    headerLexicalToken(row[index]),
  ).length;
  const suspiciousConcatenatedCount = nonEmptyIndexes.filter((index) =>
    looksLikeSuspiciousConcatenatedHeader(row[index]),
  ).length;
  const bounds = rowBounds(row);
  const firstIndex = nonEmptyIndexes.length ? nonEmptyIndexes[0] : -1;

  return {
    row,
    types,
    broadTypes: types.map(broadStructuralType),
    nonEmptyIndexes,
    nonEmpty,
    firstIndex,
    firstType: firstIndex >= 0 ? types[firstIndex] : "empty",
    textCount,
    periodCount,
    numberCount,
    dateCount,
    timeCount,
    booleanCount,
    measureValueCount,
    dataValueCount,
    textRatio: nonEmpty ? textCount / nonEmpty : 0,
    periodRatio: nonEmpty ? periodCount / nonEmpty : 0,
    measureValueRatio: nonEmpty ? measureValueCount / nonEmpty : 0,
    dataValueRatio: nonEmpty ? dataValueCount / nonEmpty : 0,
    headerLexicalCount,
    headerLexicalRatio: nonEmpty ? headerLexicalCount / nonEmpty : 0,
    suspiciousConcatenatedCount,
    suspiciousConcatenatedRatio: nonEmpty
      ? suspiciousConcatenatedCount / nonEmpty
      : 0,
    fillRatio: bounds.fillRatio,
    span: bounds.span,
  };
}

function structuralTypeCompatible(left = "", right = "") {
  if (!left || !right || left === "empty" || right === "empty") {
    return false;
  }
  if (left === right) return true;

  const leftBroad = broadStructuralType(left);
  const rightBroad = broadStructuralType(right);

  // 숫자와 기간을 같은 것으로 보지 않는다. 이 차이가 실제 헤더와
  // 데이터 행의 타입 대비를 보존한다.
  return leftBroad === rightBroad;
}

function rowTypeSignatureSimilarity(leftRow = [], rightRow = []) {
  const left = buildRowStructuralProfile(leftRow);
  const right = buildRowStructuralProfile(rightRow);
  if (!left.nonEmpty || !right.nonEmpty) return 0;

  const union = new Set([
    ...left.nonEmptyIndexes,
    ...right.nonEmptyIndexes,
  ]);
  let comparable = 0;
  let matched = 0;

  for (const index of union) {
    const leftType = left.types[index] || "empty";
    const rightType = right.types[index] || "empty";
    if (leftType === "empty" || rightType === "empty") continue;

    comparable += 1;
    if (structuralTypeCompatible(leftType, rightType)) {
      matched += 1;
    }
  }

  if (!comparable) return 0;

  const widthCompatibility =
    Math.min(left.nonEmpty, right.nonEmpty) /
    Math.max(left.nonEmpty, right.nonEmpty);

  return (matched / comparable) * widthCompatibility;
}

function nonEmptyRows(rows = [], maxRows = 4) {
  return (rows || [])
    .filter((row) => buildRowStructuralProfile(row).nonEmpty > 0)
    .slice(0, maxRows);
}

function average(values = []) {
  const safe = values.filter((value) => Number.isFinite(Number(value)));
  if (!safe.length) return 0;
  return safe.reduce((sum, value) => sum + Number(value), 0) / safe.length;
}

function scoreSubsequentRowRepeatability(row = [], nextRows = []) {
  const following = nonEmptyRows(nextRows, 4);
  if (!following.length) {
    return {
      candidateSimilarity: 0,
      followingSimilarity: 0,
      supportCount: 0,
      followingRowCount: 0,
    };
  }

  const candidateSimilarities = following.map((next) =>
    rowTypeSignatureSimilarity(row, next),
  );
  const adjacentSimilarities = [];

  for (let index = 1; index < following.length; index += 1) {
    adjacentSimilarities.push(
      rowTypeSignatureSimilarity(
        following[index - 1],
        following[index],
      ),
    );
  }

  return {
    candidateSimilarity: average(candidateSimilarities),
    followingSimilarity: adjacentSimilarities.length
      ? average(adjacentSimilarities)
      : candidateSimilarities[0] || 0,
    supportCount: candidateSimilarities.filter(
      (similarity) => similarity >= 0.7,
    ).length,
    followingRowCount: following.length,
  };
}

function analyzeHeaderDataStructure(row = [], nextRows = []) {
  const profile = buildRowStructuralProfile(row);
  const repeatability = scoreSubsequentRowRepeatability(
    row,
    nextRows,
  );

  const firstIsRecordValue = [
    "number",
    "date",
    "time",
    "period",
  ].includes(profile.firstType);

  const enoughRepeatability =
    repeatability.supportCount >=
      Math.min(2, repeatability.followingRowCount) &&
    repeatability.candidateSimilarity >= 0.68;

  const likelyDataRecord = Boolean(
    profile.nonEmpty >= 3 &&
      profile.measureValueRatio >= 0.2 &&
      enoughRepeatability &&
      (firstIsRecordValue ||
        profile.headerLexicalRatio < 0.72 ||
        profile.dataValueRatio >= 0.5),
  );

  const followingLooksLikeStableData =
    repeatability.followingRowCount >= 2 &&
    repeatability.followingSimilarity >= 0.72;

  const likelySchemaHeader = Boolean(
    profile.nonEmpty >= 2 &&
      profile.headerLexicalRatio >= 0.65 &&
      profile.measureValueRatio <= 0.15 &&
      !likelyDataRecord &&
      (followingLooksLikeStableData ||
        repeatability.candidateSimilarity <= 0.6),
  );

  let scoreAdjustment = 0;
  const reasons = [];

  if (likelySchemaHeader) {
    scoreAdjustment += 8;
    reasons.push("SCHEMA_LIKE_HEADER_EVIDENCE");
  }

  if (likelyDataRecord) {
    scoreAdjustment -= 40;
    reasons.push("REPEATABLE_DATA_RECORD_PENALTY");
  }

  if (profile.suspiciousConcatenatedRatio >= 0.25) {
    scoreAdjustment -= 18;
    reasons.push("MULTI_VALUE_CONCATENATION_PENALTY");
  }

  return {
    version: STRUCTURAL_HEADER_DATA_CLASSIFIER_VERSION,
    likelyDataRecord,
    likelySchemaHeader,
    scoreAdjustment,
    reasons,
    profile: {
      nonEmpty: profile.nonEmpty,
      firstType: profile.firstType,
      textRatio: Number(profile.textRatio.toFixed(3)),
      periodRatio: Number(profile.periodRatio.toFixed(3)),
      measureValueRatio: Number(
        profile.measureValueRatio.toFixed(3),
      ),
      dataValueRatio: Number(profile.dataValueRatio.toFixed(3)),
      headerLexicalRatio: Number(
        profile.headerLexicalRatio.toFixed(3),
      ),
      fillRatio: Number(profile.fillRatio.toFixed(3)),
      suspiciousConcatenatedRatio: Number(
        profile.suspiciousConcatenatedRatio.toFixed(3),
      ),
    },
    repeatability: {
      candidateSimilarity: Number(
        repeatability.candidateSimilarity.toFixed(3),
      ),
      followingSimilarity: Number(
        repeatability.followingSimilarity.toFixed(3),
      ),
      supportCount: repeatability.supportCount,
      followingRowCount: repeatability.followingRowCount,
    },
  };
}

function analyzeFlattenedHeaderBand(
  rows = [],
  mergedHeaders = [],
) {
  const profiles = (rows || []).map(buildRowStructuralProfile);
  const adjacentSimilarities = [];

  for (let index = 1; index < rows.length; index += 1) {
    adjacentSimilarities.push(
      rowTypeSignatureSimilarity(rows[index - 1], rows[index]),
    );
  }

  const averageMeasureValueRatio = average(
    profiles.map((profile) => profile.measureValueRatio),
  );
  const averageDataValueRatio = average(
    profiles.map((profile) => profile.dataValueRatio),
  );
  const adjacentTypeSimilarity = average(adjacentSimilarities);
  const mergedValues = (mergedHeaders || []).filter(
    (value) => !isEmpty(value),
  );
  const suspiciousMergedCount = mergedValues.filter(
    looksLikeSuspiciousConcatenatedHeader,
  ).length;
  const suspiciousMergedRatio = mergedValues.length
    ? suspiciousMergedCount / mergedValues.length
    : 0;

  const rejectAsDataBand = Boolean(
    rows.length > 1 &&
      averageMeasureValueRatio >= 0.2 &&
      adjacentTypeSimilarity >= 0.68 &&
      suspiciousMergedRatio >= 0.2,
  );

  return {
    version: STRUCTURAL_HEADER_DATA_CLASSIFIER_VERSION,
    depth: rows.length,
    adjacentTypeSimilarity: Number(
      adjacentTypeSimilarity.toFixed(3),
    ),
    averageMeasureValueRatio: Number(
      averageMeasureValueRatio.toFixed(3),
    ),
    averageDataValueRatio: Number(
      averageDataValueRatio.toFixed(3),
    ),
    suspiciousMergedCount,
    suspiciousMergedRatio: Number(
      suspiciousMergedRatio.toFixed(3),
    ),
    rejectAsDataBand,
  };
}

module.exports = {
  STRUCTURAL_HEADER_DATA_CLASSIFIER_VERSION,
  classifyCellStructuralType,
  buildRowStructuralProfile,
  rowTypeSignatureSimilarity,
  scoreSubsequentRowRepeatability,
  analyzeHeaderDataStructure,
  analyzeFlattenedHeaderBand,
  looksLikeSuspiciousConcatenatedHeader,
};
