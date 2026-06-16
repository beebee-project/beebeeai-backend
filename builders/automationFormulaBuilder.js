function quoteSheetName(sheetName = "") {
  const safe = String(sheetName || "Sheet1").replace(/'/g, "''");
  return `'${safe}'`;
}

function buildColumnRange(sheetName, columnLetter) {
  return `${quoteSheetName(sheetName)}!$${columnLetter}:$${columnLetter}`;
}

function buildAverageIfFormula({ groupRange, criteriaCell, valueRange }) {
  return `IFERROR(AVERAGEIF(${groupRange},${criteriaCell},${valueRange}),"")`;
}

function buildSumIfFormula({ groupRange, criteriaCell, valueRange }) {
  return `IFERROR(SUMIF(${groupRange},${criteriaCell},${valueRange}),"")`;
}

function buildCountIfFormula({ groupRange, criteriaCell }) {
  return `IFERROR(COUNTIF(${groupRange},${criteriaCell}),"")`;
}

function buildGroupAggregateFormula({
  operation = "average",
  sheetName = "원본데이터",
  groupLetter,
  metricLetter,
  criteriaCell,
}) {
  const groupRange = buildColumnRange(sheetName, groupLetter);
  const valueRange = buildColumnRange(sheetName, metricLetter);

  if (operation === "sum") {
    return buildSumIfFormula({
      groupRange,
      criteriaCell,
      valueRange,
    });
  }

  if (operation === "count") {
    return buildCountIfFormula({
      groupRange,
      criteriaCell,
    });
  }

  return buildAverageIfFormula({
    groupRange,
    criteriaCell,
    valueRange,
  });
}

function buildRankValueFormula({ valueRange, rank }) {
  return `IFERROR(LARGE(${valueRange},${rank}),"")`;
}

function buildRankLabelFormula({ labelRange, valueRange, rankValueCell }) {
  return `IFERROR(INDEX(${labelRange},MATCH(${rankValueCell},${valueRange},0)),"")`;
}

function buildRunningSumFormula({ valueCell, previousCell }) {
  return `IFERROR(${previousCell}+${valueCell},"")`;
}

function buildGrowthRateFormula({ currentCell, previousCell }) {
  return `IFERROR((${currentCell}-${previousCell})/${previousCell},"")`;
}

function buildMaxIfFormula({ groupRange, criteriaCell, valueRange }) {
  return `IFERROR(MAXIFS(${valueRange},${groupRange},${criteriaCell}),"")`;
}

function buildCountIfsFormula({ criteriaRange, criteriaCell }) {
  return `IFERROR(COUNTIF(${criteriaRange},${criteriaCell}),"")`;
}

function buildPivotAverageFormula({
  rowRange,
  rowCriteriaCell,
  colRange,
  colCriteriaCell,
  valueRange,
}) {
  return `IFERROR(AVERAGEIFS(${valueRange},${rowRange},${rowCriteriaCell},${colRange},${colCriteriaCell}),"")`;
}

module.exports = {
  quoteSheetName,
  buildColumnRange,
  buildAverageIfFormula,
  buildSumIfFormula,
  buildCountIfFormula,
  buildGroupAggregateFormula,
  buildRankValueFormula,
  buildRankLabelFormula,
  buildRunningSumFormula,
  buildGrowthRateFormula,
  buildMaxIfFormula,
  buildCountIfsFormula,
  buildPivotAverageFormula,
};
