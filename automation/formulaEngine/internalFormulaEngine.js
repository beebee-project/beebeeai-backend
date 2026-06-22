const DEFAULT_SOURCE_SHEET_NAME = "원본데이터";

const FORMULA_SPEC_TYPES = Object.freeze({
  GROUP_AGGREGATE: "groupAggregate",
  AVERAGE_IF: "averageIf",
  SUM_IF: "sumIf",
  COUNT_IF: "countIf",
  RANK_VALUE: "rankValue",
  RANK_LABEL: "rankLabel",
  RUNNING_SUM: "runningSum",
  GROWTH_RATE: "growthRate",
  MAX_IF: "maxIf",
  COUNT_IFS: "countIfs",
  PIVOT_AVERAGE: "pivotAverage",
  ROLLING_AVERAGE: "rollingAverage",
  IFERROR_REF: "ifErrorRef",
  RAW: "raw",
});

function quoteSheetName(sheetName = "") {
  const safe = String(sheetName || DEFAULT_SOURCE_SHEET_NAME).replace(/'/g, "''");
  return `'${safe}'`;
}

function buildColumnRange(sheetName = DEFAULT_SOURCE_SHEET_NAME, columnLetter) {
  if (!columnLetter) return "";
  return `${quoteSheetName(sheetName)}!$${columnLetter}:$${columnLetter}`;
}

function stripLeadingEquals(formula = "") {
  return String(formula || "").replace(/^=/, "");
}

function withIfError(body = "", fallback = '""') {
  return `IFERROR(${body},${fallback})`;
}

function buildAverageIfFormula({ groupRange, criteriaCell, valueRange }) {
  return withIfError(`AVERAGEIF(${groupRange},${criteriaCell},${valueRange})`);
}

function buildSumIfFormula({ groupRange, criteriaCell, valueRange }) {
  return withIfError(`SUMIF(${groupRange},${criteriaCell},${valueRange})`);
}

function buildCountIfFormula({ groupRange, criteriaCell }) {
  return withIfError(`COUNTIF(${groupRange},${criteriaCell})`);
}

function buildGroupAggregateFormula({
  operation = "average",
  sheetName = DEFAULT_SOURCE_SHEET_NAME,
  groupLetter,
  metricLetter,
  criteriaCell,
}) {
  const groupRange = buildColumnRange(sheetName, groupLetter);
  const valueRange = buildColumnRange(sheetName, metricLetter);
  const op = String(operation || "average").toLowerCase();

  if (op === "sum") {
    return buildSumIfFormula({ groupRange, criteriaCell, valueRange });
  }

  if (op === "count") {
    return buildCountIfFormula({ groupRange, criteriaCell });
  }

  return buildAverageIfFormula({ groupRange, criteriaCell, valueRange });
}

function buildRankValueFormula({ valueRange, rank }) {
  return withIfError(`LARGE(${valueRange},${rank})`);
}

function buildRankLabelFormula({ labelRange, valueRange, rankValueCell }) {
  return withIfError(`INDEX(${labelRange},MATCH(${rankValueCell},${valueRange},0))`);
}

function buildRunningSumFormula({ valueCell, previousCell }) {
  return withIfError(`${previousCell}+${valueCell}`);
}

function buildGrowthRateFormula({ currentCell, previousCell }) {
  return withIfError(`(${currentCell}-${previousCell})/${previousCell}`);
}

function buildMaxIfFormula({ groupRange, criteriaCell, valueRange }) {
  return withIfError(`MAXIFS(${valueRange},${groupRange},${criteriaCell})`);
}

function buildCountIfsFormula({ criteriaRange, criteriaCell }) {
  return withIfError(`COUNTIF(${criteriaRange},${criteriaCell})`);
}

function buildPivotAverageFormula({
  rowRange,
  rowCriteriaCell,
  colRange,
  colCriteriaCell,
  valueRange,
}) {
  return withIfError(
    `AVERAGEIFS(${valueRange},${rowRange},${rowCriteriaCell},${colRange},${colCriteriaCell})`,
  );
}

function buildRollingAverageFormula({ startCell, endCell }) {
  return withIfError(`AVERAGE(${startCell}:${endCell})`);
}

function buildIfErrorRefFormula({ cell }) {
  return withIfError(cell);
}

function buildFormulaFromSpec(spec = {}) {
  const type = spec.type || spec.formulaType || FORMULA_SPEC_TYPES.RAW;

  switch (type) {
    case FORMULA_SPEC_TYPES.GROUP_AGGREGATE:
      return buildGroupAggregateFormula(spec);
    case FORMULA_SPEC_TYPES.AVERAGE_IF:
      return buildAverageIfFormula(spec);
    case FORMULA_SPEC_TYPES.SUM_IF:
      return buildSumIfFormula(spec);
    case FORMULA_SPEC_TYPES.COUNT_IF:
      return buildCountIfFormula(spec);
    case FORMULA_SPEC_TYPES.RANK_VALUE:
      return buildRankValueFormula(spec);
    case FORMULA_SPEC_TYPES.RANK_LABEL:
      return buildRankLabelFormula(spec);
    case FORMULA_SPEC_TYPES.RUNNING_SUM:
      return buildRunningSumFormula(spec);
    case FORMULA_SPEC_TYPES.GROWTH_RATE:
      return buildGrowthRateFormula(spec);
    case FORMULA_SPEC_TYPES.MAX_IF:
      return buildMaxIfFormula(spec);
    case FORMULA_SPEC_TYPES.COUNT_IFS:
      return buildCountIfsFormula(spec);
    case FORMULA_SPEC_TYPES.PIVOT_AVERAGE:
      return buildPivotAverageFormula(spec);
    case FORMULA_SPEC_TYPES.ROLLING_AVERAGE:
      return buildRollingAverageFormula(spec);
    case FORMULA_SPEC_TYPES.IFERROR_REF:
      return buildIfErrorRefFormula(spec);
    case FORMULA_SPEC_TYPES.RAW:
    default:
      return stripLeadingEquals(spec.formula || spec.raw || "");
  }
}

function createFormulaCell({ formula, value = 0, cellType = "n" } = {}) {
  return {
    t: cellType,
    f: stripLeadingEquals(formula),
    v: value,
  };
}

function createFormulaCellFromSpec(spec = {}) {
  return createFormulaCell({
    formula: buildFormulaFromSpec(spec),
    value: spec.value,
    cellType: spec.cellType,
  });
}

module.exports = {
  DEFAULT_SOURCE_SHEET_NAME,
  SOURCE_SHEET_NAME: DEFAULT_SOURCE_SHEET_NAME,
  FORMULA_SPEC_TYPES,
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
  buildRollingAverageFormula,
  buildIfErrorRefFormula,
  buildFormulaFromSpec,
  buildFormula: buildFormulaFromSpec,
  createFormulaCell,
  createFormulaCellFromSpec,
};
