const formulaUtils = require("../utils/formulaUtils");
const { buildConditionMask } = require("../utils/conditionEngine");

const logicalFunctionBuilder = require("./formula/legacyBuilders/logicalFunctions");
const mathStatsFunctionBuilder = require("./formula/legacyBuilders/mathStatsFunctions");
const dateFunctionBuilder = require("./formula/legacyBuilders/dateFunctions");
const referenceFunctionBuilder = require("./formula/legacyBuilders/referenceFunctions");
const textFunctionBuilder = require("./formula/legacyBuilders/textFunctions");
const arrayFunctionBuilder = require("./formula/legacyBuilders/arrayFunctions");

function createBaseFormulaBuilder() {
  return {
    _formatValue: (val, opts = {}) =>
      formulaUtils.formatValue(val, { ...opts }),

    _buildConditionPairs: function (ctx) {
      const { intent, allSheetsData } = ctx;
      if (!allSheetsData) return [];
      if (!intent?.conditions?.length) return [];

      return intent.conditions
        .map((c) => {
          let headerText = "";

          if (typeof c?.target === "string") {
            headerText = c.target;
          } else if (c?.target && typeof c.target === "object") {
            headerText = c.target.header || "";
          } else if (c?.hint) {
            headerText = c.hint;
          }

          if (!headerText) return null;

          const term = formulaUtils.expandTermsFromText(headerText);
          const best = formulaUtils.findBestColumnAcrossSheets(
            allSheetsData,
            term,
            "lookup",
          );
          if (!best) return null;

          if (best.isAmbiguous) {
            const candA = best.header || "후보1";
            const candB = best.runnerUpHeader || "후보2";
            ctx.__errorFormula = `=ERROR("조건 열이 모호합니다: '${candA}' 또는 '${candB}' 중 선택이 필요합니다.")`;
            return null;
          }

          const range = `'${best.sheetName}'!${best.columnLetter}${best.startRow}:${best.columnLetter}${best.lastDataRow}`;
          const op = String(c.operator || "=").trim();
          const rawVal = c.value;

          if (
            rawVal == null ||
            (typeof rawVal === "string" && rawVal.trim() === "")
          ) {
            return null;
          }

          const val = this._formatValue(rawVal);
          const cmpOps = new Set([">", ">=", "<", "<=", "<>"]);

          if (cmpOps.has(op)) {
            if (rawVal != null && !isNaN(rawVal)) {
              return `${range}, "${op}${rawVal}"`;
            }
            return `${range}, "${op}"&${val}`;
          }
          if (/^contains$/i.test(op)) return `${range}, "*"&${val}&"*"`;
          if (/^starts?_with$/i.test(op)) return `${range}, ${val}&"*"`;
          if (/^ends?_with$/i.test(op)) return `${range}, "*"&${val}`;

          return `${range}, ${val}`;
        })
        .filter(Boolean);
    },

    _buildConditionMask: function (ctx) {
      return buildConditionMask(ctx, (v, o) =>
        formulaUtils.formatValue(v, {
          ...(ctx?.formatOptions || {}),
          ...(o || {}),
        }),
      );
    },
  };
}

function createFormulaBuilder() {
  const formulaBuilder = createBaseFormulaBuilder();

  Object.assign(formulaBuilder, logicalFunctionBuilder);
  Object.assign(formulaBuilder, mathStatsFunctionBuilder);
  Object.assign(formulaBuilder, dateFunctionBuilder);
  Object.assign(formulaBuilder, referenceFunctionBuilder);
  Object.assign(formulaBuilder, textFunctionBuilder);
  Object.assign(formulaBuilder, arrayFunctionBuilder);

  return formulaBuilder;
}

module.exports = {
  createFormulaBuilder,
};
