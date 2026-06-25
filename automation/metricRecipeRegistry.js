const {
  RATE_RECIPES,
  COMPARE_RECIPES,
} = require("./config/metricRecipeConfig");

function normalizeText(v = "") {
  return String(v)
    .toLowerCase()
    .replace(/\([^)]*\)/g, "")
    .replace(/[^가-힣a-z0-9]/gi, "")
    .trim();
}

function findMetricRecipe(message = "") {
  const s = normalizeText(message);

  return (
    RATE_RECIPES.find((recipe) =>
      recipe.match.some((kw) => s.includes(normalizeText(kw))),
    ) || null
  );
}

function findCompareRecipe(message = "") {
  const s = normalizeText(message);

  return (
    COMPARE_RECIPES.find((recipe) =>
      recipe.match.some((kw) => s.includes(normalizeText(kw))),
    ) || null
  );
}

module.exports = {
  RATE_RECIPES,
  findMetricRecipe,
  normalizeText,
  COMPARE_RECIPES,
  findCompareRecipe,
};
