const RATE_RECIPES = [
  {
    id: "termination_rate",
    match: ["퇴사율", "이탈률", "종료율", "해지율"],
    outputHeader: "퇴사율",
    numerator: {
      type: "exists",
      columnHints: ["퇴사", "종료", "해지", "이탈"],
    },
    denominator: {
      type: "count",
    },
    multiplier: 100,
  },
  {
    id: "conversion_rate",
    match: ["전환율", "구매전환율", "가입전환율"],
    outputHeader: "전환율",
    numerator: {
      type: "positive",
      columnHints: ["전환", "구매", "가입", "성공", "완료"],
    },
    denominator: {
      type: "count",
    },
    multiplier: 100,
  },
  {
    id: "defect_rate",
    match: ["불량률", "오류율", "실패율"],
    outputHeader: "불량률",
    numerator: {
      type: "positive",
      columnHints: ["불량", "오류", "실패", "결함"],
    },
    denominator: {
      type: "count",
    },
    multiplier: 100,
  },
  {
    id: "completion_rate",
    match: ["완료율", "처리율", "달성률"],
    outputHeader: "완료율",
    numerator: {
      type: "positive",
      columnHints: ["완료", "처리", "달성", "성공"],
    },
    denominator: {
      type: "count",
    },
    multiplier: 100,
  },
  {
    id: "generic_rate",
    match: ["율", "비율"],
    outputHeader: "비율",
    numerator: {
      type: "exists",
      columnHints: [
        "퇴사",
        "종료",
        "해지",
        "완료",
        "전환",
        "불량",
        "성공",
        "실패",
      ],
    },
    denominator: {
      type: "count",
    },
    multiplier: 100,
  },
];

const COMPARE_RECIPES = [
  {
    id: "growth_rate",
    match: [
      "성장률",
      "증감률",
      "증가율",
      "감소율",
      "전년대비",
      "전월대비",
      "전기대비",
    ],
    type: "compare",
    mode: "previous",
    method: "growthRate",
    outputHeader: "증감률",
    multiplier: 100,
    defaultAggregate: "sum",
  },
];

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
