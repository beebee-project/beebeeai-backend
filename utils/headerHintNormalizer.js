function stripAggregateTerms(hint = "") {
  return String(hint || "")
    .replace(/평균|합계|총합|최고|최저|최대|최소|중앙값|개수|건수|수량/g, "")
    .replace(/average|avg|sum|total|max|min|median|count/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

module.exports = {
  stripAggregateTerms,
};
