function buildGroupFormula({
  groupRange,
  valueRange,
  aggregate = "count",
  condition = null,
  sort = false,
  sortOrder = -1,
}) {
  const aggMap = {
    count: `COUNTIFS(${groupRange},k${condition ? "," + condition : ""})`,
    avg: `AVERAGEIFS(${valueRange},${groupRange},k)`,
    sum: `SUMIFS(${valueRange},${groupRange},k)`,
    max: `MAXIFS(${valueRange},${groupRange},k)`,
    min: `MINIFS(${valueRange},${groupRange},k)`,
    median: `MEDIAN(FILTER(${valueRange},${groupRange}=k))`,
  };

  const agg = aggMap[aggregate] || aggMap.count;

  const base = `LET(keys,UNIQUE(${groupRange}),HSTACK(keys,MAP(keys,LAMBDA(k,${agg}))))`;

  if (!sort) {
    return "=" + base;
  }

  return `=SORT(${base},2,${sortOrder})`;
}

module.exports = {
  buildGroupFormula,
};
