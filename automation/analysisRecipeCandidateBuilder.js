const {
  ANALYSIS_RECIPE_TYPES,
  ANALYSIS_RECIPE_DEFS,
  ANALYSIS_RECIPE_OPTIONS,
} = require("./config/analysisRecipeConfig");

const BUILDER_VERSION =
  "general_analysis_recipe_candidates_v2_2_core_count_preservation";
const MEASURE_ISOLATION_VERSION = "multi_measure_candidate_isolation_v1";

const COUNT_ONLY_RECIPE_TYPES = new Set([
  ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT,
  ANALYSIS_RECIPE_TYPES.GROUP_COUNT,
  ANALYSIS_RECIPE_TYPES.STATUS_COUNT,
  ANALYSIS_RECIPE_TYPES.TIME_COUNT,
  ANALYSIS_RECIPE_TYPES.CROSS_COUNT,
  "category_count",
]);

function normalizeText(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function getColumnHeader(column = {}, fallback = "") {
  const safeColumn = column || {};
  return String(
    safeColumn.header ||
      safeColumn.originalHeader ||
      safeColumn.name ||
      safeColumn.key ||
      safeColumn.accessor ||
      fallback ||
      "",
  ).trim();
}

function getProfile(column = {}) {
  return column.profile || column.quality || {};
}

function getUniqueCount(column = {}) {
  const profile = getProfile(column);
  const value = column.uniqueCount ?? profile.uniqueCount;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function getUniqueRatio(column = {}) {
  const profile = getProfile(column);
  const value = column.uniqueRatio ?? profile.uniqueRatio;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function getNonEmptyCount(column = {}) {
  const profile = getProfile(column);
  const value = column.nonEmptyCount ?? profile.nonEmptyCount;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function headerHasAny(column = {}, hints = []) {
  const header = normalizeText(getColumnHeader(column));
  if (!header) return false;
  return hints.some((hint) => {
    const normalizedHint = normalizeText(hint);
    return (
      normalizedHint &&
      (header.includes(normalizedHint) || normalizedHint.includes(header))
    );
  });
}

function isNameLabelColumn(column = {}) {
  const header = normalizeText(getColumnHeader(column));
  const role = String(column.role || "").toLowerCase();
  const type = String(column.type || "").toLowerCase();
  if (!header || role === "metric") return false;
  if (
    ["number", "numeric", "integer", "float", "currency", "rate"].includes(type)
  ) {
    return false;
  }

  // Patch 24.2: `metricHeaderHints` contains "명" for headcount-style metrics.
  // This must not make label columns such as 지표명/항목명/사업명 numeric metrics.
  const nameLabelPatterns = [
    "지표명",
    "항목명",
    "분류명",
    "자료명",
    "사업명",
    "과제명",
    "기관명",
    "전문기관명",
    "연구기관",
    "상품명",
    "제품명",
    "품목명",
    "성명",
    "이름",
    "명칭",
  ];
  return nameLabelPatterns.some((pattern) =>
    header.includes(normalizeText(pattern)),
  );
}

function isVirtualTable(table = {}) {
  return Boolean(
    table.isVirtual ||
    table.transformation ||
    /#(WIDE_LONG|CROSS_LONG)$/i.test(String(table.tableId || "")),
  );
}

function uniqueTextValues(values = []) {
  const seen = new Set();
  const out = [];

  for (const value of values) {
    const text = String(value || "").trim();
    const key = normalizeText(text);
    if (!text || !key || seen.has(key)) continue;
    seen.add(key);
    out.push(text);
  }

  return out;
}

function getCrossLongMeasureIsolation(table = {}) {
  const transformation = table.transformation || {};
  if (transformation.type !== "crossTableToLong") return null;

  const outputHeaders = transformation.outputHeaders || {};
  const declared = transformation.measureIsolation || {};
  const identityHeader = String(
    declared.identityHeader ||
      outputHeaders.metricIdentity ||
      outputHeaders.crossAxis ||
      "",
  ).trim();
  const metricNameHeader = String(
    declared.metricNameHeader || outputHeaders.metricName || "",
  ).trim();
  const metricValueHeader = String(
    declared.metricValueHeader || outputHeaders.metricValue || "",
  ).trim();
  const declaredMeasures = Array.isArray(declared.measures)
    ? declared.measures
    : [];
  const transformedMeasures = Array.isArray(transformation.crossMetricColumns)
    ? transformation.crossMetricColumns.map((item) =>
        typeof item === "string" ? item : item?.header,
      )
    : [];
  const rowMeasures = Array.isArray(table.rows)
    ? table.rows.map((row) => row?.[identityHeader])
    : [];
  const measures = uniqueTextValues([
    ...declaredMeasures,
    ...transformedMeasures,
    ...rowMeasures,
  ]);

  if (!identityHeader || !metricValueHeader || measures.length <= 1) {
    return null;
  }

  return {
    version: MEASURE_ISOLATION_VERSION,
    identityHeader,
    metricNameHeader,
    metricValueHeader,
    measures,
  };
}

function filterSignature(filters = []) {
  return (Array.isArray(filters) ? filters : [])
    .map((filter) =>
      [
        filter?.header || filter?.column || "",
        filter?.operator || "equals",
        Array.isArray(filter?.value)
          ? filter.value.join("~")
          : String(filter?.value ?? ""),
      ].join(":"),
    )
    .join("|");
}

function isMetricBearingCandidate(candidate = {}, isolation = null) {
  if (!isolation) return false;
  const recipeType =
    candidate.recipeType || candidate.recipeId || candidate.type || "";
  if (COUNT_ONLY_RECIPE_TYPES.has(recipeType)) return false;

  return (
    normalizeText(candidate.columns?.metric || candidate.metricHeader || "") ===
    normalizeText(isolation.metricValueHeader)
  );
}

function isMeasureIdentityDimension(header = "", isolation = null) {
  if (!header || !isolation) return false;
  const normalized = normalizeText(header);
  return [isolation.identityHeader, isolation.metricNameHeader]
    .filter(Boolean)
    .some((value) => normalizeText(value) === normalized);
}

function replaceMetricDisplayText(text = "", sourceHeader = "", label = "") {
  const value = String(text || "");
  const source = String(sourceHeader || "").trim();
  const display = String(label || "").trim();
  if (!display) return value;
  if (source && value.includes(source))
    return value.split(source).join(display);
  return `${display} · ${value}`;
}

function expandMeasureIsolatedCandidate(table = {}, candidate = {}) {
  const isolation = getCrossLongMeasureIsolation(table);
  if (!isMetricBearingCandidate(candidate, isolation)) return [candidate];

  // 동일한 metric identity 컬럼으로 다시 그룹화하는 후보는 필터 후 1행짜리
  // 자명한 결과만 만들므로 제거한다. 실제 분석 dimension 후보만 유지한다.
  if (
    isMeasureIdentityDimension(candidate.columns?.dimension, isolation) ||
    isMeasureIdentityDimension(candidate.columns?.dimension2, isolation)
  ) {
    return [];
  }

  return isolation.measures.map((measure) => ({
    ...candidate,
    id: `${candidate.id}_${stablePart(measure)}`,
    title: replaceMetricDisplayText(
      candidate.title,
      isolation.metricValueHeader,
      measure,
    ),
    description: replaceMetricDisplayText(
      candidate.description,
      isolation.metricValueHeader,
      measure,
    ),
    filters: [
      ...(Array.isArray(candidate.filters) ? candidate.filters : []),
      {
        header: isolation.identityHeader,
        operator: "equals",
        value: measure,
      },
    ],
    metricDisplayHeader: measure,
    measureIsolation: {
      ...isolation,
      selectedMeasure: measure,
    },
    reasonCodes: [
      ...(candidate.reasonCodes || []),
      "MULTI_MEASURE_ISOLATED",
      MEASURE_ISOLATION_VERSION,
    ],
  }));
}

function isAnalysisEligibleTable(table = {}) {
  if (table.tableUsage?.analysisEligible === false) return false;
  if (table.diagnostics?.tableUsage?.analysisEligible === false) return false;
  return true;
}

function columnScore(column = {}, kind = "dimension") {
  const role = String(column.role || "").toLowerCase();
  const type = String(column.type || "").toLowerCase();
  const uniqueRatio = getUniqueRatio(column);
  const uniqueCount = getUniqueCount(column);
  const nonEmptyCount = getNonEmptyCount(column);
  let score = 0.5;

  if (kind === "metric") {
    if (role === "metric") score += 0.35;
    if (
      ["number", "numeric", "integer", "float", "currency", "rate"].includes(
        type,
      )
    )
      score += 0.25;
    if (headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.metricHeaderHints))
      score += 0.2;
    if (uniqueRatio != null && uniqueRatio > 0.15) score += 0.05;
  } else if (kind === "date") {
    if (role === "date") score += 0.4;
    if (["date", "datetime", "period", "year", "month"].includes(type))
      score += 0.25;
    if (headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.dateHeaderHints))
      score += 0.25;
  } else if (kind === "status") {
    if (role === "status") score += 0.35;
    if (headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.statusHeaderHints))
      score += 0.15;
    if (uniqueCount != null && uniqueCount >= 2 && uniqueCount <= 20)
      score += 0.15;
    if (uniqueRatio != null && uniqueRatio <= 0.5) score += 0.08;
  } else {
    if (role === "dimension") score += 0.35;
    if (["category", "text", "string"].includes(type)) score += 0.15;
    if (
      uniqueCount != null &&
      uniqueCount >= 2 &&
      uniqueCount <= Math.max(20, nonEmptyCount * 0.7)
    )
      score += 0.18;
    if (uniqueRatio != null && uniqueRatio <= 0.8) score += 0.08;
    if (headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.idHeaderHints))
      score -= 0.35;
  }

  return Math.max(0, Math.min(1, score));
}

function isMetricColumn(column = {}) {
  const role = String(column.role || "").toLowerCase();
  const type = String(column.type || "").toLowerCase();
  if (["date", "id"].includes(role)) return false;
  if (isNameLabelColumn(column)) return false;
  if (role === "metric") return true;
  if (
    ["number", "numeric", "integer", "float", "currency", "rate"].includes(type)
  )
    return true;
  return (
    headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.metricHeaderHints) &&
    !["text", "string", "category"].includes(type)
  );
}

function isDateColumn(column = {}) {
  const role = String(column.role || "").toLowerCase();
  const type = String(column.type || "").toLowerCase();
  const header = normalizeText(getColumnHeader(column));

  // `전월대비증감률`, `월간달성률`처럼 월 토큰을 포함한 숫자 지표를
  // 기간 dimension으로 오인하지 않는다. 명시적 metric/rate 신호가 우선한다.
  if (role === "metric") return false;
  if (
    ["number", "numeric", "integer", "float", "currency", "rate"].includes(
      type,
    ) &&
    /(증감률|증가율|감소율|달성률|비율|점유율|마진율|rate|ratio|percent)/.test(
      header,
    )
  ) {
    return false;
  }
  if (role === "date") return true;
  if (["date", "datetime", "period", "year", "month"].includes(type))
    return true;
  return headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.dateHeaderHints);
}

function isIdLikeColumn(column = {}) {
  const role = String(column.role || "").toLowerCase();
  if (role === "id") return true;
  const uniqueRatio = getUniqueRatio(column);
  return (
    headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.idHeaderHints) &&
    (uniqueRatio == null || uniqueRatio > 0.7)
  );
}

function isDimensionColumn(column = {}) {
  const role = String(column.role || "").toLowerCase();
  const type = String(column.type || "").toLowerCase();
  if (["metric", "date"].includes(role)) return false;
  if (isDateColumn(column) || isMetricColumn(column)) return false;
  if (role === "dimension" || role === "status") return true;
  if (ANALYSIS_RECIPE_OPTIONS.categoryTypesForDimensionFallback.includes(type))
    return true;
  return false;
}

function isCompositionDimensionColumn(column = {}) {
  if (!column || !getColumnHeader(column)) return false;
  if (isIdLikeColumn(column) || isMetricColumn(column)) return false;

  const role = String(column.role || "").toLowerCase();
  const type = String(column.type || "").toLowerCase();
  const uniqueCount = getUniqueCount(column);
  const uniqueRatio = getUniqueRatio(column);
  const nonEmptyCount = getNonEmptyCount(column);

  if (uniqueCount != null && uniqueCount < 2) return false;
  if (uniqueCount != null && uniqueCount > Math.max(50, nonEmptyCount * 0.65)) {
    return false;
  }
  if (uniqueRatio != null && uniqueRatio > 0.75) return false;

  // Patch 24.1: 구성비는 일반 category뿐 아니라 연도/월처럼
  // low-cardinality인 기간성 컬럼도 그룹 기준으로 쓸 수 있어야 한다.
  if (role === "dimension" || role === "status" || role === "date") return true;
  if (
    ["category", "text", "string", "date", "period", "year", "month"].includes(
      type,
    )
  ) {
    return true;
  }
  return isDateColumn(column) || isDimensionColumn(column);
}

function uniqueColumns(columns = []) {
  const seen = new Set();
  return columns.filter((column, index) => {
    const key = normalizeText(
      column.canonicalKey ||
        column.key ||
        getColumnHeader(column, `col_${index}`),
    );
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function sortColumns(columns = [], kind = "dimension") {
  return uniqueColumns(columns)
    .map((column) => ({ column, score: columnScore(column, kind) }))
    .sort(
      (a, b) =>
        b.score - a.score ||
        getColumnHeader(a.column).localeCompare(
          getColumnHeader(b.column),
          "ko",
        ),
    )
    .map((item) => item.column);
}

function classifyColumns(columns = []) {
  const safeColumns = Array.isArray(columns) ? columns : [];
  const metrics = sortColumns(
    safeColumns.filter(isMetricColumn),
    "metric",
  ).slice(0, ANALYSIS_RECIPE_OPTIONS.maxMetricsPerTable);
  const dates = sortColumns(safeColumns.filter(isDateColumn), "date").slice(
    0,
    ANALYSIS_RECIPE_OPTIONS.maxDatesPerTable,
  );
  const statuses = sortColumns(
    safeColumns.filter(
      (column) =>
        String(column.role || "").toLowerCase() === "status" ||
        headerHasAny(column, ANALYSIS_RECIPE_OPTIONS.statusHeaderHints),
    ),
    "status",
  );
  const dimensions = sortColumns(
    safeColumns.filter(
      (column) => isDimensionColumn(column) && !isIdLikeColumn(column),
    ),
    "dimension",
  ).slice(0, ANALYSIS_RECIPE_OPTIONS.maxDimensionsPerTable);
  const labelDimensions = sortColumns(
    [
      ...dimensions,
      ...safeColumns.filter(
        (column) => !isDateColumn(column) && !isMetricColumn(column),
      ),
    ],
    "dimension",
  ).slice(0, ANALYSIS_RECIPE_OPTIONS.maxDimensionsPerTable + 2);

  return {
    metrics,
    dimensions,
    dates,
    statuses,
    labelDimensions,
    primaryMetric: metrics[0] || null,
    primaryDimension:
      dimensions[0] || statuses[0] || labelDimensions[0] || null,
    primaryDate: dates[0] || null,
  };
}

function stablePart(value = "") {
  return normalizeText(value).slice(0, 48) || "x";
}

function makeCandidate({
  table,
  recipeType,
  title,
  description,
  metric = null,
  dimension = null,
  dimension2 = null,
  date = null,
  operation = "",
  categoryId = "",
  priority = 0,
  confidence = null,
  reasonCodes = [],
}) {
  const idParts = [
    table.tableId,
    recipeType,
    getColumnHeader(dimension),
    getColumnHeader(dimension2),
    getColumnHeader(date),
    getColumnHeader(metric),
  ].filter(Boolean);
  const tableConfidence = Number.isFinite(Number(table.confidence))
    ? Number(table.confidence)
    : 0.7;
  const metricScore = metric ? columnScore(metric, "metric") : 0.7;
  const dimensionScore = dimension ? columnScore(dimension, "dimension") : 0.7;
  const dateScore = date ? columnScore(date, "date") : 0.7;
  const resolvedConfidence =
    confidence == null
      ? Math.min(
          1,
          tableConfidence * 0.45 +
            metricScore * 0.2 +
            dimensionScore * 0.18 +
            dateScore * 0.17,
        )
      : confidence;

  return {
    id: idParts.map(stablePart).join("_"),
    title,
    description,
    recipeType,
    recipeId: recipeType,
    tableId: table.tableId,
    sheetName: table.sheetName,
    sourceTableId: table.sourceTableId || table.tableId,
    confidence: Number(resolvedConfidence.toFixed(4)),
    priority,
    operation,
    categoryId,
    tableUsage: table.tableUsage || null,
    transformation: table.transformation || null,
    isVirtual: Boolean(table.isVirtual),
    columns: {
      metric: getColumnHeader(metric) || null,
      dimension: getColumnHeader(dimension) || null,
      dimension2: getColumnHeader(dimension2) || null,
      date: getColumnHeader(date) || null,
    },
    metricHeader: getColumnHeader(metric) || "",
    groupHeader: getColumnHeader(dimension) || "",
    dimension2Header: getColumnHeader(dimension2) || "",
    dateHeader: getColumnHeader(date) || "",
    recommendationReason: reasonCodes.length
      ? reasonCodes.join(", ")
      : `${title} 후보를 생성했습니다.`,
    reasonCodes: [BUILDER_VERSION, ...reasonCodes].filter(Boolean),
  };
}

function findRecipeDef(recipeType = "") {
  return (
    ANALYSIS_RECIPE_DEFS.find((def) => def.recipeType === recipeType) || null
  );
}

function buildFromDef({
  table,
  def,
  metric = null,
  dimension = null,
  dimension2 = null,
  date = null,
  reasonCodes = [],
}) {
  if (!def) return null;
  if (def.virtualOnly && !isVirtualTable(table)) return null;
  const context = { table, metric, dimension, dimension2, date };
  if ((def.requires || []).some((key) => !context[key])) return null;

  return makeCandidate({
    table,
    recipeType: def.recipeType,
    title: def.titleTemplate(context),
    description: def.descriptionTemplate(context),
    metric,
    dimension,
    dimension2,
    date,
    operation: def.operation,
    categoryId: def.categoryId,
    priority: def.priority,
    reasonCodes: [
      "GENERAL_RECIPE",
      ...(def.virtualOnly ? ["VIRTUAL_TABLE_RECIPE"] : []),
      ...reasonCodes,
    ],
  });
}

function pushCandidate(candidates, candidate) {
  if (!candidate) return;
  const key = [
    candidate.recipeType,
    candidate.tableId,
    candidate.columns?.dimension,
    candidate.columns?.dimension2,
    candidate.columns?.date,
    candidate.columns?.metric,
    filterSignature(candidate.filters),
  ].join("|");
  if (candidates.some((item) => item.__dedupeKey === key)) return;
  candidates.push({ ...candidate, __dedupeKey: key });
}

function getDimensionPairs(dimensions = []) {
  const pairs = [];
  for (let i = 0; i < dimensions.length; i += 1) {
    for (let j = i + 1; j < dimensions.length; j += 1) {
      pairs.push([dimensions[i], dimensions[j]]);
    }
  }
  return pairs.slice(0, ANALYSIS_RECIPE_OPTIONS.maxDimensionPairsPerTable);
}

function candidateDedupeKey(candidate = {}) {
  return [
    candidate.recipeType,
    candidate.tableId,
    candidate.columns?.dimension,
    candidate.columns?.dimension2,
    candidate.columns?.date,
    candidate.columns?.metric,
    filterSignature(candidate.filters),
  ].join("|");
}

function ensureRecipeTypesWithinLimit(
  sorted = [],
  allCandidates = [],
  recipeTypes = [],
  limit = 24,
) {
  const requiredRecipeTypes = Array.from(new Set(recipeTypes.filter(Boolean)));
  const result = sorted.slice(0, limit);

  for (const recipeType of requiredRecipeTypes) {
    if (result.some((candidate) => candidate.recipeType === recipeType)) {
      continue;
    }

    const rescue = allCandidates
      .filter((candidate) => candidate.recipeType === recipeType)
      .sort(
        (a, b) => b.priority - a.priority || b.confidence - a.confidence,
      )[0];

    if (!rescue) continue;

    const rescueKey = candidateDedupeKey(rescue);
    if (
      result.some((candidate) => candidateDedupeKey(candidate) === rescueKey)
    ) {
      continue;
    }

    if (result.length < limit) {
      result.push(rescue);
      continue;
    }

    // Patch 24.2: when multiple required recipe types are rescued
    // (e.g. composition_ratio + top_bottom), do not let a later rescue
    // replace an earlier required one.
    let replaceIndex = -1;
    let replaceScore = Infinity;
    result.forEach((candidate, index) => {
      if (requiredRecipeTypes.includes(candidate.recipeType)) return;
      const score =
        Number(candidate.priority || 0) + Number(candidate.confidence || 0);
      if (score < replaceScore) {
        replaceScore = score;
        replaceIndex = index;
      }
    });

    if (replaceIndex < 0) replaceIndex = result.length - 1;
    result[replaceIndex] = rescue;
  }

  return result.sort(
    (a, b) => b.priority - a.priority || b.confidence - a.confidence,
  );
}

function buildTableCandidates(table = {}) {
  const candidates = [];

  if (!isAnalysisEligibleTable(table)) return candidates;
  if (
    Number(table.confidence ?? 1) < ANALYSIS_RECIPE_OPTIONS.minTableConfidence
  ) {
    return candidates;
  }

  const columns = Array.isArray(table.columns) ? table.columns : [];
  const classified = classifyColumns(columns);
  const {
    metrics,
    dimensions,
    dates,
    statuses,
    labelDimensions,
    primaryMetric,
    primaryDimension,
    primaryDate,
  } = classified;
  const compositionDimensions = sortColumns(
    [...dimensions, ...statuses, ...dates, ...labelDimensions].filter(
      isCompositionDimensionColumn,
    ),
    "dimension",
  ).slice(0, Math.max(3, ANALYSIS_RECIPE_OPTIONS.maxDimensionsPerTable));
  const rankingDimensions = sortColumns(
    [...dimensions, ...statuses, ...labelDimensions].filter(
      (column) =>
        column &&
        getColumnHeader(column) &&
        !isIdLikeColumn(column) &&
        !isDateColumn(column) &&
        !isMetricColumn(column),
    ),
    "dimension",
  ).slice(0, Math.max(3, ANALYSIS_RECIPE_OPTIONS.maxDimensionsPerTable));

  console.log("[recipe-column-roles:v2]", {
    tableId: table.tableId,
    metrics: metrics.map(getColumnHeader),
    dimensions: dimensions.map(getColumnHeader),
    dates: dates.map(getColumnHeader),
    statuses: statuses.map(getColumnHeader),
    compositionDimensions: compositionDimensions.map(getColumnHeader),
    rankingDimensions: rankingDimensions.map(getColumnHeader),
    virtual: isVirtualTable(table),
  });

  // Backward-compatible primary recipe types for existing business template matching.
  pushCandidate(
    candidates,
    buildFromDef({
      table,
      def: findRecipeDef(ANALYSIS_RECIPE_TYPES.GROUP_SUMMARY),
      metric: primaryMetric,
      dimension: primaryDimension,
      reasonCodes: ["PRIMARY_GROUP_SUMMARY"],
    }),
  );
  pushCandidate(
    candidates,
    buildFromDef({
      table,
      def: findRecipeDef(ANALYSIS_RECIPE_TYPES.TIME_TREND),
      metric: primaryMetric,
      date: primaryDate,
      reasonCodes: ["PRIMARY_TIME_TREND"],
    }),
  );
  pushCandidate(
    candidates,
    buildFromDef({
      table,
      def: findRecipeDef(ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT),
      dimension: primaryDimension,
      reasonCodes: ["PRIMARY_CATEGORY_COUNT"],
    }),
  );
  pushCandidate(
    candidates,
    buildFromDef({
      table,
      def: findRecipeDef(ANALYSIS_RECIPE_TYPES.TOP_BOTTOM),
      metric: primaryMetric,
      dimension: labelDimensions[0] || primaryDimension,
      reasonCodes: ["PRIMARY_RANKING"],
    }),
  );

  for (const dimension of dimensions.slice(0, 3)) {
    pushCandidate(
      candidates,
      buildFromDef({
        table,
        def: findRecipeDef(ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT),
        dimension,
        reasonCodes: ["DIMENSION_COUNT"],
      }),
    );

    for (const metric of metrics.slice(0, 2)) {
      for (const recipeType of [
        ANALYSIS_RECIPE_TYPES.GROUP_SUM,
        ANALYSIS_RECIPE_TYPES.GROUP_AVG,
        ANALYSIS_RECIPE_TYPES.COMPOSITION_RATIO,
      ]) {
        pushCandidate(
          candidates,
          buildFromDef({
            table,
            def: findRecipeDef(recipeType),
            metric,
            dimension,
            reasonCodes: ["DIMENSION_METRIC"],
          }),
        );
      }
    }
  }

  // Patch 24.1: 일부 매출/기간형 데이터는 일반 dimension이 부족하거나
  // 구성비 후보가 maxCandidatesPerTable slice에서 밀려날 수 있다.
  // UI/회귀 테스트에서 구성비 후보를 안정적으로 노출하도록 별도 rescue 후보를 만든다.
  for (const dimension of compositionDimensions.slice(0, 4)) {
    for (const metric of metrics.slice(0, 2)) {
      pushCandidate(
        candidates,
        buildFromDef({
          table,
          def: findRecipeDef(ANALYSIS_RECIPE_TYPES.COMPOSITION_RATIO),
          metric,
          dimension,
          reasonCodes: ["COMPOSITION_RATIO_RESCUE", "PATCH_24_1"],
        }),
      );
    }
  }

  // Patch 24.2: WIDE_LONG sales-like tables can generate many trend recipes,
  // causing ranking candidates to be sliced out before regression checks.
  // Keep at least one dimension + metric top/bottom candidate, preferring
  // real label dimensions such as 구/업종 over date/period fields.
  for (const dimension of rankingDimensions.slice(0, 4)) {
    for (const metric of metrics.slice(0, 2)) {
      pushCandidate(
        candidates,
        buildFromDef({
          table,
          def: findRecipeDef(ANALYSIS_RECIPE_TYPES.TOP_BOTTOM),
          metric,
          dimension,
          reasonCodes: ["TOP_BOTTOM_RESCUE", "PATCH_24_2"],
        }),
      );
    }
  }

  for (const status of statuses.slice(0, 2)) {
    pushCandidate(
      candidates,
      makeCandidate({
        table,
        recipeType: ANALYSIS_RECIPE_TYPES.STATUS_COUNT,
        title: `${getColumnHeader(status)}별 상태 현황`,
        description: `${getColumnHeader(status)} 기준 상태 건수를 집계합니다.`,
        dimension: status,
        operation: "count",
        categoryId: "summary",
        priority: 810,
        reasonCodes: ["STATUS_COUNT"],
      }),
    );
  }

  for (const date of dates.slice(0, 2)) {
    pushCandidate(
      candidates,
      buildFromDef({
        table,
        def: findRecipeDef(ANALYSIS_RECIPE_TYPES.TIME_COUNT),
        date,
        reasonCodes: ["DATE_COUNT"],
      }),
    );

    for (const metric of metrics.slice(0, 2)) {
      for (const recipeType of [
        ANALYSIS_RECIPE_TYPES.TIME_SUM,
        ANALYSIS_RECIPE_TYPES.TIME_AVG,
        ANALYSIS_RECIPE_TYPES.TIME_GROWTH,
        ANALYSIS_RECIPE_TYPES.CUMULATIVE_SUM,
      ]) {
        pushCandidate(
          candidates,
          buildFromDef({
            table,
            def: findRecipeDef(recipeType),
            metric,
            date,
            reasonCodes: ["DATE_METRIC"],
          }),
        );
      }

      pushCandidate(
        candidates,
        buildFromDef({
          table,
          def: findRecipeDef(ANALYSIS_RECIPE_TYPES.WIDE_TIME_TREND),
          metric,
          date,
          reasonCodes: ["WIDE_OR_CROSS_LONG"],
        }),
      );
    }
  }

  for (const [dimension, dimension2] of getDimensionPairs(dimensions)) {
    pushCandidate(
      candidates,
      buildFromDef({
        table,
        def: findRecipeDef(ANALYSIS_RECIPE_TYPES.CROSS_COUNT),
        dimension,
        dimension2,
        reasonCodes: ["DIMENSION_PAIR"],
      }),
    );

    if (primaryMetric) {
      pushCandidate(
        candidates,
        buildFromDef({
          table,
          def: findRecipeDef(ANALYSIS_RECIPE_TYPES.CROSS_SUM),
          metric: primaryMetric,
          dimension,
          dimension2,
          reasonCodes: ["DIMENSION_PAIR_METRIC"],
        }),
      );
    }
  }

  const measureIsolation = getCrossLongMeasureIsolation(table);
  const allCandidates = candidates
    .map(({ __dedupeKey, ...candidate }) => candidate)
    .flatMap((candidate) => expandMeasureIsolatedCandidate(table, candidate))
    .sort((a, b) => b.priority - a.priority || b.confidence - a.confidence);
  const requiredRecipeTypes = [];

  // 후보 상한 적용 전 생성된 기본 구조 분석이 후순위 구제 후보에 밀리지
  // 않도록, 실제 구조 증거가 있는 count recipe를 일반 규칙으로 보존한다.
  if (primaryDimension) {
    requiredRecipeTypes.push(ANALYSIS_RECIPE_TYPES.CATEGORY_COUNT);
  }
  if (statuses.length) {
    requiredRecipeTypes.push(ANALYSIS_RECIPE_TYPES.STATUS_COUNT);
  }

  requiredRecipeTypes.push(
    ANALYSIS_RECIPE_TYPES.COMPOSITION_RATIO,
    ANALYSIS_RECIPE_TYPES.TOP_BOTTOM,
  );
  if (primaryDate && primaryMetric) {
    requiredRecipeTypes.push(ANALYSIS_RECIPE_TYPES.TIME_GROWTH);
  }
  if (measureIsolation) {
    requiredRecipeTypes.push(ANALYSIS_RECIPE_TYPES.TIME_SUM);
  }
  const sorted = ensureRecipeTypesWithinLimit(
    allCandidates,
    allCandidates,
    requiredRecipeTypes,
    ANALYSIS_RECIPE_OPTIONS.maxCandidatesPerTable,
  );

  return sorted;
}

function buildAnalysisRecipeCandidates(normalizedQueryTables = []) {
  if (!Array.isArray(normalizedQueryTables)) return [];

  const candidates = normalizedQueryTables.flatMap(buildTableCandidates);

  console.log(
    "[analysisRecipeCandidates:v2]",
    candidates.map((candidate) => ({
      recipeType: candidate.recipeType,
      title: candidate.title,
      tableId: candidate.tableId,
      sourceTableId: candidate.sourceTableId,
    })),
  );

  return candidates;
}

module.exports = {
  BUILDER_VERSION,
  buildAnalysisRecipeCandidates,
  classifyColumns,
  isCompositionDimensionColumn,
  isNameLabelColumn,
  getCrossLongMeasureIsolation,
  expandMeasureIsolatedCandidate,
};
