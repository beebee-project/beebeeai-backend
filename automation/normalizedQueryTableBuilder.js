const NORMALIZATION_VERSION = "normalized_query_table_common_v5_final_three";
const DIAGNOSTICS_VERSION = "normalized_query_diagnostics_v4";
const WIDE_TO_LONG_VERSION = "wide_to_long_normalization_v5_unit_hierarchy";
const CROSS_TO_LONG_VERSION =
  "cross_table_to_long_normalization_v5_unit_hierarchy";
const ROW_CLASSIFICATION_VERSION = "normalized_row_classification_v3";

const DEFAULT_TABLE_USAGE = Object.freeze({
  version: "table_usage_quality_v1",
  queryable: true,
  analysisEligible: true,
  templateEligible: true,
  reasons: ["TABLE_USAGE_NOT_PROVIDED"],
  metrics: {},
});

const SUMMARY_LABEL_PATTERN =
  /^(?:합계|소계|총계|전체|전국|세계|계|total|subtotal|grand\s*total)$/i;
const SUMMARY_SUFFIX_PATTERN = /(?:합계|소계|총계)\s*$/;
const UNIT_HEADER_PATTERN = /^(?:단위|측정단위|unit|measure\s*unit)$/i;
const METRIC_IDENTITY_HEADER_PATTERN =
  /^(?:항목|지표|지표명|측정항목|세부항목|metric|measure|indicator)$/i;
const IDENTIFIER_HEADER_PATTERN =
  /(?:^|[\s_\-])(id|code)(?:$|[\s_\-])|번호|코드|순번|연번/i;
const DIMENSION_HEADER_PATTERN =
  /구분|분류|유형|종류|항목|명칭|이름|지역|국가|산업|업종|사업|학교|학년|성별|죄종|침입구|특성|대분류|소분류|category|dimension|group|type|name/i;
const NON_ADDITIVE_METRIC_PATTERN =
  /%|퍼센트|백분율|비율|비중|구성비|점유율|증감률|달성률|순이동률|지수|평균|평점|점수|기대수명|수명|시간|기록|속도|성향|분포|rate|ratio|share|percent|index|average|avg|score|life\s*expectancy|duration|time/i;
const NON_ADDITIVE_UNIT_PATTERN =
  /^(?:%|퍼센트|백분율|지수|점|평점|초|분|분\s*,\s*초|분초|시간|일|개월|년|세|cm|mm|m|km\/h|명\s*\/\s*천명)$/i;
const AVERAGE_CONTEXT_PATTERN =
  /월평균|일평균|주평균|분기평균|연평균|기간평균|평균값|가구당|1인당|인당|단위당|평균\s*(?:실적|소득|지출|금액|수량)/i;
const MEASUREMENT_CONTEXT_PATTERN =
  /체력|평가|측정|검사|테스트|기록|달리기|걷기|뛰기|굽히기|악력|유연성|신체|성적/i;
const COUNT_CONTEXT_PATTERN =
  /건수|개수|수량|횟수|인원수|기업수|사업수|기관수|업체수|시설수|count|number\s+of/i;
const COMMON_UNIT_PATTERN =
  /^(?:원|천원|만원|억원|개|건|명|세|년|월|일|회|점|대|곳|식|건수|명\/천명|명\s*\/\s*천명|%|퍼센트|초|분|분\s*,\s*초|cm|mm|m|km|kg|g|톤|t|kwh|mwh|kw|mw|㎡|㎥|원\/\S+|\S+\/\S+)$/i;

function isBlank(value) {
  return value == null || String(value).trim() === "";
}

function asText(value = "") {
  return String(value ?? "");
}

function normalizeWhitespace(value = "") {
  return asText(value).replace(/\s+/g, " ").trim();
}

function normalizeHeader(value = "") {
  return normalizeWhitespace(value);
}

function normalizedHeaderKey(value = "") {
  return normalizeHeader(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>（）]+/g, "")
    .trim();
}

function canonicalKeyFromHeader(value = "", fallback = "") {
  const key = normalizeHeader(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[^\p{Letter}\p{Number}]+/gu, "_")
    .replace(/^_+|_+$/g, "");
  return key || fallback;
}

function cloneValue(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (Array.isArray(value)) return value.map(cloneValue);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, cloneValue(item)]),
    );
  }
  return value;
}

function toNumberOrNull(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;

  const normalized = value
    .normalize("NFKC")
    .replace(/,/g, "")
    .replace(/%$/g, "")
    .trim();

  if (!normalized || normalized === "-") return null;
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return null;

  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function normalizeTableUsage(table = {}) {
  const usage = table.tableUsage || table.usage || null;
  if (!usage || typeof usage !== "object") return { ...DEFAULT_TABLE_USAGE };

  return {
    version: usage.version || DEFAULT_TABLE_USAGE.version,
    queryable: usage.queryable !== false,
    analysisEligible: usage.analysisEligible !== false,
    templateEligible: usage.templateEligible !== false,
    reasons:
      Array.isArray(usage.reasons) && usage.reasons.length
        ? [...usage.reasons]
        : [...DEFAULT_TABLE_USAGE.reasons],
    metrics: cloneValue(usage.metrics || {}),
  };
}

function isAnalysisEligibleTable(table = {}) {
  return normalizeTableUsage(table).analysisEligible !== false;
}

function inheritVirtualTableUsage(sourceTable = {}, transformationType = "") {
  const usage = normalizeTableUsage(sourceTable);
  const reason = `VIRTUAL_TABLE_FROM_${String(
    transformationType || "NORMALIZATION",
  ).toUpperCase()}`;

  return {
    ...usage,
    reasons: [...new Set([...(usage.reasons || []), reason])],
  };
}

function rowValue(row, column = {}, index = 0) {
  if (Array.isArray(row)) return row[index];
  if (!row || typeof row !== "object") return undefined;

  const candidates = [
    column.key,
    column.canonicalKey,
    column.accessor,
    column.name,
    column.header,
    column.originalHeader,
  ].filter(Boolean);

  for (const key of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
  }

  const targets = new Set(candidates.map(normalizedHeaderKey).filter(Boolean));
  for (const [key, value] of Object.entries(row)) {
    if (targets.has(normalizedHeaderKey(key))) return value;
  }

  return Object.values(row)[index];
}

function columnValues(table = {}, column = {}, index = 0) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  return rows.map((row) => rowValue(row, column, column.index ?? index));
}

function analyzeValues(values = []) {
  const nonBlank = values.filter((value) => !isBlank(value));
  const numericCount = nonBlank.filter(
    (value) => toNumberOrNull(value) != null,
  ).length;
  const dateCount = nonBlank.filter(isStrictTemporalValue).length;
  const booleanCount = nonBlank.filter(
    (value) =>
      typeof value === "boolean" ||
      /^(?:true|false|yes|no|y|n|예|아니오)$/i.test(normalizeWhitespace(value)),
  ).length;
  const unique = new Set(nonBlank.map((value) => normalizeWhitespace(value)));
  const total = nonBlank.length;

  return {
    totalRows: values.length,
    nonEmptyCount: total,
    emptyRatio: values.length ? 1 - total / values.length : 1,
    numericRatio: total ? numericCount / total : 0,
    dateRatio: total ? dateCount / total : 0,
    booleanRatio: total ? booleanCount / total : 0,
    uniqueCount: unique.size,
    uniqueRatio: total ? unique.size / total : 0,
    sampleValues: [...unique].slice(0, 10),
  };
}

function extractHeaderUnit(value = "") {
  const header = normalizeHeader(value);
  const matches = [
    ...header.matchAll(/(?:\(|\[|（)\s*([^()[\]（）]{1,24})\s*(?:\)|\]|）)/g),
  ];

  if (!matches.length) return "";
  const candidate = normalizeWhitespace(matches[matches.length - 1][1]);

  if (candidate.length > 16) return "";
  if (COMMON_UNIT_PATTERN.test(candidate)) return candidate;
  return "";
}

function stripUnitSuffix(value = "") {
  const unit = extractHeaderUnit(value);
  if (!unit) return normalizeHeader(value);

  return normalizeHeader(value)
    .replace(
      new RegExp(
        `(?:\\(|\\[|（)\\s*${escapeRegExp(unit)}\\s*(?:\\)|\\]|）)\\s*$`,
      ),
      "",
    )
    .trim();
}

function escapeRegExp(value = "") {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function strictTemporalMatch(value = "") {
  const raw = normalizeHeader(value).normalize("NFKC");
  if (!raw) return null;

  const patterns = [
    {
      type: "date",
      regex:
        /(?:^|[^\d])((?:19|20|21)\d{2})[-./]\s*(0?[1-9]|1[0-2])[-./]\s*(0?[1-9]|[12]\d|3[01])(?:[^\d]|$)/,
      period: (m) =>
        `${m[1]}-${String(Number(m[2])).padStart(2, "0")}-${String(
          Number(m[3]),
        ).padStart(2, "0")}`,
    },
    {
      type: "month",
      regex:
        /(?:^|[^\d])((?:19|20|21)\d{2})\s*년\s*(0?[1-9]|1[0-2])\s*월(?:[^\d]|$)/,
      period: (m) => `${m[1]}-${String(Number(m[2])).padStart(2, "0")}`,
    },
    {
      type: "month",
      regex:
        /(?:^|[^\d])((?:19|20|21)\d{2})[-./]\s*(0?[1-9]|1[0-2])(?:[^\d]|$)/,
      period: (m) => `${m[1]}-${String(Number(m[2])).padStart(2, "0")}`,
    },
    {
      type: "quarter",
      regex:
        /(?:^|[^\d])((?:19|20|21)\d{2})\s*(?:년\s*)?(?:Q\s*([1-4])|([1-4])\s*분기)(?:[^\d]|$)/i,
      period: (m) => `${m[1]}-Q${m[2] || m[3]}`,
    },
    {
      type: "year",
      regex: /(?:^|[^\d])((?:19|20|21)\d{2})\s*년?(?:[^\d]|$)/,
      period: (m) => m[1],
    },
    {
      type: "month",
      regex: /(?:^|[^\d])(0?[1-9]|1[0-2])\s*월(?:[^\d]|$)/,
      period: (m) => `${String(Number(m[1])).padStart(2, "0")}월`,
    },
  ];

  for (const spec of patterns) {
    const match = raw.match(spec.regex);
    if (!match) continue;

    const matchIndex = Number(match.index || 0);
    const before = raw.slice(0, matchIndex);
    const after = raw.slice(matchIndex + match[0].length);
    const metricLabel = normalizeHeader(`${before} ${after}`)
      .replace(/^[\s_|/\\:;,\-–—]+|[\s_|/\\:;,\-–—]+$/g, "")
      .trim();

    return {
      raw,
      type: spec.type,
      period: spec.period(match),
      matchedText: match[0].trim(),
      year: match[1] && /^(?:19|20|21)\d{2}$/.test(match[1]) ? match[1] : "",
      metricLabel,
    };
  }

  return null;
}

function parseTemporalHeader(value = "") {
  const unit = extractHeaderUnit(value);
  const withoutUnit = stripUnitSuffix(value);
  const parsed = strictTemporalMatch(withoutUnit);
  if (!parsed) return null;

  return {
    ...parsed,
    metricLabel: normalizeHeader(parsed.metricLabel) || "지표값",
    unit,
  };
}

function isStrictTemporalValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;
  if (typeof value !== "string" && typeof value !== "number") return false;

  const raw = normalizeWhitespace(value);
  if (!raw) return false;

  return (
    /^(?:19|20|21)\d{2}$/.test(raw) ||
    /^(?:19|20|21)\d{2}[-./](?:0?[1-9]|1[0-2])$/.test(raw) ||
    /^(?:19|20|21)\d{2}\s*년\s*(?:0?[1-9]|1[0-2])\s*월$/.test(raw) ||
    /^(?:19|20|21)\d{2}[-./](?:0?[1-9]|1[0-2])[-./](?:0?[1-9]|[12]\d|3[01])$/.test(
      raw,
    ) ||
    /^(?:19|20|21)\d{2}\s*(?:Q[1-4]|[1-4]\s*분기)$/i.test(raw)
  );
}

function parseTemporalValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = String(value.getFullYear());
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return {
      raw: value,
      type: "date",
      period: `${year}-${month}-${day}`,
      year,
    };
  }

  if (!isStrictTemporalValue(value)) return null;
  return strictTemporalMatch(normalizeWhitespace(value));
}

function collectTextFragments(value, output = [], depth = 0) {
  if (depth > 4 || value == null) return output;

  if (typeof value === "string" || typeof value === "number") {
    const text = normalizeWhitespace(value);
    if (text) output.push(text);
    return output;
  }

  if (Array.isArray(value)) {
    value
      .slice(0, 100)
      .forEach((item) => collectTextFragments(item, output, depth + 1));
    return output;
  }

  if (typeof value === "object") {
    Object.values(value)
      .slice(0, 100)
      .forEach((item) => collectTextFragments(item, output, depth + 1));
  }

  return output;
}

function tableContextText(table = {}) {
  return [
    table.sheetName,
    table.tableName,
    table.title,
    table.name,
    table.source,
    table.description,
    table.caption,
    table.meta,
    table.metadata,
    table.headerRows,
    (table.columns || []).map(
      (column) => column.header || column.originalHeader || "",
    ),
  ]
    .flatMap((value) => collectTextFragments(value))
    .join(" ");
}

function normalizeUnitCandidate(value = "") {
  const text = normalizeHeader(value);
  if (!text) return "";
  if (toNumberOrNull(text) != null) return "";
  if (isStrictTemporalValue(text)) return "";
  if (text.length > 30) return "";
  if (!/[A-Za-z가-힣%‰℃℉₩$€¥㎡㎥㎏㎞㎝㎜\\/·,]/.test(text)) return "";
  const normalized = text.replace(/^[=:：,\\s]+|[=:：,\\s]+$/g, "").trim();
  if (
    !normalized ||
    /^(?:값|수치|지표값|기간|연도|년도|월|분기)$/i.test(normalized)
  )
    return "";
  return normalized;
}

function parseMetricUnitPairs(text = "") {
  const source = normalizeWhitespace(text);
  if (!source || !/[:：]/.test(source)) return [];

  return source
    .split(/\s*\/\s*(?=[^/:：]{1,80}\s*[:：])/)
    .map((part) => {
      const match = part.match(
        /^\s*([^:：]{1,80}?)\s*[:：]\s*([^:：]{1,40})\s*$/,
      );
      if (!match) return null;

      const metric = normalizeHeader(match[1])
        .replace(/^(?:단위|unit)\s*/i, "")
        .trim();
      const unit = normalizeUnitCandidate(match[2]);

      if (!metric || !unit || metric.length > 60 || unit.length > 24) {
        return null;
      }

      return { metric, unit };
    })
    .filter(Boolean);
}

function buildNormalizationContext(tables = []) {
  const unitCandidates = new Map();
  const conflicts = new Set();

  for (const table of tables || []) {
    const fragments = collectTextFragments({
      sheetName: table.sheetName,
      tableName: table.tableName,
      title: table.title,
      name: table.name,
      meta: table.meta,
      metadata: table.metadata,
      headerRows: table.headerRows,
      columns: table.columns,
      rows:
        Array.isArray(table.rows) && table.rows.length <= 20 ? table.rows : [],
    });

    for (const fragment of fragments) {
      for (const pair of parseMetricUnitPairs(fragment)) {
        const key = normalizedHeaderKey(pair.metric);
        if (!key) continue;

        if (
          unitCandidates.has(key) &&
          normalizedHeaderKey(unitCandidates.get(key)) !==
            normalizedHeaderKey(pair.unit)
        ) {
          conflicts.add(key);
          continue;
        }
        unitCandidates.set(key, pair.unit);
      }
    }
  }

  for (const key of conflicts) unitCandidates.delete(key);

  return {
    metricUnitMap: Object.fromEntries(unitCandidates),
  };
}

function contextUnitForMetric(metricLabel = "", context = {}) {
  const key = normalizedHeaderKey(metricLabel);
  if (!key) return "";

  const direct = normalizeUnitCandidate(context.metricUnitMap?.[key]);
  if (direct) return direct;

  const candidates = Object.entries(context.metricUnitMap || {})
    .filter(
      ([candidate]) =>
        candidate.length >= 3 &&
        (key.includes(candidate) || candidate.includes(key)),
    )
    .sort((left, right) => right[0].length - left[0].length);

  return normalizeUnitCandidate(candidates[0]?.[1]);
}

function numericProfileForColumn(rows = [], column = {}) {
  const values = rows
    .map((row) => rowValue(row, column, column.index ?? 0))
    .filter((value) => !isBlank(value));

  const numbers = values.map(toNumberOrNull).filter((value) => value != null);

  const fractionalCount = numbers.filter(
    (value) => !Number.isInteger(value),
  ).length;
  const exactHundredCount = numbers.filter((value) => value === 100).length;

  return {
    nonBlankCount: values.length,
    numericCount: numbers.length,
    numericRatio: values.length ? numbers.length / values.length : 0,
    min: numbers.length ? Math.min(...numbers) : null,
    max: numbers.length ? Math.max(...numbers) : null,
    fractionalCount,
    fractionalRatio: numbers.length ? fractionalCount / numbers.length : 0,
    exactHundredCount,
    integerRatio: numbers.length
      ? numbers.filter(Number.isInteger).length / numbers.length
      : 0,
  };
}

function looksLikePercentageProfile(profile = {}) {
  return (
    profile.numericCount >= 2 &&
    profile.min != null &&
    profile.max != null &&
    profile.min >= 0 &&
    profile.max <= 100 &&
    profile.fractionalCount > 0 &&
    (profile.exactHundredCount > 0 || profile.fractionalRatio >= 0.5)
  );
}

function normalizeDeclaredAggregation(value = "") {
  const text = normalizeWhitespace(value).toLowerCase();
  if (["average", "avg", "mean", "평균"].includes(text)) {
    return "average";
  }
  if (["sum", "total", "합계", "합산"].includes(text)) {
    return "sum";
  }
  return "";
}

function metricCountEntity(text = "") {
  const normalized = normalizeHeader(text);
  const match = normalized.match(/\d[\d,]*\s*대\s*([가-힣A-Za-z]{1,20})/);
  if (!match) return "";
  return `${match[1]}수`;
}

function dimensionEntityName(header = "") {
  return normalizeHeader(header)
    .replace(/(?:\(|\[)?\d+(?:\)|\])?/g, "")
    .replace(/별\s*$/g, "")
    .replace(/정보|구분|분류|유형/g, "")
    .trim();
}

function inferGenericCountMetricLabel({
  metricLabel = "",
  sourceHeader = "",
  dimensions = [],
  profile = {},
} = {}) {
  const label = normalizeHeader(metricLabel);
  const generic =
    !label || /^(?:지표값|값|수치|실적|기간(?:\(|$))/i.test(label);

  if (!generic) return label;

  const countEntity =
    metricCountEntity(sourceHeader) || metricCountEntity(label);
  if (countEntity) return countEntity;

  if (
    profile.numericCount >= 2 &&
    profile.min != null &&
    profile.min >= 0 &&
    profile.integerRatio >= 0.95
  ) {
    const entities = dimensions
      .map((column) =>
        dimensionEntityName(column.header || column.originalHeader || ""),
      )
      .filter(Boolean);

    if (entities.length >= 2 && new Set(entities).size === 1) {
      return `${entities[0]}수`;
    }
  }

  return label || "지표값";
}

function inferMetricUnit({
  metricLabel = "",
  localUnit = "",
  context = {},
  profile = {},
} = {}) {
  const explicit = normalizeUnitCandidate(localUnit);
  if (explicit) return explicit;

  const contextual = normalizeUnitCandidate(
    contextUnitForMetric(metricLabel, context),
  );
  if (contextual) return contextual;

  if (looksLikePercentageProfile(profile)) return "%";
  return "";
}

function inferType(values = [], declared = "") {
  if (declared && declared !== "unknown") return declared;

  const profile = analyzeValues(values);
  if (profile.dateRatio >= 0.7) return "date";
  if (profile.numericRatio >= 0.7) return "number";
  if (profile.booleanRatio >= 0.7) return "boolean";
  if (profile.nonEmptyCount) return "string";
  return "unknown";
}

function pairedDimensionEvidence(columns = [], index = 0) {
  const current = normalizeHeader(columns[index]?.header || "");
  const next = normalizeHeader(columns[index + 1]?.header || "");
  const match = current.match(/^(.*?)(?:\(|\[)?1(?:\)|\])?$/);
  if (!match || !next) return false;

  const base = normalizedHeaderKey(match[1]);
  const nextMatch = next.match(/^(.*?)(?:\(|\[)?2(?:\)|\])?$/);
  return Boolean(nextMatch && normalizedHeaderKey(nextMatch[1]) === base);
}

function inferRole({
  header = "",
  type = "unknown",
  profile = {},
  column = {},
  columns = [],
  index = 0,
} = {}) {
  const declared = String(
    column.role || column.semanticRole || "",
  ).toLowerCase();
  if (declared) return declared;

  const normalized = normalizeHeader(header);
  const temporal = parseTemporalHeader(normalized);

  if (
    column.semanticType === "date" ||
    type === "date" ||
    profile.dateRatio >= 0.7 ||
    temporal
  ) {
    return "date";
  }

  if (UNIT_HEADER_PATTERN.test(normalized)) return "dimension";
  if (METRIC_IDENTITY_HEADER_PATTERN.test(normalized)) return "dimension";
  if (IDENTIFIER_HEADER_PATTERN.test(normalized)) return "id";

  if (
    pairedDimensionEvidence(columns, index) &&
    profile.numericRatio >= 0.5 &&
    profile.uniqueRatio <= 0.5
  ) {
    return "id";
  }

  if (DIMENSION_HEADER_PATTERN.test(normalized)) {
    if (
      type === "number" &&
      profile.uniqueRatio >= 0.95 &&
      profile.nonEmptyCount >= 3 &&
      !pairedDimensionEvidence(columns, index)
    ) {
      return "metric";
    }
    return "dimension";
  }

  if (type === "number" && profile.numericRatio >= 0.7) return "metric";
  if (type === "string" || type === "boolean") return "dimension";
  return "unknown";
}

function normalizeColumn(column = {}, index = 0, table = {}) {
  const header = normalizeHeader(
    column.header ||
      column.name ||
      column.key ||
      column.label ||
      `column_${index + 1}`,
  );
  const values = columnValues(table, column, column.index ?? index);
  const profile = analyzeValues(values);
  const type = inferType(values, column.type || column.valueType || "");
  const role = inferRole({
    header,
    type,
    profile,
    column,
    columns: Array.isArray(table.columns) ? table.columns : [],
    index,
  });

  return {
    ...cloneValue(column),
    header,
    originalHeader: column.originalHeader || header,
    normalizedHeader: normalizedHeaderKey(header),
    canonicalKey:
      column.canonicalKey ||
      column.key ||
      canonicalKeyFromHeader(header, `column_${index + 1}`),
    index: column.index ?? index,
    type,
    role,
    unit: column.unit || column.measureUnit || extractHeaderUnit(header) || "",
    profile: {
      ...cloneValue(column.profile || {}),
      emptyRatio: Number(profile.emptyRatio.toFixed(3)),
      nonEmptyCount: profile.nonEmptyCount,
      uniqueCount: profile.uniqueCount,
      uniqueRatio: Number(profile.uniqueRatio.toFixed(3)),
      numericRatio: Number(profile.numericRatio.toFixed(3)),
      dateRatio: Number(profile.dateRatio.toFixed(3)),
      booleanRatio: Number(profile.booleanRatio.toFixed(3)),
      sampleValues: profile.sampleValues,
    },
    diagnostics: {
      ...cloneValue(column.diagnostics || {}),
      normalizationVersion: NORMALIZATION_VERSION,
    },
  };
}

function rowValues(row, columns = []) {
  return columns.map((column, index) => rowValue(row, column, index));
}

function isRepeatedHeaderRow(row, columns = []) {
  if (!row || !columns.length) return false;

  const values = rowValues(row, columns);
  let comparable = 0;
  let matched = 0;

  values.forEach((value, index) => {
    const actualText = normalizeWhitespace(value);
    const expectedText = normalizeHeader(
      columns[index]?.header || columns[index]?.originalHeader || "",
    );
    const actual = normalizedHeaderKey(actualText);
    const expected = normalizedHeaderKey(expectedText);
    if (!actual || !expected) return;

    if (
      toNumberOrNull(actualText) != null &&
      toNumberOrNull(expectedText) == null
    ) {
      return;
    }

    comparable += 1;
    if (
      actual === expected ||
      (actual.length >= 3 &&
        expected.length >= 3 &&
        (expected.includes(actual) || actual.includes(expected)))
    ) {
      matched += 1;
    }
  });

  return comparable >= 2 && matched / comparable >= 0.6;
}

function isSummaryLabel(value = "") {
  const normalized = normalizeWhitespace(value);
  if (!normalized) return false;
  return (
    SUMMARY_LABEL_PATTERN.test(normalized) ||
    SUMMARY_SUFFIX_PATTERN.test(normalized)
  );
}

function leadingIndent(value = "") {
  const source = asText(value);
  const match = source.match(/^[\s\u3000]+/);
  return match ? match[0].length : 0;
}

function dimensionColumns(columns = []) {
  return columns.filter((column) =>
    ["dimension", "status", "id"].includes(String(column.role || "")),
  );
}

function explicitHierarchyLevel(
  table = {},
  row = {},
  rowIndex = 0,
  primaryDimension = null,
) {
  const metadataCandidates = [
    row?.hierarchyLevel,
    row?.outlineLevel,
    row?.indentLevel,
    row?.meta?.hierarchyLevel,
    row?.meta?.outlineLevel,
    table?.rowMeta?.[rowIndex]?.hierarchyLevel,
    table?.rowMeta?.[rowIndex]?.outlineLevel,
    table?.rowMetadata?.[rowIndex]?.hierarchyLevel,
    table?.rowMetadata?.[rowIndex]?.outlineLevel,
  ];

  for (const value of metadataCandidates) {
    const number = Number(value);
    if (Number.isFinite(number) && number >= 0) return number;
  }

  if (!primaryDimension) return 0;
  return leadingIndent(
    rowValue(row, primaryDimension, primaryDimension.index ?? 0),
  );
}

function classifyRows(
  table = {},
  columns = [],
  valueColumns = [],
  { context = {} } = {},
) {
  const rows = Array.isArray(table.rows) ? table.rows.map(cloneValue) : [];
  const dimensions = dimensionColumns(columns);
  const primaryDimension =
    dimensions.find((column) => String(column.role || "") !== "id") ||
    columns[0] ||
    null;
  const classifications = rows.map(() => ({
    kind: "detail",
    reasons: [],
  }));

  rows.forEach((row, index) => {
    if (isRepeatedHeaderRow(row, columns)) {
      classifications[index] = {
        kind: "excluded",
        reasons: ["REPEATED_HEADER_ROW"],
      };
      return;
    }

    const labels = dimensions
      .map((column) =>
        normalizeWhitespace(rowValue(row, column, column.index ?? 0)),
      )
      .filter(Boolean);

    if (labels.some(isSummaryLabel)) {
      classifications[index] = {
        kind: "summary",
        reasons: ["SUMMARY_LABEL_ROW"],
      };
    }
  });

  if (primaryDimension) {
    for (let index = 0; index < rows.length - 1; index += 1) {
      if (classifications[index].kind !== "detail") continue;

      const currentLevel = explicitHierarchyLevel(
        table,
        rows[index],
        index,
        primaryDimension,
      );
      let nextIndex = index + 1;
      while (
        nextIndex < rows.length &&
        rowValues(rows[nextIndex], columns).every(isBlank)
      ) {
        nextIndex += 1;
      }
      if (nextIndex >= rows.length) continue;

      const nextLevel = explicitHierarchyLevel(
        table,
        rows[nextIndex],
        nextIndex,
        primaryDimension,
      );
      if (nextLevel <= currentLevel) continue;

      let childCount = 0;
      for (let scanIndex = nextIndex; scanIndex < rows.length; scanIndex += 1) {
        if (rowValues(rows[scanIndex], columns).every(isBlank)) {
          continue;
        }

        const scanLevel = explicitHierarchyLevel(
          table,
          rows[scanIndex],
          scanIndex,
          primaryDimension,
        );
        if (scanLevel <= currentLevel) break;
        childCount += 1;
      }

      if (childCount >= 2) {
        classifications[index] = {
          kind: "summary",
          reasons: ["HIERARCHY_PARENT_ROW"],
        };
      }
    }
  }

  const averageValueColumns = valueColumns.filter((entry) => {
    const temporal =
      entry.temporal ||
      parseTemporalHeader(
        entry.column?.header || entry.column?.originalHeader || "",
      );
    if (!temporal) return false;

    const unit =
      entry.column?.unit ||
      temporal.unit ||
      extractHeaderUnit(
        entry.column?.header || entry.column?.originalHeader || "",
      );
    return (
      aggregationForMetric(temporal.metricLabel, unit, entry.column || {}, {
        table,
        profile: entry.profile || {},
        context,
      }) === "average"
    );
  });

  if (
    averageValueColumns.length >= 3 &&
    averageValueColumns.length === valueColumns.length
  ) {
    const allDetailValues = rows
      .flatMap((row, rowIndex) => {
        if (classifications[rowIndex].kind !== "detail") return [];
        return averageValueColumns
          .map((entry) =>
            toNumberOrNull(
              rowValue(row, entry.column, entry.column.index ?? 0),
            ),
          )
          .filter((value) => value != null);
      })
      .map(Math.abs)
      .sort((left, right) => left - right);

    const median = allDetailValues.length
      ? allDetailValues[Math.floor(allDetailValues.length / 2)]
      : 0;

    const placeholderCandidates = [];

    rows.forEach((row, rowIndex) => {
      if (classifications[rowIndex].kind !== "detail") return;

      const values = averageValueColumns
        .map((entry) =>
          toNumberOrNull(rowValue(row, entry.column, entry.column.index ?? 0)),
        )
        .filter((value) => value != null);

      if (values.length < 3) return;
      const first = values[0];
      const repeated = values.every((value) => value === first);

      const absolute = Math.abs(first);
      const distributionOutlier =
        median > 0
          ? absolute >= Math.max(3, median * 4) ||
            (absolute >= 1 && absolute <= median / 4)
          : absolute >= 3;
      if (repeated && Number.isInteger(first) && distributionOutlier) {
        placeholderCandidates.push(rowIndex);
      }
    });

    if (placeholderCandidates.length >= 2) {
      placeholderCandidates.forEach((rowIndex) => {
        classifications[rowIndex] = {
          kind: "summary",
          reasons: ["HIERARCHY_PLACEHOLDER_ROW"],
        };
      });
    }
  }

  return { rows, classifications };
}

function buildDiagnostics(table = {}, columns = [], rows = []) {
  const profiles = columns.map((column) => column.profile || {});
  const emptyRatio = rows.length
    ? rows.reduce((sum, row) => {
        const values = rowValues(row, columns);
        return sum + values.filter(isBlank).length / Math.max(1, values.length);
      }, 0) / rows.length
    : 1;
  const headerConfidence = columns.length
    ? columns.filter(
        (column) =>
          normalizeHeader(column.header) &&
          !/^column_\d+$/i.test(normalizeHeader(column.header)),
      ).length / columns.length
    : 0;
  const typeConsistency = columns.length
    ? columns.filter((column) => column.type && column.type !== "unknown")
        .length / columns.length
    : 0;
  const roleCounts = columns.reduce((acc, column) => {
    const role = column.role || "unknown";
    acc[role] = (acc[role] || 0) + 1;
    return acc;
  }, {});

  const hasMetric = Number(roleCounts.metric || 0) > 0;
  const hasDimension =
    Number(roleCounts.dimension || 0) +
      Number(roleCounts.status || 0) +
      Number(roleCounts.id || 0) >
    0;
  const hasDate = Number(roleCounts.date || 0) > 0;

  const analysisReadiness = {
    groupSummary: {
      ready: hasMetric && hasDimension,
      reasons: [
        hasMetric ? "HAS_METRIC" : "MISSING_METRIC",
        hasDimension ? "HAS_DIMENSION" : "MISSING_DIMENSION",
      ],
    },
    timeTrend: {
      ready: hasMetric && hasDate,
      reasons: [
        hasMetric ? "HAS_METRIC" : "MISSING_METRIC",
        hasDate ? "HAS_DATE" : "MISSING_DATE",
      ],
    },
    categoryCount: {
      ready: hasDimension,
      reasons: [hasDimension ? "HAS_DIMENSION" : "MISSING_DIMENSION"],
    },
    topBottom: {
      ready: hasMetric && hasDimension,
      reasons: [
        hasMetric ? "HAS_METRIC" : "MISSING_METRIC",
        hasDimension ? "HAS_LABEL" : "MISSING_LABEL",
      ],
    },
  };

  const readyCount = Object.values(analysisReadiness).filter(
    (item) => item.ready,
  ).length;
  const confidence = Number(
    (
      headerConfidence * 0.4 +
      typeConsistency * 0.35 +
      (1 - emptyRatio) * 0.25
    ).toFixed(2),
  );

  let queryabilityGrade = "Q2";
  const queryabilityReasons = [];
  if (!rows.length || columns.length <= 1 || confidence < 0.35) {
    queryabilityGrade = "Q0";
    queryabilityReasons.push("INSUFFICIENT_QUERY_STRUCTURE");
  } else if (!readyCount || headerConfidence < 0.5) {
    queryabilityGrade = "Q1";
    queryabilityReasons.push("LOW_ANALYSIS_READINESS");
  } else if (confidence >= 0.72 && readyCount >= 2) {
    queryabilityGrade = "Q3";
    queryabilityReasons.push("HIGH_CONFIDENCE_MULTI_RECIPE_READY");
  } else {
    queryabilityReasons.push("QUERYABLE_ANALYSIS_READY");
  }

  return {
    version: DIAGNOSTICS_VERSION,
    queryabilityGrade,
    queryabilityReasons,
    metrics: {
      rowCount: rows.length,
      columnCount: columns.length,
      emptyRatio: Number(emptyRatio.toFixed(3)),
      headerConfidence: Number(headerConfidence.toFixed(3)),
      typeConsistency: Number(typeConsistency.toFixed(3)),
      confidence,
      excludedRowCount: Array.isArray(table.excludedRows)
        ? table.excludedRows.length
        : 0,
      summaryRowCount: Array.isArray(table.summaryRows)
        ? table.summaryRows.length
        : 0,
      isVirtual: Boolean(table.isVirtual),
      transformationType: table.transformation?.type || null,
      tableUsage: normalizeTableUsage(table),
    },
    transformation: cloneValue(table.transformation || null),
    tableUsage: normalizeTableUsage(table),
    roleCounts,
    analysisReadiness,
    structureSignals: {
      periodMetric: hasMetric && hasDate,
      categorySummary: hasMetric && hasDimension,
      analysisRecipeCount: readyCount,
      supportedAnalysisTypes: Object.entries(analysisReadiness)
        .filter(([, item]) => item.ready)
        .map(([key]) => key),
    },
    inheritedQuality: {
      tableBlockScore: table.score ?? table.blockScore ?? null,
      dataQuality: cloneValue(table.dataQuality || null),
      headerQuality: cloneValue(table.headerQuality || null),
    },
  };
}

function normalizeTable(table = {}, index = 0) {
  const sourceColumns = Array.isArray(table.columns) ? table.columns : [];
  const tableForColumns = {
    ...table,
    rows: Array.isArray(table.rows) ? table.rows : [],
    columns: sourceColumns,
  };
  const columns = sourceColumns.map((column, columnIndex) =>
    normalizeColumn(column, columnIndex, tableForColumns),
  );
  const rows = Array.isArray(table.rows) ? table.rows.map(cloneValue) : [];
  const diagnostics = buildDiagnostics(table, columns, rows);

  return {
    ...cloneValue(table),
    tableId: table.tableId || table.id || `table_${index + 1}`,
    sheetName: table.sheetName || table.sheet || "",
    tableType: table.tableType || "tabular",
    source: table.source || null,
    sourceTableId: table.sourceTableId || null,
    isVirtual: Boolean(table.isVirtual),
    transformation: cloneValue(table.transformation || null),
    tableUsage: normalizeTableUsage(table),
    headerRows: cloneValue(table.headerRows || []),
    dataStartRow: table.dataStartRow ?? null,
    range: table.range || null,
    dataRange: table.dataRange || null,
    columns,
    rows,
    excludedRows: Array.isArray(table.excludedRows)
      ? table.excludedRows.map(cloneValue)
      : [],
    summaryRows: Array.isArray(table.summaryRows)
      ? table.summaryRows.map(cloneValue)
      : [],
    dataQuality: cloneValue(table.dataQuality || null),
    warnings: Array.isArray(table.warnings) ? [...table.warnings] : [],
    confidence: Number.isFinite(Number(table.confidence))
      ? Number(table.confidence)
      : diagnostics.metrics.confidence,
    queryabilityGrade: diagnostics.queryabilityGrade,
    queryabilityReasons: diagnostics.queryabilityReasons,
    diagnostics,
    normalization: {
      version: NORMALIZATION_VERSION,
      physicalSourcePreserved: !table.isVirtual,
    },
  };
}

function numericRatioForColumn(rows = [], column = {}) {
  const values = rows
    .map((row) => rowValue(row, column, column.index ?? 0))
    .filter((value) => !isBlank(value));

  if (!values.length) return 0;
  return (
    values.filter((value) => toNumberOrNull(value) != null).length /
    values.length
  );
}

function makeVirtualColumn({
  header,
  type = "string",
  role = "dimension",
  unit = "",
  semanticType = "",
  aggregation = "",
  sourceColumn = null,
} = {}) {
  return {
    header,
    originalHeader: header,
    key: header,
    canonicalKey: canonicalKeyFromHeader(header, header),
    type,
    role,
    unit,
    semanticType,
    aggregation,
    sourceColumnHeader: sourceColumn?.header || null,
    sourceColumnKey: sourceColumn?.canonicalKey || sourceColumn?.key || null,
  };
}

function uniqueHeader(base = "", used = new Set()) {
  const fallback = normalizeHeader(base) || "값";
  let candidate = fallback;
  let suffix = 2;

  while (used.has(candidate)) {
    candidate = `${fallback}_${suffix}`;
    suffix += 1;
  }
  used.add(candidate);
  return candidate;
}

function unitColumn(table = {}) {
  return (table.columns || []).find((column) =>
    UNIT_HEADER_PATTERN.test(
      normalizeHeader(column.header || column.originalHeader || ""),
    ),
  );
}

function metricIdentityColumn(table = {}) {
  return (table.columns || []).find((column) =>
    METRIC_IDENTITY_HEADER_PATTERN.test(
      normalizeHeader(column.header || column.originalHeader || ""),
    ),
  );
}

function temporalMetricColumns(table = {}) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  return (table.columns || [])
    .map((column) => {
      const profile = numericProfileForColumn(rows, column);
      return {
        column,
        temporal: parseTemporalHeader(
          column.header || column.originalHeader || "",
        ),
        numericRatio: profile.numericRatio,
        profile,
      };
    })
    .filter(
      (item) =>
        item.temporal &&
        item.profile.numericCount >= 1 &&
        (item.profile.numericRatio >= 0.2 || item.column.role === "metric") &&
        !["id", "status"].includes(String(item.column.role || "")),
    );
}

function temporalContextColumns(table = {}, temporalColumns = []) {
  const temporalSet = new Set(temporalColumns.map((item) => item.column));
  return (table.columns || []).filter((column) => {
    if (temporalSet.has(column)) return false;
    const header = normalizeHeader(
      column.header || column.originalHeader || "",
    );
    return (
      column.role === "date" ||
      /^(?:기간|연도|년도|년월|연월|날짜|일자|date|year|month|period)$/i.test(
        header,
      )
    );
  });
}

function preferredDimensionColumns(
  table = {},
  valueColumns = [],
  excludedColumns = [],
) {
  const valueSet = new Set(valueColumns.map((item) => item.column || item));
  const excluded = new Set(excludedColumns.filter(Boolean));

  return (table.columns || [])
    .filter((column) => {
      if (valueSet.has(column) || excluded.has(column)) return false;
      return ["dimension", "status"].includes(String(column.role || ""));
    })
    .sort((left, right) => {
      const leftScore =
        (left.role === "dimension" ? 10 : 0) +
        (DIMENSION_HEADER_PATTERN.test(left.header || "") ? 5 : 0) +
        Number(left.index || 0) / 100;
      const rightScore =
        (right.role === "dimension" ? 10 : 0) +
        (DIMENSION_HEADER_PATTERN.test(right.header || "") ? 5 : 0) +
        Number(right.index || 0) / 100;
      return rightScore - leftScore;
    })
    .slice(0, 6);
}

function pairedDimensionGroupKey(header = "") {
  const normalized = normalizeHeader(header);
  const match = normalized.match(/^(.*?)(?:\(|\[)?[12](?:\)|\])?$/);
  return match ? normalizedHeaderKey(match[1]) : "";
}

function buildDimensionResolvers(table = {}, dimensionSpecs = []) {
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const groupCounts = new Map();

  for (const spec of dimensionSpecs) {
    const groupKey = pairedDimensionGroupKey(
      spec.column.header || spec.column.originalHeader || "",
    );
    if (groupKey) {
      groupCounts.set(groupKey, (groupCounts.get(groupKey) || 0) + 1);
    }
  }

  return dimensionSpecs.map((spec) => {
    const values = rows
      .map((row) => rowValue(row, spec.column, spec.column.index ?? 0))
      .filter((value) => !isBlank(value));
    const normalizedValues = values.map(normalizeWhitespace);
    const numericFrequencies = new Map();
    const textValues = [];

    normalizedValues.forEach((value) => {
      if (toNumberOrNull(value) != null) {
        numericFrequencies.set(value, (numericFrequencies.get(value) || 0) + 1);
      } else {
        textValues.push(value);
      }
    });

    const groupKey = pairedDimensionGroupKey(
      spec.column.header || spec.column.originalHeader || "",
    );
    const pairedGroup = groupKey && (groupCounts.get(groupKey) || 0) >= 2;

    const placeholderValues = new Set(
      [...numericFrequencies.entries()]
        .filter(
          ([, frequency]) =>
            pairedGroup && textValues.length >= 2 && frequency >= 2,
        )
        .map(([value]) => value),
    );

    return {
      ...spec,
      placeholderValues,
      fillDown: placeholderValues.size > 0 && textValues.length >= 2,
      lastAnchor: "",
    };
  });
}

function resolvedDimensionValue(resolver = {}, rawValue) {
  const normalized = normalizeWhitespace(rawValue);
  const isPlaceholder = resolver.placeholderValues?.has(normalized);

  if (
    resolver.fillDown &&
    (isBlank(rawValue) || isPlaceholder) &&
    resolver.lastAnchor
  ) {
    return resolver.lastAnchor;
  }

  if (normalized && !isPlaceholder && toNumberOrNull(normalized) == null) {
    resolver.lastAnchor = normalized;
  }

  return rawValue ?? "";
}

function tableHasAverageContext(table = {}) {
  const context = tableContextText(table);
  if (AVERAGE_CONTEXT_PATTERN.test(context)) return true;

  return (table.columns || []).some((column) =>
    /기간평균|평균값|요약.*평균/i.test(
      normalizeHeader(column.header || column.originalHeader || ""),
    ),
  );
}

function isMeasurementContext(metricLabel = "", table = {}) {
  return MEASUREMENT_CONTEXT_PATTERN.test(
    `${metricLabel} ${tableContextText(table)}`,
  );
}

function aggregationForMetric(
  metricLabel = "",
  unit = "",
  column = {},
  { table = {}, profile = {} } = {},
) {
  const declared = normalizeDeclaredAggregation(
    column.aggregation || column.metricKind || column.meta?.aggregation || "",
  );
  if (declared) return declared;

  const normalizedUnit = normalizeHeader(unit);
  const evidence = [metricLabel, column.semanticType, column.metricKind]
    .filter(Boolean)
    .join(" ");

  if (NON_ADDITIVE_UNIT_PATTERN.test(normalizedUnit)) {
    return "average";
  }

  if (NON_ADDITIVE_METRIC_PATTERN.test(evidence)) {
    return "average";
  }

  if (
    /^회$/i.test(normalizedUnit) &&
    isMeasurementContext(metricLabel, table)
  ) {
    return "average";
  }

  if (tableHasAverageContext(table)) {
    return "average";
  }

  if (
    looksLikePercentageProfile(profile) &&
    !COUNT_CONTEXT_PATTERN.test(metricLabel)
  ) {
    return "average";
  }

  return "sum";
}

function getRowTemporalContext(row, contextColumns = []) {
  for (const column of contextColumns) {
    const parsed = parseTemporalValue(rowValue(row, column, column.index ?? 0));
    if (parsed) return parsed;
  }
  return null;
}

function mergeTemporal(headerTemporal = {}, rowTemporal = null) {
  if (!rowTemporal) return { ...headerTemporal };

  if (
    headerTemporal.type === "month" &&
    /^\d{2}월$/.test(headerTemporal.period)
  ) {
    const month = headerTemporal.period.slice(0, 2);
    if (rowTemporal.year) {
      return {
        ...headerTemporal,
        year: rowTemporal.year,
        period: `${rowTemporal.year}-${month}`,
      };
    }
  }

  if (headerTemporal.type === "year" && rowTemporal?.period) {
    return { ...headerTemporal };
  }

  return { ...headerTemporal };
}

function buildWideToLongVirtualTable(table = {}, index = 0, context = {}) {
  table =
    table?.normalization?.version === NORMALIZATION_VERSION
      ? table
      : normalizeTable(table, index);

  if (!isAnalysisEligibleTable(table)) return null;
  if (table?.transformation?.type) return null;

  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];
  if (!rows.length || columns.length < 3) return null;

  const temporalColumns = temporalMetricColumns(table);
  if (temporalColumns.length < 2) return null;

  const contextColumns = temporalContextColumns(table, temporalColumns);
  const rowUnitColumn = unitColumn(table);
  const rowMetricColumn = metricIdentityColumn(table);
  const dimensions = preferredDimensionColumns(table, temporalColumns, [
    rowUnitColumn,
    rowMetricColumn,
    ...contextColumns,
  ]);

  if (!dimensions.length && !rowMetricColumn) return null;

  const classification = classifyRows(table, columns, temporalColumns, {
    context,
  });
  const usedHeaders = new Set();
  const dimensionSpecs = buildDimensionResolvers(
    table,
    dimensions.map((column) => ({
      column,
      outputHeader: uniqueHeader(
        column.header || column.originalHeader || "구분",
        usedHeaders,
      ),
    })),
  );

  const periodHeader = uniqueHeader("기간", usedHeaders);
  const yearHeader = uniqueHeader("연도", usedHeaders);
  const metricHeader = uniqueHeader("지표명", usedHeaders);
  const unitHeader = uniqueHeader("단위", usedHeaders);
  const valueHeader = uniqueHeader("지표값", usedHeaders);
  const aggregationHeader = uniqueHeader("집계유형", usedHeaders);

  const outputRows = [];
  const summaryRows = [];
  const excludedRows = [];

  rows.forEach((row, rowIndex) => {
    const rowClass = classification.classifications[rowIndex];

    if (rowClass.kind === "excluded") {
      excludedRows.push({
        row: cloneValue(row),
        reason: rowClass.reasons.join(","),
        sourceRowIndex: rowIndex,
      });
      return;
    }

    if (rowClass.kind === "summary") {
      summaryRows.push({
        row: cloneValue(row),
        reason: rowClass.reasons.join(","),
        sourceRowIndex: rowIndex,
      });
      return;
    }

    const rowTemporal = getRowTemporalContext(row, contextColumns);
    const rowMetric = rowMetricColumn
      ? normalizeWhitespace(
          rowValue(row, rowMetricColumn, rowMetricColumn.index ?? 0),
        )
      : "";
    const rowUnit = rowUnitColumn
      ? normalizeWhitespace(
          rowValue(row, rowUnitColumn, rowUnitColumn.index ?? 0),
        )
      : "";

    for (const item of temporalColumns) {
      const value = toNumberOrNull(
        rowValue(row, item.column, item.column.index ?? 0),
      );
      if (value == null) continue;

      const temporal = mergeTemporal(item.temporal, rowTemporal);
      const rawMetricLabel =
        rowMetric ||
        normalizeHeader(item.temporal.metricLabel) ||
        stripUnitSuffix(item.column.header || "") ||
        "지표값";
      const metricLabel = inferGenericCountMetricLabel({
        metricLabel: rawMetricLabel,
        sourceHeader: item.column.header || item.column.originalHeader || "",
        dimensions,
        profile: item.profile,
      });
      const unit = inferMetricUnit({
        metricLabel,
        localUnit:
          rowUnit ||
          item.temporal.unit ||
          item.column.unit ||
          extractHeaderUnit(item.column.header || ""),
        context,
        profile: item.profile,
      });
      const aggregation = aggregationForMetric(metricLabel, unit, item.column, {
        table,
        profile: item.profile,
        context,
      });

      const output = {};
      for (const spec of dimensionSpecs) {
        output[spec.outputHeader] = resolvedDimensionValue(
          spec,
          rowValue(row, spec.column, spec.column.index ?? 0),
        );
      }
      output[periodHeader] = temporal.period || item.temporal.raw;
      if (temporal.year) output[yearHeader] = temporal.year;
      output[metricHeader] = metricLabel;
      if (unit) output[unitHeader] = unit;
      output[valueHeader] = value;
      output[aggregationHeader] = aggregation;
      outputRows.push(output);
    }
  });

  if (!outputRows.length) return null;

  const virtualColumns = [
    ...dimensionSpecs.map((spec) =>
      makeVirtualColumn({
        header: spec.outputHeader,
        type: spec.column.type === "boolean" ? "boolean" : "category",
        role: spec.column.role === "status" ? "status" : "dimension",
        sourceColumn: spec.column,
      }),
    ),
    makeVirtualColumn({
      header: periodHeader,
      type: "date",
      role: "date",
      semanticType: "period",
    }),
    makeVirtualColumn({
      header: yearHeader,
      type: "string",
      role: "date",
      semanticType: "year",
    }),
    makeVirtualColumn({
      header: metricHeader,
      type: "category",
      role: "dimension",
      semanticType: "metricIdentity",
    }),
    makeVirtualColumn({
      header: unitHeader,
      type: "category",
      role: "dimension",
      semanticType: "unit",
    }),
    makeVirtualColumn({
      header: valueHeader,
      type: "number",
      role: "metric",
      semanticType: "measure",
    }),
    makeVirtualColumn({
      header: aggregationHeader,
      type: "category",
      role: "dimension",
      semanticType: "aggregation",
    }),
  ].filter((column) =>
    outputRows.some((row) =>
      Object.prototype.hasOwnProperty.call(row, column.header),
    ),
  );

  return normalizeTable(
    {
      tableId: `${table.tableId || `table_${index + 1}`}#WIDE_LONG`,
      sourceTableId: table.tableId || null,
      sheetName: table.sheetName || "",
      tableType: "wide_to_long",
      source: "normalizedWideToLong",
      isVirtual: true,
      range: table.range || null,
      dataRange: table.dataRange || null,
      headerRows: cloneValue(table.headerRows || []),
      dataStartRow: table.dataStartRow ?? null,
      columns: virtualColumns,
      rows: outputRows,
      excludedRows,
      summaryRows,
      warnings: [],
      confidence: Math.min(
        0.94,
        Math.max(0.7, Number(table.confidence || 0.75)),
      ),
      tableUsage: inheritVirtualTableUsage(table, "wide_to_long"),
      transformation: {
        version: WIDE_TO_LONG_VERSION,
        type: "wideToLong",
        sourceTableId: table.tableId || null,
        sourceColumnCount: columns.length,
        sourceRowCount: rows.length,
        generatedRowCount: outputRows.length,
        excludedRowCount: excludedRows.length,
        summaryRowCount: summaryRows.length,
        dimensionColumns: dimensionSpecs.map((spec) => spec.column.header),
        temporalMetricColumns: temporalColumns.map((item) => ({
          header: item.column.header,
          period: item.temporal.period,
          metricLabel: item.temporal.metricLabel,
          unit: item.temporal.unit || item.column.unit || "",
          numericRatio: Number(item.numericRatio.toFixed(3)),
        })),
        outputHeaders: {
          period: periodHeader,
          year: yearHeader,
          metricName: metricHeader,
          unit: unitHeader,
          metricValue: valueHeader,
          aggregation: aggregationHeader,
        },
        sourcePreserved: true,
      },
    },
    index,
  );
}

function crossMetricColumns(table = {}) {
  const rows = Array.isArray(table.rows) ? table.rows : [];

  return (table.columns || [])
    .map((column) => {
      const header = normalizeHeader(
        column.header || column.originalHeader || "",
      );
      const profile = numericProfileForColumn(rows, column);
      return {
        column,
        header,
        numericRatio: profile.numericRatio,
        profile,
      };
    })
    .filter((item) => {
      if (!item.header || item.header.length > 120) return false;
      if (parseTemporalHeader(item.header)) return false;
      if (UNIT_HEADER_PATTERN.test(item.header)) return false;
      if (METRIC_IDENTITY_HEADER_PATTERN.test(item.header)) return false;
      if (["id", "status", "date"].includes(String(item.column.role || ""))) {
        return false;
      }
      if (item.column.role === "dimension" && item.profile.numericRatio < 0.9) {
        return false;
      }
      return item.profile.numericCount >= 1 && item.profile.numericRatio >= 0.2;
    });
}

function crossDimensionColumns(table = {}, metricColumns = []) {
  const metricSet = new Set(metricColumns.map((item) => item.column));
  return (table.columns || [])
    .filter((column) => !metricSet.has(column))
    .filter((column) => !UNIT_HEADER_PATTERN.test(column.header || ""))
    .filter((column) =>
      ["dimension", "status", "id"].includes(String(column.role || "")),
    )
    .slice(0, 6);
}

function likelyCrossTable(table = {}, metrics = [], dimensions = []) {
  if (metrics.length < 2 || !dimensions.length) return false;
  if (metrics.some((item) => parseTemporalHeader(item.header))) return false;

  const total = Math.max(1, (table.columns || []).length);
  const metricRatio = metrics.length / total;
  return metricRatio >= 0.35 || metrics.length >= 3;
}

function buildCrossTableToLongVirtualTable(
  table = {},
  index = 0,
  context = {},
) {
  table =
    table?.normalization?.version === NORMALIZATION_VERSION
      ? table
      : normalizeTable(table, index);

  if (!isAnalysisEligibleTable(table)) return null;
  if (table?.transformation?.type) return null;

  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columns = Array.isArray(table.columns) ? table.columns : [];
  if (!rows.length || columns.length < 3) return null;

  if (temporalMetricColumns(table).length >= 2) return null;

  const metrics = crossMetricColumns(table);
  const dimensions = crossDimensionColumns(table, metrics);
  if (!likelyCrossTable(table, metrics, dimensions)) return null;

  const classification = classifyRows(table, columns, metrics, { context });
  const rowUnitColumn = unitColumn(table);
  const usedHeaders = new Set();
  const dimensionSpecs = buildDimensionResolvers(
    table,
    dimensions.map((column) => ({
      column,
      outputHeader: uniqueHeader(
        column.header || column.originalHeader || "구분",
        usedHeaders,
      ),
    })),
  );

  const metricHeader = uniqueHeader("지표명", usedHeaders);
  const unitHeader = uniqueHeader("단위", usedHeaders);
  const valueHeader = uniqueHeader("지표값", usedHeaders);
  const aggregationHeader = uniqueHeader("집계유형", usedHeaders);

  const outputRows = [];
  const summaryRows = [];
  const excludedRows = [];

  rows.forEach((row, rowIndex) => {
    const rowClass = classification.classifications[rowIndex];

    if (rowClass.kind === "excluded") {
      excludedRows.push({
        row: cloneValue(row),
        reason: rowClass.reasons.join(","),
        sourceRowIndex: rowIndex,
      });
      return;
    }

    if (rowClass.kind === "summary") {
      summaryRows.push({
        row: cloneValue(row),
        reason: rowClass.reasons.join(","),
        sourceRowIndex: rowIndex,
      });
      return;
    }

    const rowUnit = rowUnitColumn
      ? normalizeWhitespace(
          rowValue(row, rowUnitColumn, rowUnitColumn.index ?? 0),
        )
      : "";

    for (const item of metrics) {
      const value = toNumberOrNull(
        rowValue(row, item.column, item.column.index ?? 0),
      );
      if (value == null) continue;

      const rawMetricLabel =
        stripUnitSuffix(item.header) || item.header || "지표값";
      const metricLabel = inferGenericCountMetricLabel({
        metricLabel: rawMetricLabel,
        sourceHeader: item.header,
        dimensions,
        profile: item.profile,
      });
      const unit = inferMetricUnit({
        metricLabel,
        localUnit:
          rowUnit || item.column.unit || extractHeaderUnit(item.header),
        context,
        profile: item.profile,
      });
      const aggregation = aggregationForMetric(metricLabel, unit, item.column, {
        table,
        profile: item.profile,
        context,
      });

      const output = {};
      for (const spec of dimensionSpecs) {
        output[spec.outputHeader] = resolvedDimensionValue(
          spec,
          rowValue(row, spec.column, spec.column.index ?? 0),
        );
      }
      output[metricHeader] = metricLabel;
      if (unit) output[unitHeader] = unit;
      output[valueHeader] = value;
      output[aggregationHeader] = aggregation;
      outputRows.push(output);
    }
  });

  if (!outputRows.length) return null;

  const virtualColumns = [
    ...dimensionSpecs.map((spec) =>
      makeVirtualColumn({
        header: spec.outputHeader,
        type: spec.column.type === "boolean" ? "boolean" : "category",
        role: spec.column.role === "status" ? "status" : "dimension",
        sourceColumn: spec.column,
      }),
    ),
    makeVirtualColumn({
      header: metricHeader,
      type: "category",
      role: "dimension",
      semanticType: "metricIdentity",
    }),
    makeVirtualColumn({
      header: unitHeader,
      type: "category",
      role: "dimension",
      semanticType: "unit",
    }),
    makeVirtualColumn({
      header: valueHeader,
      type: "number",
      role: "metric",
      semanticType: "measure",
    }),
    makeVirtualColumn({
      header: aggregationHeader,
      type: "category",
      role: "dimension",
      semanticType: "aggregation",
    }),
  ].filter((column) =>
    outputRows.some((row) =>
      Object.prototype.hasOwnProperty.call(row, column.header),
    ),
  );

  return normalizeTable(
    {
      tableId: `${table.tableId || `table_${index + 1}`}#CROSS_LONG`,
      sourceTableId: table.tableId || null,
      sheetName: table.sheetName || "",
      tableType: "cross_table_long",
      source: "normalizedCrossTableToLong",
      isVirtual: true,
      range: table.range || null,
      dataRange: table.dataRange || null,
      headerRows: cloneValue(table.headerRows || []),
      dataStartRow: table.dataStartRow ?? null,
      columns: virtualColumns,
      rows: outputRows,
      excludedRows,
      summaryRows,
      warnings: [],
      confidence: Math.min(
        0.92,
        Math.max(0.68, Number(table.confidence || 0.72)),
      ),
      tableUsage: inheritVirtualTableUsage(table, "cross_table_to_long"),
      transformation: {
        version: CROSS_TO_LONG_VERSION,
        type: "crossTableToLong",
        sourceTableId: table.tableId || null,
        sourceColumnCount: columns.length,
        sourceRowCount: rows.length,
        generatedRowCount: outputRows.length,
        excludedRowCount: excludedRows.length,
        summaryRowCount: summaryRows.length,
        dimensionColumns: dimensionSpecs.map((spec) => spec.column.header),
        measures: metrics.map((item) => ({
          header: item.header,
          unit: item.column.unit || extractHeaderUnit(item.header) || "",
          numericRatio: Number(item.numericRatio.toFixed(3)),
        })),
        outputHeaders: {
          metricName: metricHeader,
          unit: unitHeader,
          metricValue: valueHeader,
          aggregation: aggregationHeader,
        },
        sourcePreserved: true,
      },
    },
    index,
  );
}

function buildNormalizedQueryTables(queryTables = []) {
  if (!Array.isArray(queryTables)) return [];

  const normalizedTables = queryTables.map((table, index) =>
    normalizeTable(table, index),
  );
  const context = buildNormalizationContext(normalizedTables);

  const virtualTables = [];

  normalizedTables.forEach((table, index) => {
    if (!isAnalysisEligibleTable(table)) return;

    const wide = buildWideToLongVirtualTable(table, index, context);
    if (wide) {
      virtualTables.push(wide);
      return;
    }

    const cross = buildCrossTableToLongVirtualTable(table, index, context);
    if (cross) virtualTables.push(cross);
  });

  return [...normalizedTables, ...virtualTables];
}

module.exports = {
  NORMALIZATION_VERSION,
  buildNormalizedQueryTables,
  buildWideToLongVirtualTable,
  buildCrossTableToLongVirtualTable,
  parseTemporalHeader,
  parseTemporalValue,
  isStrictTemporalValue,
  buildNormalizationContext,
  aggregationForMetric,
  normalizeUnitCandidate,
};
