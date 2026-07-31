const {
  normalizeText,
  sha256,
  getQueryTables,
} = require("./queryCandidateObservation");

const QUERY_JSON_SEMANTIC_PROFILE_VERSION = "query_json_semantic_profile_v1";
const QUERY_JSON_TABLE_PROFILE_VERSION = "query_json_table_profile_v1";
const QUERY_JSON_COLUMN_PROFILE_VERSION = "query_json_column_profile_v1";
const NORMALIZER_SAMPLE_ROW_LIMIT = 500;

const ISSUE_LEVELS = Object.freeze(["INFO", "WARNING", "BLOCKING"]);
const DATA_TYPES = Object.freeze([
  "string",
  "number",
  "date",
  "boolean",
  "unknown",
]);

const ROLE_ALIASES = Object.freeze({
  period: ["period", "date", "time", "year", "month", "quarter"],
  revenue: ["revenue", "sales", "sales_amount"],
  amount: ["amount", "money", "currency", "value"],
  cost: ["cost", "expense", "expenditure", "budget_spent"],
  quantity: ["quantity", "qty", "volume"],
  count: ["count", "count_value", "number_of"],
  ratio: ["ratio", "rate", "percentage", "percent"],
  score: ["score", "rating", "grade"],
  status: ["status", "state", "attendance_status"],
  person: ["person", "name", "applicant", "participant", "employee"],
  organization: ["organization", "department", "division", "affiliation"],
  product: ["product", "item", "goods"],
  customer: ["customer", "client", "vendor", "account"],
  location: ["location", "region", "area", "address"],
  category: ["category", "type", "class", "group"],
  identifier: ["identifier", "id", "code", "number"],
  text: ["text", "memo", "note", "description", "content"],
  measure: ["measure", "numeric_measure", "metric"],
  dimension: ["dimension", "categorical_dimension", "entity"],
});

const HEADER_RULES = Object.freeze([
  {
    role: "period",
    patterns: [
      /거래일/,
      /판매일/,
      /신청일/,
      /접수일/,
      /등록일/,
      /작성일/,
      /일자/,
      /날짜/,
      /기간/,
      /연월/,
      /년월/,
      /\byear\b/i,
      /\bmonth\b/i,
      /\bdate\b/i,
      /\bperiod\b/i,
    ],
  },
  {
    role: "revenue",
    patterns: [
      /매출/,
      /판매금액/,
      /매상/,
      /\brevenue\b/i,
      /\bsales(?:amount)?\b/i,
    ],
  },
  {
    role: "cost",
    patterns: [
      /비용/,
      /원가/,
      /지출/,
      /집행액/,
      /소요액/,
      /\bcost\b/i,
      /\bexpense\b/i,
      /\bexpenditure\b/i,
    ],
  },
  {
    role: "quantity",
    patterns: [
      /수량/,
      /물량/,
      /판매량/,
      /재고량/,
      /\bqty\b/i,
      /\bquantity\b/i,
      /\bvolume\b/i,
    ],
  },
  {
    role: "ratio",
    patterns: [
      /비율/,
      /증감률/,
      /달성률/,
      /참여율/,
      /응답률/,
      /만족률/,
      /퍼센트/,
      /\bpercentage\b/i,
      /\bpercent\b/i,
      /\brate\b/i,
    ],
  },
  {
    role: "score",
    patterns: [
      /점수/,
      /평점/,
      /평가점/,
      /등급점수/,
      /\bscore\b/i,
      /\brating\b/i,
    ],
  },
  {
    role: "status",
    patterns: [
      /출석상태/,
      /참석상태/,
      /진행상태/,
      /처리상태/,
      /승인여부/,
      /상태/,
      /여부/,
      /\bstatus\b/i,
      /\bstate\b/i,
      /\battendance\b/i,
    ],
  },
  {
    role: "customer",
    patterns: [
      /거래처/,
      /고객/,
      /업체명/,
      /공급업체/,
      /수요처/,
      /\bcustomer\b/i,
      /\bclient\b/i,
      /\bvendor\b/i,
    ],
  },
  {
    role: "product",
    patterns: [
      /품목/,
      /상품/,
      /제품/,
      /물품/,
      /자재명/,
      /\bproduct\b/i,
      /\bitem\b/i,
      /\bgoods\b/i,
    ],
  },
  {
    role: "person",
    patterns: [
      /성명/,
      /이름/,
      /신청자/,
      /참석자/,
      /지원자/,
      /학생명/,
      /교수명/,
      /담당자/,
      /연구자/,
      /\bperson\b/i,
      /\bname\b/i,
      /\bapplicant\b/i,
      /\bparticipant\b/i,
    ],
  },
  {
    role: "organization",
    patterns: [
      /소속/,
      /부서/,
      /기관/,
      /조직/,
      /학과/,
      /연구소/,
      /사업단/,
      /\borganization\b/i,
      /\bdepartment\b/i,
      /\bdivision\b/i,
      /\baffiliation\b/i,
    ],
  },
  {
    role: "location",
    patterns: [
      /지역/,
      /권역/,
      /주소/,
      /장소/,
      /시도/,
      /시군구/,
      /\blocation\b/i,
      /\bregion\b/i,
      /\barea\b/i,
      /\baddress\b/i,
    ],
  },
  {
    role: "category",
    patterns: [
      /유형/,
      /종류/,
      /분류/,
      /구분/,
      /카테고리/,
      /분야/,
      /등급/,
      /\bcategory\b/i,
      /\btype\b/i,
      /\bclass\b/i,
      /\bgroup\b/i,
    ],
  },
  {
    role: "count",
    patterns: [
      /건수/,
      /인원수/,
      /횟수/,
      /개수/,
      /명수/,
      /\bcount\b/i,
      /\bnumberof\b/i,
    ],
  },
  {
    role: "amount",
    patterns: [
      /금액/,
      /예산액/,
      /총액/,
      /단가/,
      /잔액/,
      /\bamount\b/i,
      /\bprice\b/i,
      /\bbalance\b/i,
    ],
  },
  {
    role: "identifier",
    patterns: [
      /식별/,
      /관리번호/,
      /접수번호/,
      /신청번호/,
      /학번/,
      /사번/,
      /코드/,
      /번호$/,
      /^id$/i,
      /_id$/i,
      /\bcode\b/i,
    ],
  },
  {
    role: "text",
    patterns: [
      /내용/,
      /비고/,
      /메모/,
      /설명/,
      /의견/,
      /사유/,
      /\btext\b/i,
      /\bmemo\b/i,
      /\bnote\b/i,
      /\bdescription\b/i,
    ],
  },
]);

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const result = [];
  for (const value of asArray(values)) {
    const text = normalizeText(value);
    if (!text || seen.has(text)) continue;
    seen.add(text);
    result.push(text);
  }
  return result;
}

function uniqueSorted(values = []) {
  return unique(values).sort((left, right) => left.localeCompare(right, "ko"));
}

function finiteNumber(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function round(value, digits = 6) {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  return Number(number.toFixed(digits));
}

function normalizeKey(value = "") {
  return normalizeText(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s_./:\-()[\]{}]+/g, "")
    .replace(/[^가-힣a-z0-9%]/g, "");
}

function isEmptyValue(value) {
  if (value == null) return true;
  if (typeof value === "string") return normalizeText(value) === "";
  return false;
}

function parseNumeric(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string") return null;
  const text = normalizeText(value);
  if (!text) return null;
  const negative = /^\(.*\)$/.test(text);
  const normalized = text
    .replace(/[₩$€£¥,%\s]/g, "")
    .replace(/^\((.*)\)$/, "$1");
  if (!/^[-+]?\d+(?:\.\d+)?$/.test(normalized)) return null;
  const number = Number(normalized);
  if (!Number.isFinite(number)) return null;
  return negative ? -number : number;
}

function looksLikeDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;
  if (typeof value !== "string") return false;
  const text = normalizeText(value);
  if (!text) return false;
  return (
    /^\d{4}[-/.]\d{1,2}(?:[-/.]\d{1,2})?(?:\s.*)?$/.test(text) ||
    /^\d{4}년\s*\d{1,2}월(?:\s*\d{1,2}일)?$/.test(text) ||
    /^\d{4}-?Q[1-4]$/i.test(text) ||
    /^\d{4}[-/.]\d{1,2}$/.test(text)
  );
}

function looksLikeBoolean(value) {
  if (typeof value === "boolean") return true;
  if (typeof value !== "string") return false;
  return /^(true|false|yes|no|y|n|예|아니오|유|무|완료|미완료)$/i.test(
    normalizeText(value),
  );
}

function normalizeDataType(value = "") {
  const key = normalizeKey(value);
  if (!key) return "unknown";
  if (/^(date|datetime|time|period|year|month|quarter|timestamp)/.test(key)) {
    return "date";
  }
  if (
    /^(number|numeric|integer|int|float|double|decimal|currency|money|percentage|percent)/.test(
      key,
    )
  ) {
    return "number";
  }
  if (/^(boolean|bool)/.test(key)) return "boolean";
  if (/^(string|text|category|categorical|varchar|char|object)/.test(key)) {
    return "string";
  }
  return "unknown";
}

function inferDataType(values = [], explicitType = "") {
  const normalizedExplicit = normalizeDataType(explicitType);
  if (normalizedExplicit !== "unknown") {
    return {
      dataType: normalizedExplicit,
      confidence: 0.99,
      source: "explicit",
    };
  }
  const nonEmpty = values.filter((value) => !isEmptyValue(value));
  if (!nonEmpty.length) {
    return { dataType: "unknown", confidence: 0, source: "empty" };
  }
  const dateCount = nonEmpty.filter(looksLikeDate).length;
  const numericCount = nonEmpty.filter(
    (value) => parseNumeric(value) != null,
  ).length;
  const booleanCount = nonEmpty.filter(looksLikeBoolean).length;
  const denominator = nonEmpty.length;
  if (dateCount / denominator >= 0.8) {
    return {
      dataType: "date",
      confidence: round(dateCount / denominator),
      source: "values",
    };
  }
  if (numericCount / denominator >= 0.8) {
    return {
      dataType: "number",
      confidence: round(numericCount / denominator),
      source: "values",
    };
  }
  if (booleanCount / denominator >= 0.8) {
    return {
      dataType: "boolean",
      confidence: round(booleanCount / denominator),
      source: "values",
    };
  }
  return { dataType: "string", confidence: 0.8, source: "values" };
}

function normalizeExplicitRole(value = "") {
  const key = normalizeKey(value);
  if (!key) return "";
  for (const [role, aliases] of Object.entries(ROLE_ALIASES)) {
    if (role === key || aliases.some((alias) => normalizeKey(alias) === key)) {
      return role;
    }
  }
  return normalizeText(value).toLowerCase();
}

function headerRole(header = "") {
  const text = normalizeText(header);
  if (!text) return "";
  const compact = normalizeKey(text);
  for (const rule of HEADER_RULES) {
    if (
      rule.patterns.some(
        (pattern) => pattern.test(text) || pattern.test(compact),
      )
    ) {
      return rule.role;
    }
  }
  return "";
}

function isGenericRole(role = "") {
  return [
    "measure",
    "dimension",
    "entity",
    "category",
    "group",
    "string",
    "number",
    "metric",
  ].includes(normalizeKey(role));
}

function inferSemanticRole({
  header = "",
  explicitRole = "",
  dataType = "unknown",
} = {}) {
  const normalizedExplicit = normalizeExplicitRole(explicitRole);
  const inferredFromHeader = headerRole(header);
  if (normalizedExplicit && !isGenericRole(normalizedExplicit)) {
    return { role: normalizedExplicit, confidence: 0.99, source: "explicit" };
  }
  if (inferredFromHeader) {
    return { role: inferredFromHeader, confidence: 0.94, source: "header" };
  }
  if (normalizedExplicit) {
    return {
      role: normalizedExplicit,
      confidence: 0.88,
      source: "explicit_generic",
    };
  }
  if (dataType === "date") {
    return { role: "period", confidence: 0.72, source: "data_type" };
  }
  if (dataType === "number") {
    return { role: "measure", confidence: 0.62, source: "data_type" };
  }
  if (dataType === "string") {
    return { role: "dimension", confidence: 0.55, source: "data_type" };
  }
  return { role: "unknown", confidence: 0, source: "unknown" };
}

function semanticTypeFor(role = "", dataType = "unknown") {
  if (role === "period") return "dimension";
  if (
    [
      "revenue",
      "amount",
      "cost",
      "quantity",
      "count",
      "ratio",
      "score",
      "measure",
    ].includes(role)
  ) {
    return "measure";
  }
  if (role === "identifier") return "identifier";
  if (role === "text") return "text";
  if (
    [
      "status",
      "person",
      "organization",
      "product",
      "customer",
      "location",
      "category",
      "dimension",
    ].includes(role)
  ) {
    return "dimension";
  }
  if (dataType === "number") return "measure";
  if (dataType === "date" || dataType === "string" || dataType === "boolean")
    return "dimension";
  return "unknown";
}

function metricFamilyFor(role = "", header = "") {
  if (role === "revenue") return "sales";
  if (["cost", "quantity", "count", "ratio", "score", "amount"].includes(role))
    return role;
  const compact = normalizeKey(header);
  if (/예산|집행|잔액|budget/.test(compact)) return "budget";
  if (/인건비|급여|salary|payroll/.test(compact)) return "labor_cost";
  return "";
}

function operationCapabilities({
  role = "",
  semanticType = "",
  dataType = "unknown",
} = {}) {
  const operations = [];
  if (semanticType === "measure" && dataType === "number") {
    operations.push("sum", "average", "min", "max");
  }
  if (semanticType === "dimension") {
    operations.push("groupBy", "countDistinct");
  }
  if (role === "period" || dataType === "date") {
    operations.push("timeBucket", "min", "max");
  }
  return unique(operations);
}

function derivedRoles(role = "", semanticType = "") {
  const roles = [role];
  if (semanticType === "dimension") roles.push("dimension");
  if (semanticType === "measure") roles.push("measure");
  if (
    [
      "organization",
      "category",
      "customer",
      "product",
      "status",
      "location",
      "person",
    ].includes(role)
  ) {
    roles.push("group");
  }
  if (
    ["revenue", "cost", "quantity", "count", "ratio", "score"].includes(role)
  ) {
    roles.push("amount");
  }
  return unique(roles);
}

function issue(code, level, message, details = {}) {
  return {
    code: normalizeText(code),
    level: ISSUE_LEVELS.includes(level) ? level : "WARNING",
    message: normalizeText(message),
    details: details && typeof details === "object" ? details : {},
  };
}

function explicitColumns(table = {}) {
  const source = asArray(table.columns || table.headers || table.fields);
  return source.map((column, index) => {
    if (typeof column === "string") {
      return { index, header: normalizeText(column), source: {} };
    }
    const object = column && typeof column === "object" ? column : {};
    return {
      index,
      header: normalizeText(
        object.header ||
          object.name ||
          object.label ||
          object.key ||
          object.sourceHeader ||
          object.columnName ||
          "",
      ),
      source: object,
    };
  });
}

function firstObjectRow(rows = []) {
  return rows.find(
    (row) => row && typeof row === "object" && !Array.isArray(row),
  );
}

function tableColumns(table = {}) {
  const direct = explicitColumns(table);
  if (direct.length) return direct;
  const rows = asArray(table.rows || table.data || table.records);
  const first = firstObjectRow(rows);
  if (first) {
    return Object.keys(first).map((header, index) => ({
      index,
      header,
      source: {},
    }));
  }
  const maxWidth = rows.reduce(
    (max, row) => (Array.isArray(row) ? Math.max(max, row.length) : max),
    0,
  );
  return Array.from({ length: maxWidth }, (_, index) => ({
    index,
    header: `column_${index + 1}`,
    source: {},
  }));
}

function tableRows(table = {}) {
  return asArray(table.rows || table.data || table.records || table.sampleRows);
}

function sampleRowsForAnalysis(rows = [], limit = NORMALIZER_SAMPLE_ROW_LIMIT) {
  const source = asArray(rows);
  if (source.length <= limit) return source;
  const tailCount = Math.min(50, Math.floor(limit / 5));
  const headCount = limit - tailCount;
  return [...source.slice(0, headCount), ...source.slice(-tailCount)];
}

function valueForColumn(row, column) {
  if (Array.isArray(row)) return row[column.index];
  if (!row || typeof row !== "object") return undefined;
  const source = column.source || {};
  const keys = unique([
    column.header,
    source.key,
    source.name,
    source.header,
    source.sourceHeader,
    source.columnName,
    source.id,
  ]);
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
  }
  const rowKeys = Object.keys(row);
  return row[rowKeys[column.index]];
}

function safeSampleValue(value) {
  if (value == null) return null;
  if (value instanceof Date) return value.toISOString();
  if (["string", "number", "boolean"].includes(typeof value)) return value;
  return normalizeText(JSON.stringify(value)).slice(0, 120);
}

function subtotalKind(row) {
  const values = Array.isArray(row)
    ? row
    : row && typeof row === "object"
      ? Object.values(row)
      : [];
  const text = values.slice(0, 4).map(normalizeText).join(" ");
  if (!text) return "";
  if (/(총계|전체합계|grand\s*total)/i.test(text)) return "total";
  if (/(소계|부분합계|subtotal)/i.test(text)) return "subtotal";
  if (/(합계|total)/i.test(text)) return "total";
  return "";
}

function headerDiagnostics(columns = []) {
  const normalized = columns.map((column) => normalizeKey(column.header));
  const blankHeaderIndexes = columns
    .filter((column) => !normalizeText(column.header))
    .map((column) => column.index);
  const seen = new Map();
  for (const key of normalized.filter(Boolean))
    seen.set(key, (seen.get(key) || 0) + 1);
  const duplicateHeaders = Array.from(seen.entries())
    .filter(([, count]) => count > 1)
    .map(([header, count]) => ({ header, count }));
  return { blankHeaderIndexes, duplicateHeaders };
}

function firstDefinedBoolean(...values) {
  for (const value of values) {
    if (typeof value === "boolean") return value;
  }
  return false;
}

function tableFlags(table = {}) {
  const tableUsage =
    table.tableUsage &&
    typeof table.tableUsage === "object" &&
    !Array.isArray(table.tableUsage)
      ? table.tableUsage
      : {};
  const usage =
    table.usage &&
    typeof table.usage === "object" &&
    !Array.isArray(table.usage)
      ? table.usage
      : {};
  const sourceTablePolicy =
    table.sourceTablePolicy &&
    typeof table.sourceTablePolicy === "object" &&
    !Array.isArray(table.sourceTablePolicy)
      ? table.sourceTablePolicy
      : {};
  const primarySelection =
    table.primarySelection &&
    typeof table.primarySelection === "object" &&
    !Array.isArray(table.primarySelection)
      ? table.primarySelection
      : {};

  return {
    primary: firstDefinedBoolean(
      table.isPrimary,
      table.primary,
      primarySelection.selected,
      tableUsage.primary,
      usage.primary,
      sourceTablePolicy.primary,
    ),
    analysisEligible: firstDefinedBoolean(
      table.analysisEligible,
      table.eligible,
      table.isEligible,
      tableUsage.analysisEligible,
      tableUsage.analysis,
      tableUsage.eligible,
      usage.analysisEligible,
      usage.analysis,
      usage.eligible,
      sourceTablePolicy.analysisEligible,
      sourceTablePolicy.analysis,
      sourceTablePolicy.eligible,
    ),
    templateEligible: firstDefinedBoolean(
      table.templateEligible,
      tableUsage.templateEligible,
      tableUsage.template,
      usage.templateEligible,
      usage.template,
      sourceTablePolicy.templateEligible,
      sourceTablePolicy.template,
    ),
    virtual: firstDefinedBoolean(
      table.isVirtual,
      table.virtual,
      table.transformation?.virtual,
    ),
  };
}

function detectMergedHeader(table = {}) {
  return Boolean(
    table.mergedHeader === true ||
    table.hasMergedHeader === true ||
    asArray(table.mergedCells || table.merges).length > 0 ||
    finiteNumber(table.headerRowCount || table.headerRows?.length, 1) > 1,
  );
}

function buildColumnProfile({ tableId, column, rows }) {
  const values = rows.map((row) => valueForColumn(row, column));
  const nonEmpty = values.filter((value) => !isEmptyValue(value));
  const dataTypeResult = inferDataType(
    nonEmpty,
    column.source?.dataType ||
      column.source?.type ||
      column.source?.inferredType ||
      "",
  );
  const roleResult = inferSemanticRole({
    header: column.header,
    explicitRole:
      column.source?.semanticRole ||
      column.source?.role ||
      column.source?.detectedRole ||
      column.source?.metricRole ||
      "",
    dataType: dataTypeResult.dataType,
  });
  const semanticType = semanticTypeFor(
    roleResult.role,
    dataTypeResult.dataType,
  );
  const metricFamily = metricFamilyFor(roleResult.role, column.header);
  const operations = operationCapabilities({
    role: roleResult.role,
    semanticType,
    dataType: dataTypeResult.dataType,
  });
  const roleSet = derivedRoles(roleResult.role, semanticType);
  const capabilities = unique([
    `data_type:${dataTypeResult.dataType}`,
    `semantic_type:${semanticType}`,
    ...roleSet.map((role) => `column_role:${role}`),
    metricFamily ? `metric_family:${metricFamily}` : "",
    ...operations.map((operation) => `operation:${operation}`),
  ]);
  const uniqueValues = new Set(nonEmpty.map((value) => normalizeText(value)))
    .size;
  const sampleValues = unique(nonEmpty.slice(0, 8).map(safeSampleValue)).slice(
    0,
    5,
  );
  const issues = [];
  if (!column.header) {
    issues.push(
      issue("blank_header", "WARNING", "열 머리글이 비어 있습니다.", {
        columnIndex: column.index,
      }),
    );
  }
  if (!nonEmpty.length) {
    issues.push(
      issue("empty_column", "WARNING", "유효한 값이 없는 열입니다.", {
        columnIndex: column.index,
      }),
    );
  }
  if (roleResult.confidence < 0.7 && nonEmpty.length) {
    issues.push(
      issue("low_role_confidence", "INFO", "열 역할 추론 신뢰도가 낮습니다.", {
        confidence: roleResult.confidence,
      }),
    );
  }

  const profile = {
    version: QUERY_JSON_COLUMN_PROFILE_VERSION,
    columnId: normalizeText(
      column.source?.columnId ||
        column.source?.id ||
        `${tableId}.column_${column.index + 1}`,
    ),
    index: column.index,
    sourceHeader: normalizeText(column.header),
    normalizedHeader: normalizeText(column.header).normalize("NFKC"),
    dataType: dataTypeResult.dataType,
    dataTypeConfidence: round(dataTypeResult.confidence),
    dataTypeSource: dataTypeResult.source,
    semanticRole: roleResult.role,
    roleAliases: roleSet,
    semanticType,
    roleConfidence: round(roleResult.confidence),
    roleSource: roleResult.source,
    metricFamily,
    supportedOperations: operations,
    capabilities,
    stats: {
      sampledValueCount: values.length,
      nonEmptyCount: nonEmpty.length,
      nonEmptyRatio: values.length ? round(nonEmpty.length / values.length) : 0,
      uniqueCount: uniqueValues,
      uniqueRatio: nonEmpty.length ? round(uniqueValues / nonEmpty.length) : 0,
    },
    sampleValues,
    evidence: {
      explicitDataType: normalizeText(
        column.source?.dataType ||
          column.source?.type ||
          column.source?.inferredType ||
          "",
      ),
      explicitSemanticRole: normalizeText(
        column.source?.semanticRole ||
          column.source?.role ||
          column.source?.detectedRole ||
          column.source?.metricRole ||
          "",
      ),
    },
    issues,
  };
  profile.columnSha256 = sha256({ ...profile, columnSha256: undefined });
  return profile;
}

function buildTableProfile(table = {}, index = 0) {
  const flags = tableFlags(table);
  const sourceRows = tableRows(table);
  const rows = sampleRowsForAnalysis(sourceRows);
  const columns = tableColumns(table);
  const tableId = normalizeText(
    table.tableId || table.id || `table_${index + 1}`,
  );
  const diagnostics = headerDiagnostics(columns);
  const subtotalRows = rows.filter(
    (row) => subtotalKind(row) === "subtotal",
  ).length;
  const totalRows = rows.filter((row) => subtotalKind(row) === "total").length;
  const profiles = columns.map((column) =>
    buildColumnProfile({ tableId, column, rows }),
  );
  const availableRoles = uniqueSorted(
    profiles.flatMap((column) => column.roleAliases),
  );
  const metricFamilies = uniqueSorted(
    profiles.map((column) => column.metricFamily),
  );
  const supportedOperations = unique([
    "countRows",
    ...profiles.flatMap((column) => column.supportedOperations),
  ]);
  const dimensionCount = profiles.filter(
    (column) => column.semanticType === "dimension",
  ).length;
  const measureCount = profiles.filter(
    (column) => column.semanticType === "measure",
  ).length;
  if (dimensionCount && measureCount) supportedOperations.push("rank");
  if (
    profiles.some((column) => column.semanticRole === "period") &&
    measureCount
  ) {
    supportedOperations.push("timeSeries");
  }
  const capabilities = uniqueSorted([
    "operation:countRows",
    dimensionCount ? "group_by" : "",
    dimensionCount ? "metric_kind:aggregate" : "",
    dimensionCount && measureCount ? "ranking" : "",
    dimensionCount && measureCount ? "metric_kind:rank" : "",
    dimensionCount && measureCount ? "metric_dependency" : "",
    ...profiles.flatMap((column) => column.capabilities),
    ...unique(supportedOperations).map((operation) => `operation:${operation}`),
    flags.analysisEligible ? "table:analysis_eligible" : "",
    flags.templateEligible ? "table:template_eligible" : "",
    flags.primary ? "table:primary" : "",
    flags.virtual ? "table:virtual" : "table:physical",
  ]);
  const issues = [];
  if (!flags.analysisEligible) {
    issues.push(
      issue(
        "analysis_ineligible_table",
        "INFO",
        "분석 대상에서 제외된 표입니다.",
      ),
    );
  }
  if (!columns.length)
    issues.push(issue("no_columns", "BLOCKING", "표에 열이 없습니다."));
  if (!rows.length)
    issues.push(issue("no_rows", "WARNING", "표에 데이터 행이 없습니다."));
  if (diagnostics.blankHeaderIndexes.length) {
    issues.push(
      issue("blank_headers", "WARNING", "빈 머리글이 존재합니다.", diagnostics),
    );
  }
  if (diagnostics.duplicateHeaders.length) {
    issues.push(
      issue(
        "duplicate_headers",
        "WARNING",
        "중복 머리글이 존재합니다.",
        diagnostics,
      ),
    );
  }
  if (detectMergedHeader(table)) {
    issues.push(
      issue(
        "merged_header_detected",
        "WARNING",
        "병합 또는 다중 머리글 구조가 감지됐습니다.",
      ),
    );
  }
  if (subtotalRows || totalRows) {
    issues.push(
      issue(
        "summary_rows_detected",
        "INFO",
        "합계 또는 소계 행이 감지됐습니다.",
        { subtotalRows, totalRows },
      ),
    );
  }
  if (
    profiles.length &&
    !profiles.some((column) => column.semanticType === "measure")
  ) {
    issues.push(
      issue("no_measure_column", "INFO", "측정값 역할의 열이 없습니다."),
    );
  }

  const nonEmptyCellCount = profiles.reduce(
    (sum, column) => sum + column.stats.nonEmptyCount,
    0,
  );
  const sampledCellCount = profiles.reduce(
    (sum, column) => sum + column.stats.sampledValueCount,
    0,
  );
  const profile = {
    version: QUERY_JSON_TABLE_PROFILE_VERSION,
    index,
    tableId,
    sourceTableId: normalizeText(
      table.sourceTableId || table.transformation?.sourceTableId || "",
    ),
    sourceSheetName: normalizeText(
      table.sourceSheetName || table.sheetName || table.name || "",
    ),
    flags,
    shape: {
      rowCount: finiteNumber(
        table.rowCount || table.dataRowCount || table.stats?.rowCount,
        sourceRows.length,
      ),
      sampledRowCount: rows.length,
      columnCount: columns.length,
      blankHeaderCount: diagnostics.blankHeaderIndexes.length,
      duplicateHeaderCount: diagnostics.duplicateHeaders.length,
      mergedHeader: detectMergedHeader(table),
      subtotalRowCount: subtotalRows,
      totalRowCount: totalRows,
      multiHeaderRowCount: finiteNumber(
        table.headerRowCount || table.headerRows?.length,
        1,
      ),
    },
    quality: {
      sampledCellCount,
      nonEmptyCellCount,
      nonEmptyRatio: sampledCellCount
        ? round(nonEmptyCellCount / sampledCellCount)
        : 0,
    },
    availableRoles,
    metricFamilies,
    supportedOperations: unique(supportedOperations),
    capabilities,
    columns: profiles,
    issues: [...issues, ...profiles.flatMap((column) => column.issues)],
  };
  profile.tableSha256 = sha256({ ...profile, tableSha256: undefined });
  return profile;
}

function buildQueryJsonSemanticProfile({
  queryJson = {},
  caseId = "",
  fileName = "",
} = {}) {
  const tables = getQueryTables(queryJson).map(buildTableProfile);
  const issues = [];
  if (!tables.length) {
    issues.push(
      issue(
        "no_query_tables",
        "BLOCKING",
        "queryJson에서 표를 찾을 수 없습니다.",
      ),
    );
  }
  if (tables.length > 1) {
    issues.push(
      issue("multiple_tables", "INFO", "queryJson에 여러 표가 존재합니다.", {
        tableCount: tables.length,
      }),
    );
  }
  if (tables.length && !tables.some((table) => table.flags.analysisEligible)) {
    issues.push(
      issue(
        "no_analysis_eligible_table",
        "WARNING",
        "분석 가능한 표가 없습니다.",
      ),
    );
  }
  const allIssues = [...issues, ...tables.flatMap((table) => table.issues)];
  const allColumns = tables.flatMap((table) => table.columns);
  const profile = {
    version: QUERY_JSON_SEMANTIC_PROFILE_VERSION,
    tableProfileVersion: QUERY_JSON_TABLE_PROFILE_VERSION,
    columnProfileVersion: QUERY_JSON_COLUMN_PROFILE_VERSION,
    source: {
      caseId: normalizeText(caseId),
      fileName: normalizeText(
        fileName || queryJson.fileName || queryJson.sourceFileName || "",
      ),
      querySchemaVersion: normalizeText(
        queryJson.schemaVersion || queryJson.version || "",
      ),
      querySha256: sha256(queryJson),
      tableSourceKey: Array.isArray(queryJson.normalizedQueryTables)
        ? "normalizedQueryTables"
        : Array.isArray(queryJson.tables)
          ? "tables"
          : Array.isArray(queryJson.queryTables)
            ? "queryTables"
            : "none",
    },
    counts: {
      totalTables: tables.length,
      analysisEligibleTables: tables.filter(
        (table) => table.flags.analysisEligible,
      ).length,
      templateEligibleTables: tables.filter(
        (table) => table.flags.templateEligible,
      ).length,
      primaryTables: tables.filter((table) => table.flags.primary).length,
      virtualTables: tables.filter((table) => table.flags.virtual).length,
      totalColumns: allColumns.length,
      measureColumns: allColumns.filter(
        (column) => column.semanticType === "measure",
      ).length,
      dimensionColumns: allColumns.filter(
        (column) => column.semanticType === "dimension",
      ).length,
      blockingIssues: allIssues.filter((item) => item.level === "BLOCKING")
        .length,
      warningIssues: allIssues.filter((item) => item.level === "WARNING")
        .length,
      infoIssues: allIssues.filter((item) => item.level === "INFO").length,
    },
    availableRoles: uniqueSorted(
      tables.flatMap((table) => table.availableRoles),
    ),
    metricFamilies: uniqueSorted(
      tables.flatMap((table) => table.metricFamilies),
    ),
    supportedOperations: uniqueSorted(
      tables.flatMap((table) => table.supportedOperations),
    ),
    availableCapabilities: uniqueSorted([
      ...tables.flatMap((table) => table.capabilities),
      tables.length === 1 ? "single_table" : "",
      tables.length > 1 ? "multi_table" : "",
      `table_count:${tables.length}`,
    ]),
    tables,
    issues: allIssues,
  };
  profile.profileSha256 = sha256({ ...profile, profileSha256: undefined });
  return profile;
}

function validationIssue(path, code, message) {
  return { path, code, message };
}

function validateColumnProfile(column = {}, path = "column") {
  const errors = [];
  const warnings = [];
  if (column.version !== QUERY_JSON_COLUMN_PROFILE_VERSION) {
    errors.push(
      validationIssue(
        `${path}.version`,
        "invalid_version",
        "column profile version이 유효하지 않습니다.",
      ),
    );
  }
  if (!normalizeText(column.columnId)) {
    errors.push(
      validationIssue(`${path}.columnId`, "required", "columnId가 필요합니다."),
    );
  }
  if (!DATA_TYPES.includes(column.dataType)) {
    errors.push(
      validationIssue(
        `${path}.dataType`,
        "invalid_enum",
        "dataType이 유효하지 않습니다.",
      ),
    );
  }
  for (const field of ["dataTypeConfidence", "roleConfidence"]) {
    const value = Number(column[field]);
    if (!Number.isFinite(value) || value < 0 || value > 1) {
      errors.push(
        validationIssue(
          `${path}.${field}`,
          "invalid_range",
          `${field}는 0~1이어야 합니다.`,
        ),
      );
    }
  }
  for (const field of ["nonEmptyRatio", "uniqueRatio"]) {
    const value = Number(column.stats?.[field]);
    if (!Number.isFinite(value) || value < 0 || value > 1) {
      errors.push(
        validationIssue(
          `${path}.stats.${field}`,
          "invalid_range",
          `${field}는 0~1이어야 합니다.`,
        ),
      );
    }
  }
  const expectedSha = sha256({ ...column, columnSha256: undefined });
  if (column.columnSha256 !== expectedSha) {
    errors.push(
      validationIssue(
        `${path}.columnSha256`,
        "sha_mismatch",
        "column SHA-256이 일치하지 않습니다.",
      ),
    );
  }
  if (column.roleConfidence < 0.7) {
    warnings.push(
      validationIssue(
        `${path}.roleConfidence`,
        "low_role_confidence",
        "역할 신뢰도가 낮습니다.",
      ),
    );
  }
  return { errors, warnings };
}

function validateTableProfile(table = {}, index = 0) {
  const path = `tables[${index}]`;
  const errors = [];
  const warnings = [];
  if (table.version !== QUERY_JSON_TABLE_PROFILE_VERSION) {
    errors.push(
      validationIssue(
        `${path}.version`,
        "invalid_version",
        "table profile version이 유효하지 않습니다.",
      ),
    );
  }
  if (!normalizeText(table.tableId)) {
    errors.push(
      validationIssue(`${path}.tableId`, "required", "tableId가 필요합니다."),
    );
  }
  if (!Array.isArray(table.columns)) {
    errors.push(
      validationIssue(
        `${path}.columns`,
        "invalid_type",
        "columns는 배열이어야 합니다.",
      ),
    );
  } else {
    const ids = new Set();
    table.columns.forEach((column, columnIndex) => {
      const validation = validateColumnProfile(
        column,
        `${path}.columns[${columnIndex}]`,
      );
      errors.push(...validation.errors);
      warnings.push(...validation.warnings);
      if (ids.has(column.columnId)) {
        errors.push(
          validationIssue(
            `${path}.columns[${columnIndex}].columnId`,
            "duplicate",
            "columnId가 중복됩니다.",
          ),
        );
      }
      ids.add(column.columnId);
    });
  }
  if (Number(table.shape?.columnCount || 0) !== asArray(table.columns).length) {
    errors.push(
      validationIssue(
        `${path}.shape.columnCount`,
        "count_mismatch",
        "columnCount가 실제 열 수와 다릅니다.",
      ),
    );
  }
  const expectedSha = sha256({ ...table, tableSha256: undefined });
  if (table.tableSha256 !== expectedSha) {
    errors.push(
      validationIssue(
        `${path}.tableSha256`,
        "sha_mismatch",
        "table SHA-256이 일치하지 않습니다.",
      ),
    );
  }
  return { errors, warnings };
}

function validateQueryJsonSemanticProfile(profile = {}) {
  const errors = [];
  const warnings = [];
  if (profile.version !== QUERY_JSON_SEMANTIC_PROFILE_VERSION) {
    errors.push(
      validationIssue(
        "version",
        "invalid_version",
        "semantic profile version이 유효하지 않습니다.",
      ),
    );
  }
  if (profile.tableProfileVersion !== QUERY_JSON_TABLE_PROFILE_VERSION) {
    errors.push(
      validationIssue(
        "tableProfileVersion",
        "invalid_version",
        "table profile version이 유효하지 않습니다.",
      ),
    );
  }
  if (profile.columnProfileVersion !== QUERY_JSON_COLUMN_PROFILE_VERSION) {
    errors.push(
      validationIssue(
        "columnProfileVersion",
        "invalid_version",
        "column profile version이 유효하지 않습니다.",
      ),
    );
  }
  if (!Array.isArray(profile.tables)) {
    errors.push(
      validationIssue("tables", "invalid_type", "tables는 배열이어야 합니다."),
    );
  } else {
    profile.tables.forEach((table, index) => {
      const validation = validateTableProfile(table, index);
      errors.push(...validation.errors);
      warnings.push(...validation.warnings);
    });
  }
  if (
    Number(profile.counts?.totalTables || 0) !== asArray(profile.tables).length
  ) {
    errors.push(
      validationIssue(
        "counts.totalTables",
        "count_mismatch",
        "totalTables가 실제 표 수와 다릅니다.",
      ),
    );
  }
  const totalColumns = asArray(profile.tables).reduce(
    (sum, table) => sum + asArray(table.columns).length,
    0,
  );
  if (Number(profile.counts?.totalColumns || 0) !== totalColumns) {
    errors.push(
      validationIssue(
        "counts.totalColumns",
        "count_mismatch",
        "totalColumns가 실제 열 수와 다릅니다.",
      ),
    );
  }
  const expectedSha = sha256({ ...profile, profileSha256: undefined });
  if (profile.profileSha256 !== expectedSha) {
    errors.push(
      validationIssue(
        "profileSha256",
        "sha_mismatch",
        "profile SHA-256이 일치하지 않습니다.",
      ),
    );
  }
  if (!profile.tables?.length) {
    warnings.push(
      validationIssue("tables", "no_tables", "정규화할 표가 없습니다."),
    );
  }
  return {
    valid: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings,
  };
}

module.exports = {
  QUERY_JSON_SEMANTIC_PROFILE_VERSION,
  QUERY_JSON_TABLE_PROFILE_VERSION,
  QUERY_JSON_COLUMN_PROFILE_VERSION,
  NORMALIZER_SAMPLE_ROW_LIMIT,
  ISSUE_LEVELS,
  DATA_TYPES,
  buildQueryJsonSemanticProfile,
  validateQueryJsonSemanticProfile,
  inferDataType,
  inferSemanticRole,
  metricFamilyFor,
  operationCapabilities,
};
