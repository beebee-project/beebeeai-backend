const NARRATIVE_VERSION = "report_narrative_v3";

const DOMAIN_DEFS = Object.freeze({
  sales: Object.freeze({
    label: "매출·영업",
    reportName: "매출 분석 보고서",
    focus: "기간별 증감, 상위 항목, 구성비를 중심으로 해석했습니다.",
    check:
      "매출 변동이 큰 항목과 반복적으로 상위에 나타나는 항목을 우선 확인하세요.",
  }),
  budget: Object.freeze({
    label: "예산·지출·정산",
    reportName: "예산·지출 분석 보고서",
    focus:
      "집행 규모, 비목별 구성, 잔액 또는 초과 가능성을 중심으로 해석했습니다.",
    check:
      "집행 비중이 큰 항목과 증빙 확인이 필요한 지출 항목을 우선 점검하세요.",
  }),
  hr: Object.freeze({
    label: "인사·조직",
    reportName: "인사 현황 분석 보고서",
    focus:
      "부서·직급 구성, 인원 변동, 주요 그룹별 차이를 중심으로 해석했습니다.",
    check: "특정 부서 또는 직급에 인원이 과도하게 집중되는지 확인하세요.",
  }),
  inventory: Object.freeze({
    label: "재고·물류·자산",
    reportName: "재고·자산 분석 보고서",
    focus:
      "입출고 흐름, 보유 현황, 상위 품목과 이동 패턴을 중심으로 해석했습니다.",
    check: "재고 과다·부족 가능성이 있는 품목과 이동량이 큰 창고를 확인하세요.",
  }),
  survey: Object.freeze({
    label: "설문·평가",
    reportName: "설문·평가 분석 보고서",
    focus: "평균 점수, 낮은 점수 항목, 문항별 편차를 중심으로 해석했습니다.",
    check:
      "평균보다 낮은 문항과 응답 분포가 크게 갈리는 문항을 우선 확인하세요.",
  }),
  operation: Object.freeze({
    label: "운영·상태",
    reportName: "운영 현황 분석 보고서",
    focus: "처리 상태, 완료율, 미처리·지연 항목을 중심으로 해석했습니다.",
    check: "미완료 또는 지연 상태의 항목을 우선 조치 대상으로 분류하세요.",
  }),
  project: Object.freeze({
    label: "프로젝트·성과",
    reportName: "성과 지표 분석 보고서",
    focus: "목표 대비 실적, 달성률, 성과 편차를 중심으로 해석했습니다.",
    check: "목표 대비 미달 항목과 가중치가 큰 핵심 지표를 우선 확인하세요.",
  }),
  energy: Object.freeze({
    label: "에너지·비용",
    reportName: "에너지 사용량 분석 보고서",
    focus: "사용량, 절감량, 비용, 회수기간을 중심으로 해석했습니다.",
    check: "사용량 또는 비용이 큰 항목과 절감 효과가 낮은 항목을 확인하세요.",
  }),
  general: Object.freeze({
    label: "일반",
    reportName: "분석 보고서",
    focus: "주요 그룹, 수치 규모, 상위·하위 항목을 중심으로 해석했습니다.",
    check: "수치가 큰 항목과 변동 폭이 큰 항목을 우선 확인하세요.",
  }),
});

const OPERATION_LABELS = Object.freeze({
  sum: "합계",
  groupSum: "그룹별 합계",
  group_sum: "그룹별 합계",
  average: "평균",
  avg: "평균",
  groupAvg: "그룹별 평균",
  group_avg: "그룹별 평균",
  count: "건수",
  groupCount: "그룹별 건수",
  group_count: "그룹별 건수",
  compositionRatio: "구성비",
  composition_ratio: "구성비",
  topBottom: "상위·하위",
  top_bottom: "상위·하위",
  timeSum: "기간별 합계",
  time_sum: "기간별 합계",
  timeAverage: "기간별 평균",
  time_average: "기간별 평균",
  multiAggregate: "복수 지표 요약",
  pipelineCombine: "복합 분석",
  pivot: "교차 분석",
  grouped: "그룹 분석",
  businessTemplate: "업무 템플릿 분석",
});

function toNumber(v) {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (v == null || v === "") return null;
  const n = Number(String(v).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

function text(value = "") {
  return String(value ?? "");
}

function compactText(value = "") {
  return text(value)
    .replace(/undefined|null|NaN|\[object Object\]/gi, "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function uniqueItems(items = []) {
  const seen = new Set();
  const result = [];
  for (const item of items.map(compactText).filter(Boolean)) {
    const key = item.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(item);
  }
  return result;
}

function inferLabelKey(row = {}) {
  const keys = Object.keys(row || {});

  return (
    keys.find((key) =>
      /name|이름|성명|title|제목|label|항목|부서|팀|구분|분류|연월|월|일자|date|id$/i.test(
        key,
      ),
    ) ||
    keys.find((key) => typeof row[key] === "string" && row[key].trim()) ||
    keys[0] ||
    "항목"
  );
}

function formatValue(v, field = "") {
  const n = toNumber(v);
  if (n == null) return compactText(v);

  const source = text(field);
  const digits = /율|비율|rate|ratio|증감|달성/.test(source)
    ? 1
    : Number.isInteger(n)
      ? 0
      : 2;
  return Number(n.toFixed(digits)).toLocaleString("ko-KR");
}

function numericKeys(row = {}, exclude = []) {
  const excluded = new Set(exclude);

  return Object.keys(row || {}).filter((key) => {
    if (excluded.has(key)) return false;
    const n = toNumber(row[key]);
    return n != null;
  });
}

function getRows(result = {}) {
  return Array.isArray(result.rows) ? result.rows : [];
}

function getGroupHeader(result = {}) {
  if (result.groupBy?.header) return result.groupBy.header;
  if (result.pivot?.rowGroup?.header) return result.pivot.rowGroup.header;

  return inferLabelKey(result.rows?.[0] || {});
}

function getMetricHeader(result = {}) {
  if (result.metric?.header) return result.metric.header;
  const rows = getRows(result);
  const first = rows[0] || {};
  return numericKeys(first, [getGroupHeader(result), "rowCount"])[0] || "값";
}

function getPrimaryNumericKey(result = {}) {
  const rows = getRows(result);
  if (!rows.length) return null;

  const groupHeader = getGroupHeader(result);

  if (
    result.resultType === "grouped" &&
    result.operation !== "multiAggregate" &&
    result.operation !== "pipelineCombine" &&
    Object.prototype.hasOwnProperty.call(rows[0], "value")
  ) {
    return "value";
  }

  return (
    numericKeys(rows[0], [groupHeader, "rowCount", "기준값", "비교값"])[0] ||
    null
  );
}

function normalizeFieldLabel(field = "", context = {}) {
  const key = compactText(field);
  const source = compactText(
    [key, context.title, context.message, context.fileName]
      .filter(Boolean)
      .join(" "),
  );
  const lower = key.toLowerCase();

  if (!key) return "값";
  if (lower === "value") {
    if (/연봉|급여/.test(source)) return "연봉";
    if (/매출|비용|손익|예산|금액|집행/.test(source)) return "금액";
    if (/달성|성과|kpi/i.test(source)) return "실적값";
    if (/에너지|사용량/.test(source)) return "사용량";
    return "값";
  }
  if (lower === "count" || lower === "cnt")
    return /인원|직원|명단/.test(source) ? "인원 수" : "건수";
  if (lower === "sum" || lower === "total") return "합계";
  if (lower === "average" || lower === "avg" || lower === "mean") return "평균";
  if (lower === "label") return "항목";
  if (lower === "rowcount") return "행 수";
  return key.replace(/_/g, " ").replace(/\s+/g, " ").trim();
}

function operationLabel(result = {}) {
  const key = compactText(result.operation || result.resultType || "");
  return OPERATION_LABELS[key] || key || "분석";
}

function inferDomain(result = {}, context = {}) {
  const source = compactText(
    [
      context.domain,
      context.templateId,
      context.message,
      context.fileName,
      result.templateId,
      result.title,
      result.operation,
      getGroupHeader(result),
      getMetricHeader(result),
    ]
      .filter(Boolean)
      .join(" "),
  ).toLowerCase();

  if (/energy|에너지|전력|전기|사용량|절감|회수기간/.test(source))
    return "energy";
  if (/kpi|성과|달성률|목표|실적|프로젝트|project/.test(source))
    return "project";
  if (/매출|영업|판매|revenue|sales|손익|비용|profit|cost/.test(source))
    return "sales";
  if (/예산|집행|정산|카드|출장|회의비|budget|expense/.test(source))
    return "budget";
  if (/인사|직원|부서|직급|입사|퇴사|연봉|hr|employee/.test(source))
    return "hr";
  if (/재고|입고|출고|창고|자산|장비|inventory|asset|warehouse/.test(source))
    return "inventory";
  if (/만족도|설문|문항|평점|점수|survey|feedback/.test(source))
    return "survey";
  if (/상태|완료|처리|이수|점검|배송|운영|status|operation/.test(source))
    return "operation";
  return "general";
}

function domainDef(domain = "general") {
  return DOMAIN_DEFS[domain] || DOMAIN_DEFS.general;
}

function topBottomHighlight(
  rows = [],
  labelKey = "",
  valueKey = "",
  valueLabel = valueKey,
) {
  const values = rows
    .map((r) => ({
      label: compactText(r[labelKey]),
      value: toNumber(r[valueKey]),
    }))
    .filter((r) => r.value != null && r.label !== "");

  const label = normalizeFieldLabel(valueLabel || valueKey || "값");

  if (!values.length) return [];

  const max = values.reduce((a, b) => (b.value > a.value ? b : a));
  const min = values.reduce((a, b) => (b.value < a.value ? b : a));

  if (max.label === min.label) {
    return [
      `${max.label}의 ${label}은 ${formatValue(max.value, valueKey)}입니다.`,
    ];
  }

  return [
    `${label}이 가장 큰 항목은 ${max.label}(${formatValue(max.value, valueKey)})입니다.`,
    `${label}이 가장 작은 항목은 ${min.label}(${formatValue(min.value, valueKey)})입니다.`,
  ];
}

function growthHighlight(result = {}) {
  const rows = getRows(result);
  const groupHeader = getGroupHeader(result);

  const growthRows = rows
    .map((r) => ({
      label: compactText(r[groupHeader]),
      value: toNumber(r["증감률"]),
    }))
    .filter((r) => r.value != null);

  if (!growthRows.length) return [];

  const max = growthRows.reduce((a, b) => (b.value > a.value ? b : a));
  const min = growthRows.reduce((a, b) => (b.value < a.value ? b : a));

  return [
    `증감률은 ${max.label}에서 가장 높았고(${formatValue(max.value, "증감률")}%), ${min.label}에서 가장 낮았습니다(${formatValue(min.value, "증감률")}%).`,
  ];
}

function windowHighlight(result = {}) {
  const rows = getRows(result);
  if (!rows.length) return [];

  const groupHeader = getGroupHeader(result);
  const first = rows[0] || {};
  const windowKey = Object.keys(first).find((key) =>
    /누적|이동평균|rolling|cumulative/i.test(key),
  );

  if (!windowKey) return [];

  const last = rows[rows.length - 1];

  return [
    `${windowKey}의 마지막 값은 ${compactText(last?.[groupHeader]) || "마지막 구간"} 기준 ${formatValue(last?.[windowKey], windowKey)}입니다.`,
  ];
}

function pivotHighlight(result = {}) {
  const rows = getRows(result);
  const columns = result.pivot?.columns || [];

  if (!rows.length || !columns.length) return [];

  const filledCount = rows.reduce(
    (sum, row) =>
      sum +
      columns.filter((col) => {
        const v = row[col];
        return v !== "" && v != null;
      }).length,
    0,
  );

  const totalCells = rows.length * columns.length;
  const density = totalCells ? (filledCount / totalCells) * 100 : 0;

  return [
    `교차표는 ${rows.length.toLocaleString()}개 행과 ${columns.length.toLocaleString()}개 열 항목으로 구성되었습니다.`,
    `값이 채워진 셀 비율은 ${formatValue(density, "비율")}%입니다.`,
  ];
}

function buildTitle(result = {}, context = {}) {
  const message = compactText(context.message || "");
  if (message) return `${message} 분석 보고서`;

  const domain = inferDomain(result, context);
  const def = domainDef(domain);
  const groupHeader = normalizeFieldLabel(getGroupHeader(result), context);
  const metricHeader = normalizeFieldLabel(getMetricHeader(result), context);

  if (result.resultType === "pivot") {
    return `${groupHeader} 기준 교차 분석 보고서`;
  }

  if (domain !== "general") return def.reportName;
  return `${groupHeader}별 ${metricHeader} 분석 보고서`;
}

function buildSummary(result = {}, context = {}) {
  const rows = getRows(result);
  const rowCount = rows.length;
  const opLabel = operationLabel(result);
  const domain = inferDomain(result, context);
  const def = domainDef(domain);
  const groupHeader = normalizeFieldLabel(getGroupHeader(result), context);
  const metricHeader = normalizeFieldLabel(getMetricHeader(result), context);

  const scale = rowCount
    ? `${rowCount.toLocaleString()}건의 결과 행`
    : "결과 요약";

  return `${def.label} 관점에서 ${groupHeader} 기준 ${metricHeader} ${opLabel}을 ${scale}으로 정리했습니다. ${def.focus}`;
}

function buildHighlights(result = {}, context = {}) {
  const rows = getRows(result);
  const domain = inferDomain(result, context);
  const def = domainDef(domain);
  const groupHeader = getGroupHeader(result);
  const groupLabel = normalizeFieldLabel(groupHeader, context);

  if (!rows.length)
    return [
      `${def.label} 데이터의 구조와 주요 지표를 확인했습니다.`,
      def.check,
    ];

  if (result.resultType === "pivot") {
    return uniqueItems([...pivotHighlight(result), def.check]).slice(0, 5);
  }

  const highlights = [];

  if (
    result.operation === "multiAggregate" ||
    result.operation === "pipelineCombine"
  ) {
    const metricKeys = numericKeys(rows[0], [groupHeader, "rowCount"]);

    highlights.push(
      `${groupLabel} 기준으로 ${rows.length.toLocaleString()}개 그룹을 비교했습니다.`,
    );

    highlights.push(
      `${metricKeys.length.toLocaleString()}개 지표를 함께 계산해 규모와 편차를 확인했습니다.`,
    );

    const primaryKey = metricKeys[0];
    if (primaryKey) {
      highlights.push(
        ...topBottomHighlight(rows, groupHeader, primaryKey, primaryKey),
      );
    }

    highlights.push(def.check);
    return uniqueItems(highlights).slice(0, 5);
  }

  highlights.push(
    `${groupLabel} 기준으로 ${rows.length.toLocaleString()}개 결과를 비교했습니다.`,
  );

  highlights.push(...growthHighlight(result));
  highlights.push(...windowHighlight(result));

  const primaryKey = getPrimaryNumericKey(result);
  if (primaryKey) {
    const valueLabel =
      primaryKey === "value" ? result.metric?.header || "값" : primaryKey;
    highlights.push(
      ...topBottomHighlight(rows, groupHeader, primaryKey, valueLabel),
    );
  }

  highlights.push(def.check);
  return uniqueItems(highlights).slice(0, 5);
}

function buildNarrativeSections(result = {}, context = {}) {
  const domain = inferDomain(result, context);
  const def = domainDef(domain);
  return {
    version: NARRATIVE_VERSION,
    domain,
    domainLabel: def.label,
    title: buildTitle(result, context),
    summary: buildSummary(result, context),
    highlights: buildHighlights(result, context),
    resultType: result.resultType || "",
    operation: result.operation || "",
    qualitySignals: {
      domainAware: true,
      sanitizedNarrative: true,
      duplicateHighlightsRemoved: true,
    },
  };
}

module.exports = {
  buildNarrativeSections,
};
