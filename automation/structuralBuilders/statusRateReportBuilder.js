const {
  findColumnHeader,
  makeTemplateCandidate,
  makeTemplateSection,
  executeTemplateSections,
  getRows,
  getRowValue,
  toNumber,
} = require("../businessTemplates/commonTemplateHelpers");

const STATUS_RATE_REPORT_VERSION =
  "status_rate_report_builder_v2_percentage_contract";
const PERCENTAGE_VALUE_CONTRACT_VERSION = "percentage_value_contract_v1";

const DEFAULT_STATUS_HINTS = [
  "상태",
  "처리상태",
  "진행상태",
  "승인상태",
  "정산상태",
  "검수상태",
  "발주상태",
  "신청상태",
  "문의상태",
  "계약상태",
  "참여상태",
  "재직상태",
  "장비상태",
  "자산상태",
  "보유상태",
  "완료여부",
  "이수여부",
  "참석여부",
  "취소여부",
  "처리결과",
  "결과",
  "status",
  "state",
  "result",
];

const DEFAULT_DATE_HINTS = [
  "일자",
  "날짜",
  "월",
  "연월",
  "기준월",
  "등록일",
  "신청일",
  "접수일",
  "요청일",
  "처리일",
  "완료일",
  "승인일",
  "발주일",
  "검수일",
  "계약일",
  "시작일",
  "종료일",
  "date",
  "month",
  "period",
];

const DEFAULT_METRIC_HINTS = [
  "금액",
  "비용",
  "사용금액",
  "승인금액",
  "지출금액",
  "구매금액",
  "취득금액",
  "발주금액",
  "계약금액",
  "집행금액",
  "정산금액",
  "수량",
  "건수",
  "승인건수",
  "점수",
  "평점",
  "amount",
  "cost",
  "price",
  "count",
  "score",
];

const DEFAULT_OWNER_HINTS = [
  "담당자",
  "사용자",
  "신청자",
  "요청자",
  "처리자",
  "검수자",
  "출장자",
  "성명",
  "이름",
  "직원명",
  "owner",
  "manager",
  "user",
  "name",
];

const DEFAULT_DEPARTMENT_HINTS = [
  "부서",
  "소속",
  "조직",
  "팀",
  "부서명",
  "소속부서",
  "기관",
  "department",
  "team",
  "organization",
];

const DEFAULT_CATEGORY_HINTS = [
  "유형",
  "구분",
  "분류",
  "카테고리",
  "항목",
  "비목",
  "세목",
  "품목",
  "물품",
  "장비",
  "자산",
  "문의유형",
  "지원분야",
  "전형",
  "회차",
  "업체",
  "거래처",
  "공급사",
  "가맹점",
  "category",
  "type",
  "item",
  "vendor",
];

const COMPLETED_KEYWORDS = [
  "완료",
  "승인",
  "정상",
  "처리완료",
  "검수완료",
  "입고완료",
  "수령완료",
  "지급완료",
  "정산완료",
  "계약완료",
  "참석",
  "참여",
  "이수",
  "합격",
  "선정",
  "완료됨",
  "completed",
  "complete",
  "approved",
  "done",
  "closed",
  "success",
  "active",
];

const PENDING_KEYWORDS = [
  "대기",
  "진행",
  "접수",
  "요청",
  "예정",
  "검토",
  "확인중",
  "처리중",
  "미완료",
  "미처리",
  "미정산",
  "미검수",
  "보류",
  "pending",
  "progress",
  "open",
  "requested",
  "review",
];

const CANCELLED_KEYWORDS = [
  "취소",
  "반려",
  "거절",
  "불합격",
  "미승인",
  "불가",
  "탈락",
  "해지",
  "폐기",
  "cancel",
  "rejected",
  "declined",
  "failed",
  "inactive",
];

const DELAYED_KEYWORDS = [
  "지연",
  "연체",
  "초과",
  "미납",
  "기한초과",
  "delayed",
  "late",
  "overdue",
];

function normalizeStatusText(value = "") {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()\[\]{}]/g, "")
    .trim();
}

function includesAny(text = "", keywords = []) {
  const normalized = normalizeStatusText(text);
  return keywords.some((keyword) => {
    const target = normalizeStatusText(keyword);
    return target && normalized.includes(target);
  });
}

function classifyStatus(value = "") {
  const text = normalizeStatusText(value);
  if (!text) return "unknown";
  if (includesAny(text, DELAYED_KEYWORDS)) return "delayed";
  if (includesAny(text, CANCELLED_KEYWORDS)) return "cancelled";
  if (includesAny(text, COMPLETED_KEYWORDS)) return "completed";
  if (includesAny(text, PENDING_KEYWORDS)) return "pending";
  return "other";
}

function safeRate(numerator = 0, denominator = 0) {
  const n = Number(numerator || 0);
  const d = Number(denominator || 0);
  if (!Number.isFinite(n) || !Number.isFinite(d) || d === 0) return null;
  return n / d;
}

function makePercent(value) {
  return value == null ? null : value * 100;
}

function findStatusRateHeaders(table = {}, config = {}) {
  const hints = config.hints || {};

  const statusHeader = findColumnHeader(table, [
    ...(hints.status || []),
    ...DEFAULT_STATUS_HINTS,
  ]);

  const dateHeader = findColumnHeader(table, [
    ...(hints.date || []),
    ...DEFAULT_DATE_HINTS,
  ]);

  const metricHeader = findColumnHeader(
    table,
    [...(hints.metric || []), ...DEFAULT_METRIC_HINTS],
    { type: "number" },
  );

  const departmentHeader = findColumnHeader(table, [
    ...(hints.department || []),
    ...DEFAULT_DEPARTMENT_HINTS,
  ]);

  const ownerHeader = findColumnHeader(table, [
    ...(hints.owner || []),
    ...DEFAULT_OWNER_HINTS,
  ]);

  const categoryHeader = findColumnHeader(table, [
    ...(hints.category || []),
    ...DEFAULT_CATEGORY_HINTS,
  ]);

  return {
    statusHeader,
    dateHeader,
    metricHeader,
    departmentHeader,
    ownerHeader,
    categoryHeader,
  };
}

function makeCustomMetricSection({
  sectionId,
  sectionType,
  title,
  table,
  rows,
  columns = {},
  chartHint = {},
  narrativeHint = {},
  meta = {},
}) {
  return makeTemplateSection({
    sectionId,
    sectionType,
    title,
    candidate: {
      recipeType: "custom_metric",
      title,
      tableId: table.tableId,
      columns,
      meta: {
        ...meta,
        statusRateReportVersion: STATUS_RATE_REPORT_VERSION,
        percentageValueContractVersion: PERCENTAGE_VALUE_CONTRACT_VERSION,
      },
    },
    result: {
      ok: true,
      recipeType: "custom_metric",
      resultType: sectionType,
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
      meta: {
        ...meta,
        statusRateReportVersion: STATUS_RATE_REPORT_VERSION,
        percentageValueContractVersion: PERCENTAGE_VALUE_CONTRACT_VERSION,
      },
    },
    chartHint,
    narrativeHint,
  });
}

function buildStatusCounts(rows = [], statusHeader = "") {
  const total = rows.length;
  const statusMap = new Map();
  const classCounts = {
    completed: 0,
    pending: 0,
    cancelled: 0,
    delayed: 0,
    other: 0,
    unknown: 0,
  };

  rows.forEach((row) => {
    const rawStatus =
      String(getRowValue(row, statusHeader) ?? "").trim() || "미입력";
    const normalizedClass = classifyStatus(rawStatus);

    classCounts[normalizedClass] = (classCounts[normalizedClass] || 0) + 1;

    if (!statusMap.has(rawStatus)) {
      statusMap.set(rawStatus, {
        상태: rawStatus,
        건수: 0,
        상태그룹: normalizedClass,
      });
    }

    statusMap.get(rawStatus).건수 += 1;
  });

  const statusRows = Array.from(statusMap.values())
    .map((item) => ({
      ...item,
      전체건수: total,
      비율: safeRate(item.건수, total),
      비율Percent: makePercent(safeRate(item.건수, total)),
    }))
    .sort((a, b) => Number(b.건수 || 0) - Number(a.건수 || 0));

  return {
    total,
    classCounts,
    statusRows,
  };
}

function buildStatusOverviewSection({ table, headers, config = {} }) {
  const { statusHeader } = headers || {};
  if (!table?.tableId || !statusHeader) return null;

  const rows = getRows(table);
  const { total, classCounts } = buildStatusCounts(rows, statusHeader);
  if (!total) return null;

  const completed = classCounts.completed || 0;
  const pending = (classCounts.pending || 0) + (classCounts.delayed || 0);
  const cancelled = classCounts.cancelled || 0;
  const other = (classCounts.other || 0) + (classCounts.unknown || 0);

  const resultRows = [
    {
      지표: "전체 건수",
      값: total,
      비율: 1,
      비율Percent: 100,
    },
    {
      지표: config.labels?.completed || "완료·승인 건수",
      값: completed,
      비율: safeRate(completed, total),
      비율Percent: makePercent(safeRate(completed, total)),
    },
    {
      지표: config.labels?.pending || "진행·대기 건수",
      값: pending,
      비율: safeRate(pending, total),
      비율Percent: makePercent(safeRate(pending, total)),
    },
    {
      지표: config.labels?.cancelled || "취소·반려 건수",
      값: cancelled,
      비율: safeRate(cancelled, total),
      비율Percent: makePercent(safeRate(cancelled, total)),
    },
  ];

  if (other > 0) {
    resultRows.push({
      지표: "기타·미분류 건수",
      값: other,
      비율: safeRate(other, total),
      비율Percent: makePercent(safeRate(other, total)),
    });
  }

  return makeCustomMetricSection({
    sectionId: config.sectionIds?.overview || "status_rate_overview",
    sectionType: "status_rate_overview",
    title: config.titles?.overview || "상태 처리율 요약",
    table,
    rows: resultRows,
    columns: {
      status: statusHeader,
      total: "전체 건수",
      completed: "완료·승인 건수",
      pending: "진행·대기 건수",
      cancelled: "취소·반려 건수",
    },
    chartHint: {
      preferredType: "metric_card",
      valueField: "값",
      ratioField: "비율Percent",
    },
    narrativeHint: {
      focus: "status_rate_overview",
      status: statusHeader,
    },
  });
}

function buildStatusRatioSection({ table, headers, config = {} }) {
  const { statusHeader } = headers || {};
  if (!table?.tableId || !statusHeader) return null;

  const { statusRows } = buildStatusCounts(getRows(table), statusHeader);
  if (!statusRows.length) return null;

  return makeCustomMetricSection({
    sectionId: config.sectionIds?.statusRatio || "status_ratio_breakdown",
    sectionType: "status_ratio_breakdown",
    title: config.titles?.statusRatio || `${statusHeader}별 구성비`,
    table,
    rows: statusRows,
    columns: {
      status: statusHeader,
      count: "건수",
      ratio: "비율Percent",
    },
    chartHint: {
      preferredType: "donut",
      categoryField: "상태",
      valueField: "건수",
      ratioField: "비율Percent",
    },
    narrativeHint: {
      focus: "status_ratio",
      status: statusHeader,
    },
  });
}

function buildDimensionStatusRateSection({
  table,
  headers,
  dimensionHeader = "",
  title = "",
  sectionId = "",
}) {
  const { statusHeader } = headers || {};
  if (!table?.tableId || !statusHeader || !dimensionHeader) return null;

  const map = new Map();

  getRows(table).forEach((row) => {
    const dimension =
      String(getRowValue(row, dimensionHeader) ?? "").trim() || "미입력";
    const statusClass = classifyStatus(getRowValue(row, statusHeader));

    if (!map.has(dimension)) {
      map.set(dimension, {
        [dimensionHeader]: dimension,
        전체건수: 0,
        완료승인건수: 0,
        진행대기건수: 0,
        취소반려건수: 0,
        지연건수: 0,
      });
    }

    const item = map.get(dimension);
    item.전체건수 += 1;

    if (statusClass === "completed") item.완료승인건수 += 1;
    else if (statusClass === "pending") item.진행대기건수 += 1;
    else if (statusClass === "delayed") item.지연건수 += 1;
    else if (statusClass === "cancelled") item.취소반려건수 += 1;
  });

  const resultRows = Array.from(map.values())
    .map((item) => {
      const completionRate = safeRate(item.완료승인건수, item.전체건수);
      const incompleteRate = safeRate(
        item.진행대기건수 + item.지연건수,
        item.전체건수,
      );
      const cancelledRate = safeRate(item.취소반려건수, item.전체건수);

      return {
        ...item,
        완료율: completionRate,
        완료율Percent: makePercent(completionRate),
        미완료율: incompleteRate,
        미완료율Percent: makePercent(incompleteRate),
        취소반려율: cancelledRate,
        취소반려율Percent: makePercent(cancelledRate),
      };
    })
    .sort((a, b) => Number(b.전체건수 || 0) - Number(a.전체건수 || 0));

  if (!resultRows.length) return null;

  return makeCustomMetricSection({
    sectionId: sectionId || `status_rate_by_${dimensionHeader}`,
    sectionType: "status_rate_by_dimension",
    title: title || `${dimensionHeader}별 상태 처리율`,
    table,
    rows: resultRows,
    columns: {
      dimension: dimensionHeader,
      status: statusHeader,
      total: "전체건수",
      completed: "완료승인건수",
      completionRate: "완료율Percent",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "완료율Percent",
    },
    narrativeHint: {
      focus: "status_rate_by_dimension",
      dimension: dimensionHeader,
      status: statusHeader,
    },
  });
}

function buildStatusRateCandidates({ table, headers, config = {} }) {
  if (!table?.tableId) return [];

  const {
    statusHeader,
    dateHeader,
    metricHeader,
    departmentHeader,
    ownerHeader,
    categoryHeader,
  } = headers || {};

  const candidates = [];
  const tableId = table.tableId;

  if (statusHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.statusCount || "status_count",
        sectionType: "status_count",
        recipeType: "category_count",
        title: config.titles?.statusCount || `${statusHeader}별 건수`,
        tableId,
        columns: {
          dimension: statusHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: statusHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "status_count",
          status: statusHeader,
        },
      }),
    );

    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.statusComposition || "status_composition",
        sectionType: "status_composition",
        recipeType: "composition_ratio",
        title: config.titles?.statusComposition || `${statusHeader} 구성비`,
        tableId,
        columns: {
          dimension: statusHeader,
        },
        chartHint: {
          preferredType: "donut",
          categoryField: statusHeader,
          valueField: "value",
        },
        narrativeHint: {
          focus: "status_composition",
          status: statusHeader,
        },
      }),
    );
  }

  if (dateHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.timeCount || "status_time_count",
        sectionType: "status_time_count",
        recipeType: "time_count",
        title: config.titles?.timeCount || `${dateHeader}별 건수 추이`,
        tableId,
        columns: {
          date: dateHeader,
        },
        chartHint: {
          preferredType: "line",
          categoryField: dateHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "time_count",
          date: dateHeader,
        },
      }),
    );
  }

  const summaryDimensions = [
    departmentHeader,
    categoryHeader,
    ownerHeader,
  ].filter((value, index, arr) => value && arr.indexOf(value) === index);

  for (const dimensionHeader of summaryDimensions) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: `count_by_${dimensionHeader}`,
        sectionType: "dimension_count",
        recipeType: "category_count",
        title: `${dimensionHeader}별 건수`,
        tableId,
        columns: {
          dimension: dimensionHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: dimensionHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "dimension_count",
          dimension: dimensionHeader,
        },
      }),
    );

    if (metricHeader) {
      candidates.push(
        makeTemplateCandidate({
          sectionId: `metric_by_${dimensionHeader}`,
          sectionType: "dimension_metric_summary",
          recipeType: "group_summary",
          title: `${dimensionHeader}별 ${metricHeader} 요약`,
          tableId,
          columns: {
            dimension: dimensionHeader,
            metric: metricHeader,
          },
          chartHint: {
            preferredType: "bar",
            categoryField: dimensionHeader,
            valueField: metricHeader,
          },
          narrativeHint: {
            focus: "dimension_metric_summary",
            dimension: dimensionHeader,
            metric: metricHeader,
          },
        }),
      );
    }
  }

  if (metricHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: "status_metric_top_bottom",
        sectionType: "status_metric_top_bottom",
        recipeType: "top_bottom",
        title: `${metricHeader} 상위·하위 항목`,
        tableId,
        columns: {
          dimension:
            categoryHeader || ownerHeader || departmentHeader || statusHeader,
          metric: metricHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField:
            categoryHeader || ownerHeader || departmentHeader || statusHeader,
          valueField: metricHeader,
        },
        narrativeHint: {
          focus: "top_bottom",
          metric: metricHeader,
        },
      }),
    );
  }

  return candidates.filter((candidate) => {
    if (!candidate.columns) return true;
    return Object.values(candidate.columns).every(Boolean);
  });
}

function buildStatusRateReportSections({
  normalizedQueryTables = [],
  table,
  templateCandidate = {},
  config = {},
}) {
  if (!table?.tableId) return [];

  const headers = findStatusRateHeaders(table, config);

  if (!headers.statusHeader) {
    const fallbackCandidates = Array.isArray(templateCandidate.candidates)
      ? templateCandidate.candidates
      : [];

    if (!fallbackCandidates.length) return [];

    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const customSections = [
    buildStatusOverviewSection({ table, headers, config }),
    buildStatusRatioSection({ table, headers, config }),
    buildDimensionStatusRateSection({
      table,
      headers,
      dimensionHeader: headers.departmentHeader,
      sectionId: "status_rate_by_department",
      title: headers.departmentHeader
        ? `${headers.departmentHeader}별 상태 처리율`
        : "부서별 상태 처리율",
    }),
    buildDimensionStatusRateSection({
      table,
      headers,
      dimensionHeader: headers.categoryHeader,
      sectionId: "status_rate_by_category",
      title: headers.categoryHeader
        ? `${headers.categoryHeader}별 상태 처리율`
        : "유형별 상태 처리율",
    }),
  ].filter(Boolean);

  const candidates = buildStatusRateCandidates({ table, headers, config });
  const recipeSections = executeTemplateSections({
    normalizedQueryTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });

  return [...customSections, ...recipeSections];
}

module.exports = {
  STATUS_RATE_REPORT_VERSION,
  findStatusRateHeaders,
  classifyStatus,
  buildStatusRateCandidates,
  buildStatusRateReportSections,
};
